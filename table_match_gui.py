#!/usr/bin/env python3
"""
通用表格核对工具 v4.0 - GUI版
================================
用「源表」某字段去「目标表」检索，回填字段 + 多规格标记

GUI 功能:
  1. 文件选择(源表/目标表)
  2. 检索键列 + 回填映射配置(表格形式)
  3. 多规格判断列 + 备注列
  4. "已有数据不覆盖"开关
  5. 一键运行 + 结果预览 + 未匹配清单

纯 tkinter 标准库实现, 绿色版/Windows 原生兼容
"""
import os
import sys
import subprocess
from threading import Thread
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from collections import defaultdict

import openpyxl


# ============ 核心逻辑 (与 CLI 版一致) ============
def find_header_row(ws, keywords):
    for r in range(1, min(6, ws.max_row + 1)):
        strs = [str(ws.cell(row=r, column=c + 1).value or '').strip()
                for c in range(ws.max_column)]
        if sum(1 for k in keywords if any(k.lower() in s.lower() for s in strs)) >= len(keywords):
            return r
    return None


def col_index_by_header(ws, header_row, header_name):
    for c in range(ws.max_column):
        if str(ws.cell(row=header_row, column=c + 1).value or '').strip() == header_name:
            return c
    return None


def run_match(src_path, tgt_path, src_key, tgt_key, fill_map, sku_col, remark_col,
              skip_existing, progress_cb=None):
    """执行核对, 返回 (matched, multi, notfound, src_ws)"""
    # 加载目标表 (收集全部行, 按键索引)
    twb = openpyxl.load_workbook(tgt_path, data_only=True)
    tws = twb[twb.sheetnames[0]]
    thr = find_header_row(tws, [tgt_key])
    if thr is None:
        raise ValueError(f"目标表找不到表头(含 '{tgt_key}')")
    key_col = col_index_by_header(tws, thr, tgt_key)
    if key_col is None:
        raise ValueError(f"目标表没有列 '{tgt_key}'")
    tgt_cols = {}
    for sc, tc in fill_map:
        ci = col_index_by_header(tws, thr, tc)
        if ci is not None:
            tgt_cols.setdefault(sc, []).append(ci)
    sku_idx = col_index_by_header(tws, thr, sku_col) if sku_col else None

    idx = defaultdict(list)
    for r in tws.iter_rows(min_row=thr + 1, values_only=True):
        k = r[key_col]
        if not k:
            continue
        for part in str(k).replace('，', ',').replace(' ', '').split(','):
            if part.strip():
                idx[part.strip()].append(r)

    # 源表
    swb = openpyxl.load_workbook(src_path)
    sws = swb[swb.sheetnames[0]]
    shr = find_header_row(sws, [src_key])
    if shr is None:
        raise ValueError(f"源表找不到表头(含 '{src_key}')")
    src_key_col = col_index_by_header(sws, shr, src_key)
    src_fill_cols = {}
    for sc, tc in fill_map:
        ci = col_index_by_header(sws, shr, sc)
        if ci is not None:
            src_fill_cols[sc] = ci
    remark_col_idx = col_index_by_header(sws, shr, remark_col) if remark_col else None

    matched = 0
    multi = 0
    notfound = []
    total = 0
    for row in sws.iter_rows(min_row=shr + 1):
        k = row[src_key_col].value
        if not k:
            continue
        total += 1
        k = str(k).strip()
        hits = idx.get(k)
        if hits:
            matched += 1
            # 增量填充
            for sc, tc in fill_map:
                sci = src_fill_cols.get(sc)
                tci_list = tgt_cols.get(sc)
                if sci is not None and tci_list and not row[sci].value:
                    row[sci].value = hits[0][tci_list[0]]
            # 多规格
            if remark_col_idx is not None and not row[remark_col_idx].value:
                if sku_idx is not None and len(hits) >= 2:
                    uniq = {str(h[sku_idx] or '').strip() for h in hits if h[sku_idx]}
                    if len(uniq) >= 2:
                        row[remark_col_idx].value = "多规格"
                        multi += 1
        else:
            notfound.append(k)
        if progress_cb and total % 10 == 0:
            progress_cb(total)

    swb.save(src_path)  # 直接保存源表(或副本)
    return matched, multi, notfound


# ============ GUI ============
class App:
    def __init__(self, root):
        self.root = root
        root.title("通用表格核对工具 v4.0")
        root.geometry("760x680")
        root.minsize(680, 600)

        # --- 文件选择 ---
        frm = ttk.LabelFrame(root, text="1. 选择文件")
        frm.pack(fill="x", padx=10, pady=5)
        self.src_var = tk.StringVar()
        self.tgt_var = tk.StringVar()
        ttk.Label(frm, text="源表(被填的):").grid(row=0, column=0, sticky="w", padx=5, pady=3)
        ttk.Entry(frm, textvariable=self.src_var, width=50).grid(row=0, column=1, padx=5)
        ttk.Button(frm, text="浏览", command=lambda: self.pick(self.src_var)).grid(row=0, column=2)
        ttk.Label(frm, text="目标表(数据源):").grid(row=1, column=0, sticky="w", padx=5)
        ttk.Entry(frm, textvariable=self.tgt_var, width=50).grid(row=1, column=1, padx=5)
        ttk.Button(frm, text="浏览", command=lambda: self.pick(self.tgt_var)).grid(row=1, column=2)

        # --- 键列 ---
        frm2 = ttk.LabelFrame(root, text="2. 检索键列 (表头名)")
        frm2.pack(fill="x", padx=10, pady=5)
        self.sk = tk.StringVar(value="快递单号")
        self.tk = tk.StringVar(value="快递单号")
        ttk.Label(frm2, text="源表键列:").grid(row=0, column=0, padx=5)
        ttk.Entry(frm2, textvariable=self.sk, width=22).grid(row=0, column=1, padx=5)
        ttk.Label(frm2, text="目标表键列:").grid(row=0, column=2, padx=5)
        ttk.Entry(frm2, textvariable=self.tk, width=22).grid(row=0, column=3, padx=5)

        # --- 回填映射 ---
        frm3 = ttk.LabelFrame(root, text="3. 回填映射 (源列 → 目标列)")
        frm3.pack(fill="x", padx=10, pady=5)
        self.map_rows = []
        self._add_row(frm3)
        ttk.Button(frm3, text="+ 添加映射", command=lambda: self._add_row(frm3)).grid(row=100, column=0, pady=3)

        # --- 多规格 ---
        frm4 = ttk.LabelFrame(root, text="4. 多规格标记 (可空)")
        frm4.pack(fill="x", padx=10, pady=5)
        self.sku_col = tk.StringVar()
        self.rmk_col = tk.StringVar()
        ttk.Label(frm4, text="多规格判断列:").grid(row=0, column=0, padx=5)
        ttk.Entry(frm4, textvariable=self.sku_col, width=20).grid(row=0, column=1, padx=5)
        ttk.Label(frm4, text="备注列(写'多规格'):").grid(row=0, column=2, padx=5)
        ttk.Entry(frm4, textvariable=self.rmk_col, width=20).grid(row=0, column=3, padx=5)

        # --- 选项 + 运行 ---
        frm5 = ttk.Frame(root)
        frm5.pack(fill="x", padx=10, pady=5)
        self.skip_existing = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm5, text="已有数据不覆盖", variable=self.skip_existing).pack(side="left", padx=5)
        self.run_btn = ttk.Button(frm5, text="▶ 开始核对", command=self.run)
        self.run_btn.pack(side="right", padx=5)

        # --- 日志/预览 ---
        frm6 = ttk.LabelFrame(root, text="运行日志 / 未匹配清单")
        frm6.pack(fill="both", expand=True, padx=10, pady=5)
        self.log = scrolledtext.ScrolledText(frm6, height=12, font=("Consolas", 9))
        self.log.pack(fill="both", expand=True, padx=5, pady=5)

    def pick(self, var):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有", "*.*")])
        if p:
            var.set(p)

    def _add_row(self, parent):
        row = tk.Frame(parent)
        idx = len(self.map_rows)
        s = tk.StringVar(value="规格商家编码" if idx == 0 else "")
        t = tk.StringVar(value="规格商家编码" if idx == 0 else "")
        ttk.Label(row, text=f"映射{idx + 1}:").pack(side="left")
        ttk.Entry(row, textvariable=s, width=18).pack(side="left", padx=3)
        ttk.Label(row, text="→").pack(side="left")
        ttk.Entry(row, textvariable=t, width=18).pack(side="left", padx=3)
        row.pack(fill="x", padx=5, pady=2)
        self.map_rows.append((s, t))

    def log_write(self, txt):
        self.log.insert("end", txt + "\n")
        self.log.see("end")

    def run(self):
        src = self.src_var.get().strip().strip('"')
        tgt = self.tgt_var.get().strip().strip('"')
        if not src or not tgt:
            messagebox.showwarning("提示", "请先选择源表和目标表")
            return
        fill_map = [(s.get().strip(), t.get().strip())
                    for s, t in self.map_rows if s.get().strip() and t.get().strip()]
        if not fill_map:
            messagebox.showwarning("提示", "请至少填写一个回填映射")
            return
        self.run_btn.config(state="disabled")
        self.log_write("▶ 开始核对...")
        try:
            matched, multi, notfound = run_match(
                src, tgt, self.sk.get().strip(), self.tk.get().strip(),
                fill_map, self.sku_col.get().strip() or None,
                self.rmk_col.get().strip() or None, self.skip_existing.get())
            self.log_write(f"✅ 匹配: {matched}, 多规格标记: {multi}, 未匹配: {len(notfound)}")
            self.log_write(f"   结果已保存到源表 (同文件)")
            if notfound:
                self.log_write(f"\n⚠️ 未匹配清单 ({len(notfound)}):")
                for v in notfound[:50]:
                    self.log_write(f"   {v}")
                if len(notfound) > 50:
                    self.log_write(f"   ... 等 {len(notfound)} 个")
            messagebox.showinfo("成功", f"核对完成!\n匹配 {matched}\n多规格 {multi}\n未匹配 {len(notfound)}")
        except Exception as e:
            self.log_write(f"❌ 错误: {e}")
            messagebox.showerror("错误", str(e))
        finally:
            self.run_btn.config(state="normal")


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        style.theme_use("clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
