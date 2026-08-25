#!/usr/bin/env python3
"""主版 invoice_gen.py 同步到 flex, 保留 flex 独有差异:
1. DEFAULT_FIXED 项目/编码/单位/税率 为空
2. validate_invoice 有灵活版必填校验
"""
import os, ast

ROOT = "/mnt/ssd/table-match-tool"
main = open(os.path.join(ROOT, "invoice_gen.py"), encoding="utf-8").read()
flex_path = os.path.join(ROOT, "flex", "invoice_gen.py")

new = main

# 差异1: DEFAULT_FIXED 空值
import re
old_block = re.search(r"# ---------- 默认固定内容.*?\nDEFAULT_FIXED = \{[^}]+\}", new, re.S)
if old_block:
    new = new.replace(old_block.group(0),
        '# ---------- 默认固定内容 (灵活版: 项目/编码/单位/税率不固定, 留空手填) ----------\n'
        'DEFAULT_FIXED = {\n'
        '    "invoice_type": "普通发票",      # 发票类型\n'
        '    "tax_included": "是",            # 是否含税\n'
        '    "item_name": "",                 # 项目名称 (灵活版: 每次手填)\n'
        '    "tax_code": "",                  # 商品和服务税收编码 (灵活版: 每次手填)\n'
        '    "unit": "",                      # 单位 (灵活版: 每次手填)\n'
        '    "tax_rate": "",                  # 税率 (灵活版: 每次手填)\n'
        '}')
    print("差异1: DEFAULT_FIXED 空值已应用")

# 差异2: validate 必填校验(在 数量不是数字 校验后、return 前)
old_val = '''            errs.append(f"第{idx}行: 数量不是数字({qty})")
    return errs'''
new_val = '''            errs.append(f"第{idx}行: 数量不是数字({qty})")
    # 灵活版: 项目名称/税收编码/单位/税率 必填(不固定, 生成时校验)
    for field, label in (("item_name", "项目名称"), ("tax_code", "税收编码"),
                         ("unit", "单位"), ("tax_rate", "税率")):
        v = str(inv.get(field, "") or "").strip()
        if not v:
            errs.append(f"第{idx}行: {label}为空(请在上方固定内容栏填写)")
    return errs'''
if old_val in new:
    new = new.replace(old_val, new_val, 1)
    print("差异2: 灵活版必填校验已应用")
else:
    print("⚠️ 校验差异未找到")

with open(flex_path, "w", encoding="utf-8") as f:
    f.write(new)
ast.parse(new)
print("✅ flex/invoice_gen.py 同步完成(保留flex差异)")