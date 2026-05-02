---
name: pdf-to-excel
description: 将 PDF 订单文件中的数据提取并汇总到 Excel 表格。触发场景：用户要求从 PDF 提取数据写入 Excel、PDF 订单汇总、PDF 表格数据导出、按日期填写 Excel 数据。支持跨月表格和多种文件名格式。
---

# PDF 订单数据汇总到 Excel

## 功能说明

将 PDF 订单文件中的表格数据提取出来，按日期汇总到 Excel 表格的对应列中。

## 支持的文件名格式

| 格式 | 示例 | 说明 |
|------|------|------|
| `{日期}号单{序号}.pdf` | `10号单1.pdf` | 传统的"X号单"格式 |
| `{日期}-{名称}.pdf` | `1-单1-发货单.pdf` | 日期为开头数字，解析第一个 `-` 前的数字 |

## 数据提取规则

从 PDF 表格中提取三列数据：
- **平台SKC/商家货号**: 取第二行作为商品名
- **属性集**: 用正则提取尺码（XS/S/M/L/XL）
- **数量**: 直接取值

## Excel 结构要求

- 第 2 列：商家货号（商品名）
- 第 3 列：尺码
- 第 4 列起：日期（Excel 序列号格式，第 1 行为日期）

## 日期匹配规则

文件名中的数字是"几号"，程序会在 Excel 的所有日期列中匹配相同的"几号"。

**支持跨月表格**：Excel 表格可以跨越两个月，程序只比较"几号"，不关心是哪个月。

## 处理流程

1. 解析 PDF 文件名，提取日期号
2. 从 PDF 表格中提取订单数据
3. 按日期分组，合并相同商品名+尺码的数量
4. 在 Excel 中找到对应日期的列
5. 清空该列原有数据，填入新数据
6. 从 Excel 重新统计验证数量
7. 如不一致，检测并报告重复行问题
8. 保存 Excel 文件，删除已处理的 PDF 文件

## 数据验证

写入 Excel 后会自动验证：
- 重新统计 Excel 中该列的总数量
- 与 PDF 提取的数量对比
- 如不一致，显示差异并检测 Excel 中的重复行

## 常见问题

**数量不一致**：Excel 中同一商品+尺码出现多次，导致总和翻倍。检查 Excel 是否有重复行。

## 依赖

- Python 3.12
- pdfplumber
- openpyxl

```bash
pip install pdfplumber openpyxl
```

## 示例代码

```python
import pdfplumber
import openpyxl
import re
from datetime import datetime, timedelta
from collections import defaultdict

def extract_day_from_filename(filename):
    """从文件名提取日期数字"""
    # 匹配 "X号" 格式
    match = re.search(r'(\d+)号', filename)
    if match:
        return int(match.group(1))
    # 匹配开头数字（遇到 - 停止）
    match = re.search(r'^(\d+)(?:-|$)', filename)
    if match:
        return int(match.group(1))
    return None

def find_column_for_date(ws, day):
    """根据日期(几号)找到对应的 Excel 列"""
    base = datetime(1899, 12, 30)
    for col in range(4, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and isinstance(val, (int, float)):
            actual_date = base + timedelta(days=int(val))
            if actual_date.day == day:
                return col
    return None

def extract_pdf_data(pdf_path):
    """从 PDF 提取订单数据"""
    data_list = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and len(row) >= 6:
                        if row[0] in ['序号', '合计', None, ''] or str(row[0]).startswith('合计'):
                            continue
                        try:
                            skc_col = row[2]
                            attr_col = row[4]
                            qty_col = row[5]
                            if skc_col and attr_col and qty_col:
                                skc_parts = skc_col.split('\n')
                                product_name = skc_parts[1].strip() if len(skc_parts) >= 2 else skc_col
                                attr_clean = attr_col.replace('\n', '')
                                size_match = re.search(r'-([XS|S|M|L|XL]+)-', attr_clean)
                                if size_match:
                                    size = size_match.group(1)
                                else:
                                    size_match = re.search(r'-([XS|S|M|L|XL]+)$', attr_clean)
                                    if size_match:
                                        size = size_match.group(1)
                                    else:
                                        continue
                                if size and product_name:
                                    data_list.append({
                                        'product': product_name,
                                        'size': size,
                                        'qty': int(qty_col)
                                    })
                        except:
                            pass
    return data_list
```

## 注意事项

- 文件名中的日期在 Excel 表格中必须存在
- 如果同一个月出现两次（如 4 月 1 日和 5 月 1 日），会匹配第一次出现的列
- 处理完成后会自动删除已处理的 PDF 文件
