---
name: pdf-to-excel
description: 将 PDF 订单文件中的数据提取并汇总到 Excel 表格。触发场景：用户要求从 PDF 提取数据写入 Excel、PDF 订单汇总、PDF 表格数据导出、按日期填写 Excel 数据、PDF 数据统计。支持跨月表格和多种文件名格式。
---

# PDF 订单数据汇总到 Excel

## 概述

将 PDF 订单文件中的表格数据提取出来，按日期汇总到 Excel 表格的对应列中。支持自动验证数据一致性。

## 触发场景

- "从 PDF 提取数据写入 Excel"
- "PDF 订单汇总到表格"
- "按日期填写 Excel 数据"
- "PDF 表格数据导出"
- "统计 PDF 中的订单数据"

## 支持的文件名格式

| 格式 | 示例 | 说明 |
|------|------|------|
| `{日期}号单{序号}.pdf` | `10号单1.pdf` | 传统的"X号单"格式 |
| `{日期}-{名称}.pdf` | `1-单1-发货单.pdf` | 日期为开头数字，解析第一个 `-` 前的数字 |

## 数据提取规则

从 PDF 表格中提取三列数据：

| 列位置 | 列名 | 提取方式 |
|--------|------|----------|
| 第3列 | 平台SKC/商家货号 | 取第二行作为商品名 |
| 第5列 | 属性集 | 用正则提取尺码（XS/S/M/L/XL） |
| 第6列 | 数量 | 直接取值 |

## Excel 结构要求

| 列 | 内容 |
|----|------|
| 第2列 | 商家货号（商品名） |
| 第3列 | 尺码 |
| 第4列起 | 日期（Excel 序列号格式，第1行为日期） |

## 处理流程

```
PDF文件 → 解析文件名(提取日期) → 提取表格数据 → 合并相同商品+尺码 → 写入Excel → 验证数量 → 完成
```

1. 解析 PDF 文件名，提取日期号
2. 从 PDF 表格中提取订单数据
3. 按日期分组，合并相同商品名+尺码的数量
4. 在 Excel 中找到对应日期的列
5. 清空该列原有数据，填入新数据
6. 从 Excel 重新统计验证数量
7. 如不一致，检测并报告重复行问题
8. 保存 Excel 文件，删除已处理的 PDF 文件

## 数据验证

写入 Excel 后自动验证：
- 重新统计 Excel 中该列的总数量
- 与 PDF 提取的数量对比
- 如不一致，显示差异并检测 Excel 中的重复行

## 常见问题

| 问题 | 原因 | 解决方案 |
|------|------|----------|
| 数量不一致 | Excel 中同一商品+尺码出现多次 | 检查并删除 Excel 重复行 |
| 日期列未找到 | 文件名日期在 Excel 中不存在 | 确认 Excel 包含该日期列 |
| 尺码提取失败 | 属性集格式不匹配 | 检查 PDF 中属性集格式 |

## 依赖

```bash
pip install pdfplumber openpyxl
```

## 核心代码

### 文件名解析

```python
def extract_day_from_filename(filename):
    # 匹配 "X号" 格式
    match = re.search(r'(\d+)号', filename)
    if match:
        return int(match.group(1))
    # 匹配开头数字（遇到 - 停止）
    match = re.search(r'^(\d+)(?:-|$)', filename)
    if match:
        return int(match.group(1))
    return None
```

### 日期列查找

```python
def find_column_for_date(ws, day):
    base = datetime(1899, 12, 30)
    for col in range(4, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and isinstance(val, (int, float)):
            actual_date = base + timedelta(days=int(val))
            if actual_date.day == day:
                return col
    return None
```

### 数据合并

```python
merged_data = defaultdict(int)
for item in all_data:
    merged_data[(item['product'], item['size'])] += item['qty']
```

## 注意事项

- 文件名中的日期在 Excel 表格中必须存在
- 如果同一个月出现两次（如 4月1日 和 5月1日），会匹配第一次出现的列
- 处理完成后会自动删除已处理的 PDF 文件
- Excel 中避免同一商品+尺码出现多次，否则会导致数量翻倍

## 下载

Windows exe 文件：https://github.com/wz525252-alex/pdf-tool/releases
