# -*- coding: utf-8 -*-
"""
PDF订单数据汇总到Excel

用法:
    python process.py <日期> [PDF文件...]

示例:
    python process.py 10              # 处理当前目录下所有"10号单*.pdf"
    python process.py 10 10号单1.pdf 10号单2.pdf  # 指定PDF文件
"""

import pdfplumber
import openpyxl
import re
import sys
import os
import glob
from collections import defaultdict
from datetime import datetime, timedelta


def extract_pdf_data(pdf_path):
    """从PDF提取订单数据"""
    data_list = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    if row and len(row) >= 6:
                        # 跳过标题行、合计行
                        if row[0] in ['序号', '合计', None, ''] or str(row[0]).startswith('合计'):
                            continue
                        try:
                            skc_col = row[2]      # 平台SKC/商家货号
                            attr_col = row[4]     # 属性集
                            qty_col = row[5]      # 数量

                            if skc_col and attr_col and qty_col:
                                # 提取商品名 (第二行)
                                skc_parts = skc_col.split('\n')
                                product_name = skc_parts[1].strip() if len(skc_parts) >= 2 else skc_col

                                # 去掉属性集中的换行符，用正则提取尺码
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


def find_column_for_date(ws, day):
    """根据日期(几号)找到对应的Excel列
    从Excel所有日期列中，找到"几号"匹配的文件名数字
    支持跨月表格（如4月20日-5月18日）
    """
    base = datetime(1899, 12, 30)

    for col in range(4, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val and isinstance(val, (int, float)):
            # Excel序列号转日期
            actual_date = base + timedelta(days=int(val))
            # 只比较"几号"，不比较月份
            if actual_date.day == day:
                return col
    return None


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    # 解析日期参数
    day = int(sys.argv[1])

    # 获取PDF文件列表
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if len(sys.argv) > 2:
        pdf_files = [os.path.join(script_dir, f) if not os.path.isabs(f) else f
                     for f in sys.argv[2:]]
    else:
        # 自动匹配 "{day}号单*.pdf" 文件
        pattern = os.path.join(script_dir, f"{day}号单*.pdf")
        pdf_files = sorted(glob.glob(pattern))

    if not pdf_files:
        print(f"未找到 {day} 号的PDF文件")
        return

    print(f"处理 {day} 号订单...")
    print(f"PDF文件: {[os.path.basename(f) for f in pdf_files]}")

    # 提取并合并数据
    all_data = []
    for pdf_path in pdf_files:
        data = extract_pdf_data(pdf_path)
        total_qty = sum(d['qty'] for d in data)
        print(f"  {os.path.basename(pdf_path)}: {len(data)}条, {total_qty}件")
        all_data.extend(data)

    merged_data = defaultdict(int)
    for item in all_data:
        merged_data[(item['product'], item['size'])] += item['qty']

    print(f"\n合并后: {len(merged_data)}条, {sum(merged_data.values())}件")

    # 读取Excel
    excel_path = os.path.join(script_dir, 'A1.xlsx')
    wb = openpyxl.load_workbook(excel_path)
    ws = wb['Sheet1']

    # 找到目标列
    target_col = find_column_for_date(ws, day)
    if not target_col:
        print(f"未找到 {day} 号对应的列")
        return
    print(f"目标列: 第{target_col}列")

    # 建立Excel商品索引
    excel_products = {}
    for row_num in range(2, ws.max_row + 1):
        product = ws.cell(row=row_num, column=2).value
        size = ws.cell(row=row_num, column=3).value
        if product and size:
            excel_products[(product, size)] = row_num

    # 清空目标列原有数据
    for row_num in range(2, ws.max_row + 1):
        ws.cell(row=row_num, column=target_col).value = None

    # 匹配并填写
    matched = 0
    unmatched = []
    for (product, size), qty in merged_data.items():
        key = (product, size)
        if key in excel_products:
            row_num = excel_products[key]
            ws.cell(row=row_num, column=target_col).value = qty
            matched += 1
        else:
            unmatched.append((product, size, qty))

    print(f"\n匹配成功: {matched}条")
    if unmatched:
        print(f"未匹配: {len(unmatched)}条")
        for product, size, qty in unmatched:
            print(f"  {product}, {size}, {qty}")

    # 保存
    wb.save(excel_path)
    print(f"\n已保存到 {excel_path}")

    # 删除已处理的PDF文件
    for pdf_path in pdf_files:
        os.remove(pdf_path)
        print(f"已删除: {os.path.basename(pdf_path)}")


if __name__ == '__main__':
    main()
