# PDF订单数据汇总

## 一、需求说明

### 任务目标
将PDF订单文件中的数据提取并汇总到Excel表格的对应日期列。

### 数据来源
PDF文件命名格式：`{日期}号单{序号}.pdf`（如 `10号单1.pdf`、`10号单2.pdf`）

### 数据提取规则
从PDF表格中提取三列数据：

| 列名 | 提取方式 |
|------|----------|
| 平台SKC/商家货号 | 取第二行作为商品名 |
| 属性集 | 用正则提取尺码（XS/S/M/L/XL） |
| 数量 | 直接取值 |

### 填写规则
1. 填写前先清空目标日期列的所有已有数据
2. 在Excel中找到对应日期的列（第1行为日期）
3. 通过「商品名 + 尺码」匹配Excel中的行
4. 将数量填入对应单元格
5. 相同商品名+尺码的数据合并（数量相加）
6. 执行完毕后删除已处理的PDF文件

---

## 二、执行思路

### Excel结构
- 第2列：商家货号（商品名）
- 第3列：尺码
- 第4列起：日期（Excel序列号格式）

### 日期与列号对应
Excel序列号转换：`日期 → 序列号 → 匹配列标题`

| 日期 | 序列号 | 列号 |
|------|--------|------|
| 4月8日 | 46120 | 18 |
| 4月9日 | 46121 | 19 |
| 4月10日 | 46122 | 20 |
| ... | ... | ... |

### 执行方法

**方式一：GUI工具（推荐）**
```bash
python pdf_tool.py
```
1. 选择目标Excel文件
2. 添加PDF文件（自动识别文件名中的日期）
3. 点击"开始处理"

**方式二：命令行**
```bash
# 处理指定日期的所有PDF
python process.py 10

# 或指定具体PDF文件
python process.py 10 10号单1.pdf 10号单2.pdf
```

### 环境要求
- Python 3.12
- 依赖：`pdfplumber`, `openpyxl`

```bash
pip install pdfplumber openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 打包为exe（Windows）
```bash
pip install pyinstaller
pyinstaller --onefile --windowed pdf_tool.py
```
生成的exe文件在 `dist/` 目录下，可独立运行。
