@echo off
chcp 65001 >nul
echo 正在安装依赖...
pip install pdfplumber openpyxl pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo 正在打包...
pyinstaller --onefile --windowed --name "PDF订单统计工具" pdf_tool.py

echo.
echo 打包完成！
echo 生成的文件: dist\PDF订单统计工具.exe
pause
