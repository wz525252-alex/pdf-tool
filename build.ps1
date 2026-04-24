# 设置编码
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "正在安装依赖..." -ForegroundColor Yellow
pip install pdfplumber openpyxl pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple

Write-Host "`n正在打包..." -ForegroundColor Yellow
pyinstaller --onefile --windowed --name "PDF订单统计工具" pdf_tool.py

Write-Host "`n打包完成！" -ForegroundColor Green
Write-Host "生成的文件: dist\PDF订单统计工具.exe" -ForegroundColor Cyan
Read-Host "按回车键退出"
