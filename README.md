# PdfProcessor v0.2
a simple tool for pdf processing, based on PyMuPDF & tk

更新日志v0.2:
新增功能:
1. 新增了gui界面
2. 新增了拆分pdf时的功能(现在支持等间隔拆分和指定范围提取)
3. 新增了pdf单页操作,预览,旋转,文字提取等功能
4. 支持了拖拽打开文件
5. 使用的核心库从PyPDF2改为了PyMuPDF
6. License从MIT改为了GPLv3
修复bug:
1. 修复了指定页码拆分pdf时无法正确处理结尾页数的问题