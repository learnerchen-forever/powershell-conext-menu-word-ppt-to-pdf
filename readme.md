# Office Word/PPT 转换 PDF

## 功能概述

在很多时候，需要把MS Word, PowerPoint 文件转换为PDF格式。比如，在Obsidian或Logseq这类的笔记应用中，PDF是可以直接导入到知识库的格式。


通常转换的过程中，每一个文件，都需要打开Office应用做转换，还是有些麻烦。

为了操作的简单便捷，把这个功能做到文件管理器的上下文（右键）菜单里。选择一个文件，或转换目录中所有的Word/PPT文件。



## 使用方法
1. 复制 AddContextMenu.reg 和 convert2pdf.ps1到windows系统中
2. 根据你的具体保存convert2pdf.ps1所在目录的位置，修改AddContextMenu.reg中的相应路径
3. 双击AddContextMenu.reg, 导入注册表，添加右键菜单

## 注意事项
1. 需要安装MS Office
2. 需要管理员权限
3. 需要安装PowerShell
4. 注意 AddContextMenu.reg 文件的编码格式，必须是 __UTF-16 LE__ ，否则右键菜单的中文会出现乱码
5. Windows 11中，右键菜单，“显示更多选项” ,才能显示。
6. 如果文件本身存在问题，脚本无法处理这类错误。用户自己要确保文件是可以正常打开，没有损坏。



## Version
- 1.0.0, 2024-08-17
