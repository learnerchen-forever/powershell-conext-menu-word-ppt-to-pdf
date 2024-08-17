# Office Word/PPT 转换 PDF

## 功能概述
在很多时候，需要把MS Word, PowerPoint 文件转换为PDF格式。每次打开Office应用，还是有些麻烦。比如，在Obsidian或Logseq这类的笔记应用中，PDF是可以比较好的导入到知识库的格式，选择一个文件，或者把一个目录下的这两类文件转换为PDF是一个经常的工作。

为了操作的简单便捷，充分利用Windows和MS Office的功能，把这个功能做到上下文（右键）菜单里。使用本功能，需要有安装好的MS Office。

由于，文件本身可能存在问题，脚本无法安全处理这类错误，所以，用户自己要确保文件是可以正常打开，没有损坏。

## 使用方法
1. 复制 AddContextMenu.reg 和 convert2pdf.ps1到windows系统中
2. 根据你的具体目录的位置，修改AddContextMenu.reg中的路径
3. 双击AddContextMenu.reg, 导入注册表，添加右键菜单

## 注意事项
1. 需要安装MS Office
2. 需要管理员权限
3. 需要安装PowerShell
4. 注意 AddContextMenu.reg 文件的编码格式，必须是 __UTF-16 LE__ ，否则右键菜单的中文会出现乱码
5. Windows 11中，右键菜单，“显示更多选项” ,才能显示。



## Version
- 1.0.0, 2024-08-17
