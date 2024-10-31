# Convert MS Office Word/PPT to PDF with PowerShell and context menu

[中文](readme-cn.md)

## Functionality

Microsoft Office document like Word (doc/docx) and PowerPoint (ppt/ppx) are widely used in our daily work. While some sutiation, PDF version might be more proper for us. 

For example, in note taking tools like `obdisian` or `logseq`, PDF version makes the reading, annotation and reference much easier, and saves a lot of time. 

Office itself provides a way to convert these files to PDF, but you have to open document in the application first. If we are going to convert a batch of files, it will be a little bit tedious.
 
We provide a PowerShell script to do the job. and add a context menu to Windows Explorer. It allows you to convert a single file or all Word/PPT files in a directory. Just right click the file or directory.

You don't need to install any software. Just download the script and edit it according to your configuration.



## Installation
1. Download following 2 files to your Windows system:
    - convert2pdf.ps1
    - AddContextMenu.reg
2. save them to a folder, for instance: `C:\\convert2pdf` 
3. replace the path in AddContextMenu.reg with the actual path where you saved the script.
4. double click AddContextMenu.reg, and import the registry, add context menu.


## note
1. You must have Microsoft Office installed
2. to inject .reg file, you must have admin privilege
3. Powershell is required
4. The encoding of AddContextMenu.reg must be __UTF-16 LE__, else, the context menu text could be in error, if you are using Chinese or other language. 
5. In Windows 11, the menu is in second level, that is you should cick on the "more options" to see the context menu. 
6. If the office document is corrupted, the script cannot handle it. User should ensure the file is readable. 



## Version history
- 1.0.0, 2024-08-17 first release
- 1.0.1, 2024-10-31 update to English version, to release on GitHub

