XDOC Office Server
==========================================

简介
----------------------------------------------
一个JavaEE服务，将Office文档转换为PDF，格式兼容性好。
![docx](https://raw.githubusercontent.com/myeeboy/xoffice/master/web/image/docx.png)
![pdf](https://raw.githubusercontent.com/myeeboy/xoffice/master/web/image/pdf.png)

安装部署
----------------------------------------------
1. 安装微软Office 2010或以上版本
2. 解压release/xoffice.zip，运行startup.bat

调用
----------------------------------------------
http://locahost:9090/xoffice?_xformat=*文档格式*&_file=*文档地址*&_watermark=*水印*

文档格式：doc、docx、xls、xlsx、ppt、pptx

文档地址：http协议地址，需要用UTF-8编码

水印：pdf水印文本

例：http://locahost:9090/xoffice?_xformat=docx&_file=http%3A%2F%2Flocahost%3A9090%2Fxoffice%2Fdemo.docx&_watermark=https%3A%2F%2Fview.xdocin.com

开源协议
----------------------------------------------
**MIT**

技术支持
----------------------------------------------
https://view.xdocin.com
