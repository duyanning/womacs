For more information see: https://github.com/duyanning/womacs

```
womacs\
    womacs-dev.docm         used to develop VBA project, DO NOT open this document directly, use develop.bat to open it.
    install.bat             copy publish\womacs.dotm to Word\STARTUP
    uninstall.bat           remove womacs.dotm from Word\STARTUP
    develop.bat             open womacs.docm for developing
    develop_src\            VBA code exported from womacs-dev.docm
    publish\                staff generated to publish
        src\                VBA code to be embodied in womacs.dotm
        womacs.dotm         the resulting word template to publish
```

Note: womacs-dev.docm (for developing) and womacs.dotm (to publish)
have different filename extension. If womacs-dev.docm does NOT exist,
create it and import file `develop_src\DeveloperTools.bas`. Run marco
`import_vba_source_code_to_this_project` to import all sources from
develop_src.


If you get an error saying `user-defined type not defined` when
executing import, add a reference to `Microsoft Visual Basic For
Applications Extensibility 5.3`.

See also:
http://www.cpearson.com/excel/vbe.aspx


if you get an message saying

    Run-time error '6068':
    Programmatic access to Visual Basic Project is not trusted.

when calling macro `export_vba_source_code_from_this_project`, go to
`File > Options > Trust Center > Trust Center Settings > Developer
Macro Settings` and check `Trust access to the VBA project object
model`.



# 调试办法

首先，执行uninstall.bat以删除文件

`"%userprofile%\Application Data\Microsoft\Word\Startup\womacs.dotm"`

这个文件的自动加载会干扰我们开发。



区分两个模板：

- `womacs-dev.dotm`	这是写代码的地方（但我们并不直接打开该模板dotm，而是打开使用该模板的文档docx）

- `womacs.dotm`		这是发布的模板（发布前，直接打开`womacs-dev.dotm`，然后运行其中的generate宏来产生`womacs.dotm`）



