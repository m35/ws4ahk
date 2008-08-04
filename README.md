# ws4ahk

Include VBScript or JScript code directly in your Autohotkey program!
No temporary files. Full and easy access to COM.

***Note: ws4ahk is made for Autohotkey Basic. It may work with Autohotkey_L, but it is not supported.***

I've been thinking long and hard about using Microsoft Scripting Control to provide easy COM usage to AHK, and the more I have, the more it became the ultimate choice--far better than anything I could develop myself.

Pros
* Automatic objct management (objects are automatically deallocated--no memory leaks!)
* Able to use either VBScript or JScript to write COM related code
* ByRef argument handling is all taken care of
* Almost all `VARIANT` handling is taken care of
* Can very easily write compound COM statements (e.g. `objExcel.Workbooks.Add().Sheets(1).Cells(1,1).Value = 50`)
* There's no need for an extra dll. It can be done entirely in AHK.
* Much more easily implemented!

The ONLY disadvantage with using Microsoft Scripting Control is that there may be some computers that do not have it installed (but probably not very many). HOWEVER, not only is it available to download from Microsoft and very easily installed, but thanks to the  `WS_CreateObjectFromDll()` function, it doesn't even have to be installed! If a computer doesn't have Microsoft Scripting Control installed, you just need to supply the msscript.ocx file in the same folder as your script, and it will still work perfectly. No need to register the OCX. Your script remains completely portable!
