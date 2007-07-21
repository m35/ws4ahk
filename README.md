# EasyScript

I've been working on using Microsoft Scripting Control for COM, and so far things seem to be ok. I'm not quite sure what to call it, so for now I've named it "EasyScript". The `#Include` file is looking to be a beefy 1000 lines. Right now the API looks something like this:

```
LoadEasyScript(Language)
UnloadEasyScript()

ES_Exec(VB or JS string) ; for no return value
ES_Eval(VB or JS string) ; for return value

ES_AddDispObject(IDispatchPtr)
ES_AddObject(ClsId, Name)
ES_AddObjectFromDll(ClsId, Name)

VBStr(String)
JSStr(String)
```

`ES_AddObject()` is useful for creating objects by ClassId, since you can't do that in VBScript/JScript. `ES_AddObjectFromDll()` is thrown in just because it's easy, but I can't imagine anyone actually using it. `ES_AddDispObject()` is useful for putting ActiveX controls under the care of VBScript/JScript (assuming that works).

`VBStr()` and `JSStr()` wrap strings in quote and escapes any quotes in the actual string.

Here's a short usage example

```
sFile := "C:\WINDOWS"
ES_Exec("Set oFSO = CreateObject('Scripting.FileSystemObject')")
IsFolder := ES_Eval( "oFSO.FolderExists(" VBStr(sFile) ")" )
; IsFolder = 1
ES_Exec("Set ByRefObj = CreateObject('Object.With.Byref.Param')")
ES_Exec("ByRefObj.ByRefMethod ", ByRefArg )
```

As you can see I'm using single quotes (`'`) inside the strings. This isn't normally available in VBScript, but to make things easier I may try to allow for it. No guarantee it will end up in a final version.

`ByRef` arguments can be passed as variables between strings. After the call, the `byref` variable will be updated with the new value.

I still haven't looked into how to handle script errors.

All this may be too much encapsulation however. It may be better to just expose the Scripting Control methods and let the scripters write VBScript/JScript however they like.

I dunno, I'm still working on it. I'm curious what others think.


Also, It occurred to me that while this "EasyScript" API is handy for the COM basics, it still lacks support for creating ActiveX controls and handling COM Events. I haven't looked too closely at Sean's ActiveX control related scripts, so maybe those will be fine for handling that. I also haven't looked at COM event handling.

But I think those things are better suited for separate `#Includes`, and the separate APIs can be written to interface with each other.

In any case, I may have bitten off more than I can chew with all this COM stuff, so I'm going to have to pass on delving into COM Events, at least for now.


-------------------------------------------------------------------------------

The functions haven't really changed, but their parameters have. Plus I thought that adding a simple 'printf' style syntax would make the code much easier to write (no need for lots of ugly double quotes and string concatenation).

```
vbwmp := "oWMP"
If (!ES_Exec("Set %v = CreateObject(%s)", vbwmp, "WMPlayer.OCX"))
	Msgbox % ES_Error()
If (!ES_Eval(pwmp, vbwmp))
	Msgbox % ES_Error()
```

`%v` inserts the value, and `%s` inserts the value wrapped in quotes, with special characters escaped.

`ES_Eval()` will return the value of most script variables. I don't think I'll implement getting return values for Dates and Currency variables, and especially not for Arrays (you should write a script function to convert the array to something you can use before returning it).

You can create an object in the script, put that same object into a COM control, and then control from the script.

```
DEdit := "DEdit"
If (!ES_Exec("Set %v = CreateObject(%s)", DEdit, "DhtmlEdit.DhtmlEdit"))
		Msgbox % ES_Error()
If (!ES_Eval(ppvDEdit, "DEdit"))
		Msgbox % ES_Error()
		
hContainerCtrl := AtlAxCreateContainer(hWnd, 0, 25, 800, 575) 
AtlAxAttachControl(ppvDEdit, hContainerCtrl)

DE_LoadUrl(DEdit, "http://www.autohotkey.com")

DE_LoadUrl(sDhtmlEdit, url)
{
	If (!ES_Exec("%v.LoadUrl %s", sDhtmlEdit, url))
		Msgbox % ES_Error()
}
```
You could alternately create an object using one of the EasyScript's or Sean's create/get object functions, then pass the object into the script (via AddObject) to be controlled there.

You can always execute more than one line of code at a time.
```
DE_BrowseMode(sDHtmlEdit)
{
	sCode = 
(
If `%v.Browsemode = 0 Then
 `%v.Browsemode = 1
Else
 `%v.Browsemode = 0
End If
)
	If (!ES_Exec(sCode, sDHtmlEdit, sDHtmlEdit, sDHtmlEdit))
		Msgbox % ES_Error()
}
```
This allows you to add functions inside your script. Heck, you could even create classes inside the script, instantiate it, and return it for use in AHK. I haven't pondered just what this could mean. I wonder if it would help with COM event handling at all...?

But speaking of event handling: There is the VBScript `GetRef()` function which may be used to insert a script call-back function for objects that allow it. These are usually the "OnClick" kind of functions. From what I've seen, it's still limited to a subset of the available object events. There are many other events I can't get access to.

The code is currently working, but I'm still running tests.

