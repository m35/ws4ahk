;#Include CoHelper.ahk
#Include EasyScript.ahk
#Include ES_IEControl.ahk

If (!ES_Initialize())
	Msgbox % ES_Initialize("VBScript", "C:\WINDOWS\system32\msscript.ocx")

Msgbox % UrlHistoryEnum()


; .........................
; Add array to string
ArrayToString =
(
Function ArrayToString(ary)
	s = ""
	For Each v in ary
		s = s & v
	Next
	ArrayToString = s
End Function
)

If (!ES_Exec(ArrayToString))
	Msgbox % ES_Error()
; .........................

/*
adapter := CreateObjectFromDll("CoScriptAdapter.dll", "{18AB6A4F-0FEE-11D4-9D21-009027133993}")
Msgbox % "adapter=" adapter "  error=" GetComError()
IfEqual adapter,
{
	ES_Uninitialize()
	Return
}	
Msgbox % "AddObject adapter: " ES_AddObject(adapter, "adapter")

puh := CreateObject("{3C374A40-BAE4-11CF-BF7D-00AA006946EE}", "{3C374A41-BAE4-11CF-BF7D-00AA006946EE}")

Msgbox % "AddObject puh: " ES_AddObject(puh, "puh")

If (!ES_Exec("set x = puh"))
	Msgbox % A_lineNumber ":set..."  ES_Error()
*/

;If (!ES_Exec("hist = adapter.WrapObject(puh);"))
;	Msgbox % A_lineNumber ":Set hist..."  ES_Error()

	
;If (!ES_Exec("Set puh = adapter.CreateAndWrap(%s)", "{3C374A40-BAE4-11CF-BF7D-00AA006946EE}"))
;	Msgbox % ES_Error()


;ppvDllObj := _CreateObjectFromDll("TestObj.dll", "{4760F41F-7D56-4869-AFA9-5210675D9BDF}")
;Msgbox % ppvDllObj

;blnSuccess := ES_Exec("Set WSH = CreateObject(""WSH.WScript"")")
;If (Not blnSuccess)
;	Msgbox % ES_Error()

;ES_Eval(ret, "%v.Msgbox(%s) ' hello comment with % in it", "myobj", "Happy day!")

;ES_Eval(ret,"%s%a%%b%sc%d%%s", 10)

;Msgbox % IScriptControl_AddObject(__iScriptControlObj__, "testobj", ppvDllObj, -1)

;ES_Exec("s = cstr(""wanda"")")
;ES_Exec("i = clng(100)")
;ES_Exec("call testobj.RefArgs(s, i)")
;__iScriptErrorObj__ := IScriptControl_Error(__iScriptControlObj__)
;Msgbox % "Column: " IScriptError_Column(__iScriptErrorObj__)
;Msgbox % "Number: " IScriptError_Number(__iScriptErrorObj__)

__GetUniqueTempVar()
{
	static iTempIndex := 0
	Critical On
	sTempName := "TempVar" iTempIndex++	
	Critical Off
	Return sTempName
}

;Msgbox % ES_Exec("GetRef(""Return2Str"")")

;ES_Exec("ReDim ary(3)")
;ES_Exec("ary(0) = ""first""")
;Msgbox % ES_Eval("ary") 

;oFSO := CreateObject("Scripting.Dictionary")
;
;Msgbox % oFSO
;Msgbox % Get(oFSO, "Count")
;Call(oFSO, "Add", 1, """fish""")
;Msgbox % Get(oFSO, "Count")
;Msgbox % Get(oFSO, "Item", 1)

/*

CreateObject(sProgId)
{
	Return ES_Eval("CreateObject(""" . sProgId . """)")
}

GetObject(sProgId, sPathName="")
{
	Return ES_Eval("GetObject(""" . sProgId . """)")
}

Get(sObjHandle, sCall, Arg1="""", Arg2="""", Arg3="""")
{
	If (Arg1 != """")
	{
		sCall := sCall "("
		Loop 3
		{
			If (Arg%A_Index% = """")
				Break
			If (A_Index = 1)
				sCall .= Arg%A_Index%
			Else
				sCall .= ", " Arg%A_Index%
		}
		sCall .= ")"
	}
	Return ES_Eval(sObjHandle "." sCall)
}
;GetArray()
;Put(sObjHandle, sCall, sValue)
Call(sObjHandle, sCall, Arg1="""", Arg2="""", Arg3="""")
{
	sCall := sCall " "
	Loop 3
	{
		If (Arg%A_Index% = """")
			Break
		If (A_Index = 1)
			sCall .= Arg%A_Index%
		Else
			sCall .= ", " Arg%A_Index%
	}
	ES_Exec(sObjHandle "." sCall)
}
;
;objWB := Invoke(objXL, "Workbooks.Add()")
;objWB := Invoke(objXL, "Workbooks.Add()")
;
;Invoke(objXL, "Workbooks.Add()")
;Eval(objXL, 
;
;

*/

ES_Uninitialize()

