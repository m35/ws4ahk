;               ***********************************************
;               *   STDLIB Windows Scripting for Autohotkey   *
;               *              v0.11  beta                    *
;               ***********************************************

/****h* /ws4ahk
* About
*	Windows Scripting for Autohotkey (stdlib) v0.11 beta
*	
*	Requires Autohotkey v1.0.47 or above.
*	
*	This module contains functions to embed VBScript or JScript into your AHK
*	program, and as such, provides simple access to COM though these languages. 
*	This module also provides functions to create COM controls which can be 
*	controlled by VBScript or JScript.
*	
*	Note that this module requires use of the "Microsoft Scripting Control" 
*	which is usually installed on most machines. In the rare case it is not 
*	installed on a system, it may be downloaded from Microsoft and installed.
*	
*	http://www.microsoft.com/downloads/details.aspx?FamilyId=D7E31492-2595-49E6-8C02-1426FEC693AC
*	
*	As an alternative, the Microsoft Scripting Control file "msscript.ocx" may
*	be used directly (e.g. placed in the same folder as the AHK script), so  
*	there is no need to actually install it.
*	
* Links
*	List of Automation errors
*	http://support.microsoft.com/kb/186063
*
*	VBScript Language Reference
*	http://msdn2.microsoft.com/en-us/library/d1wf56tt.aspx
*	
*	JScript Language Reference
*	http://msdn2.microsoft.com/en-us/library/yek4tbz0.aspx
*	
*	The MSDN guru on WSH
*	http://blogs.msdn.com/ericlippert/archive/2004/07/14/183241.aspx
* To Do
*	* Figure out Locale ID handling (e.g. try using English VB in German locale)
*	* Create test suite
*	* Make variable naming conventions more consistent
******
*/

IID_IDispatch := "{00020400-0000-0000-C000-000000000046}"

; ..............................................................................
/****** ws4ahk/WS_Initialize
* Description
*	Initializes the Windows Scripting environment.
* Usage
*	WS_Initialize( [ sLanguage = "VBScript" [, sMSScriptOCX ] ] )
* Parameters
*	* sLanguage -- (Optional) (String) Either "VBScript" or "JScript".
*	* sMSScriptOCX -- (Optional) (String) Path to msscript.ocx file.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	This function must be called before any other functions may be used.
*	
*	Normally the scripting environment is setup using the installed 
*	msscript.ocx file on the system. Alternatively, you may specify the path 
*	directly to a msscript.ocx file, even if it is not registered with the
*	system (useful if the user does not have Microsoft Scripting Control
*	installed). Repeated calls to this function are ignored.
* Related
*	WS_Uninitialize
* Example
	WS_Initialize()
	; do stuff
	WS_Uninitialize()
	
	WS_Initialize("VBScript", "C:\Windows\system32\msscript.ocx")
	; do stuff
	WS_Uninitialize()
******
*/
WS_Initialize(sLanguage = "VBScript", sMSScriptOCX="")
{
	global __iScriptControlObj__, __iScriptErrorObj__, __sScriptLanguage__
	
	static ProgId_ScriptControl := "MSScriptControl.ScriptControl"
	static CLSID_ScriptControl  := "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}"
	static IID_ScriptControl    := "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}"
	
	IfNotEqual, __iScriptControlObj__,
		Return True ; Windows Scripting has already been initialized
				   
	; Init COM
	iErr := DllCall("ole32\CoInitialize", "UInt", 0, "Int")
	
	If (__IsComError("CoInitialize", iErr))
		Return False
	
	; Create Scripting Control
	If (sMSScriptOCX="")
		__iScriptControlObj__ := WS_CreateObject(CLSID_ScriptControl, IID_ScriptControl)
	Else
		__iScriptControlObj__ := WS_CreateObjectFromDll(sMSScriptOCX, CLSID_ScriptControl, IID_ScriptControl)
		
	IfEqual, __iScriptControlObj__,
		Return False

	; Set the language
	If (!__IScriptControl_Language(__iScriptControlObj__, sLanguage))
	{	; Failed to set language
		WS_Uninitialize()
		Return False
	}
		
	__sScriptLanguage__ := sLanguage
	
	; Get Error object
	__iScriptErrorObj__ := __IScriptControl_Error(__iScriptControlObj__)
	
	Return True
}


; ..............................................................................
/****** ws4ahk/WS_Uninitialize
* Description
*	Uninitializes the Windows Scripting environment 
*	and releases the allocated resources.
* Usage
*	WS_Uninitialize()
* Return Value
*	None ("").
* ErrorLevel
*	May hold an error description on error.
* Remarks
*	Call this function to free the memory used by this library. It is not
*	necessary to call this function before exiting your script (but it is 
*	good practice). This function may be called repeatedly.
* Related
*	WS_Initialize
* Example
	; see WS_Initialize example
******
*/
WS_Uninitialize()
{
	global __iScriptControlObj__, __iScriptErrorObj__   
	
	IfNotEqual __iScriptErrorObj__,
		__IUnknown_Release(__iScriptErrorObj__)
	
	IfNotEqual __iScriptControlObj__,
		__IUnknown_Release(__iScriptControlObj__)
	
	ErrLvl := ErrorLevel ; save ErrorLevel
	DllCall("ole32\CoUninitialize")
	ErrorLevel := ErrLvl ; restore ErrorLevel
	
	__iScriptControlObj__ := ""
	__iScriptErrorObj__   := ""
	__sScriptLanguage__   := ""
}


; ..............................................................................
/****** ws4ahk/WS_Exec
* Description
*	Executes scripting code.
* Usage
*	WS_Exec(sScriptCode [, value1 [, value2 [,...]]])
* Parameters
*	* sScriptCode - (String) Scripting code to execute.
*	* value1, value2, ... -- (Optional) Values to insert into the ScriptCode.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	There are two special codes that may be used within the ScriptCode: 
*	* %v - inserts the value
*	* %s - inserts the value wrapped in quotes, with special characters escaped.
*	These codes will be replaced with the value1, value2,... values.
*	Up to 10 values can be inserted into the ScriptCode
*
*	If WS_Initialize() has not been called, an error message will be displayed
*	and the program will exit.
* Related
*	WS_Eval, ScriptStr, VBStr, JStr
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	; These three lines execute the same script.
	WS_Exec("foo = ""bar""")
	WS_Exec("foo = %s", "bar")
	WS_Exec("%v = %s", "foo", "bar")
	
	; Using a block of code like this makes it easy to
	; add functions to the scripting environment.
	Code = 
	(
		foo = "bar"
		Sub MsgFoo()
			Msgbox foo
		End Sub
	)
	WS_Exec(Code)
	WS_Exec("MsgFoo")
******
*/
WS_Exec(sCode, arg1="`b`b", arg2="`b`b", arg3="`b`b", arg4="`b`b", arg5="`b`b"
             , arg6="`b`b", arg7="`b`b", arg8="`b`b", arg9="`b`b", arg10="`b`b"))
{
	global __iScriptControlObj__, __iScriptErrorObj__
	
	
	IfEqual __iScriptControlObj__,
	{
		Msgbox % "Windows Scripting has not been initialized!`nExiting application."
		ExitApp
	}
	
	; Merge the arguments into the code string
	iArg := 1
	iPos := 1
	Loop
	{
		val := arg%iArg%
		If (val = "`b`b")
			Break
		
		If (iPos := InStr(sCode, "%", True, iPos))
		{
			sNextChar := SubStr(sCode, iPos+1, 1)
			If (sNextChar == "v")
			{
				sCode := SubStr(sCode, 1, iPos-1) . val . SubStr(sCode, iPos+2)
				iArg++
				iPos += StrLen(val)
			}
			Else If (sNextChar == "s")
			{
				val := ScriptStr(val)
				sCode := SubStr(sCode, 1, iPos-1) . val . SubStr(sCode, iPos+2)
				iArg++
				iPos += StrLen(val)
			}
			Else
				iPos++
		}
		Else
			Break
	}
	; Run the code
	Critical, On ; For thread safty
	iErr := __IScriptControl_ExecuteStatement(__iScriptControlObj__, sCode)
	If (iErr = 0)
	{
		Critical, Off
		Return True
	}
	Else
	{
		; Probably an exception. Get the deatils.
		; TODO: Find out what HRESULT code(s) mean there is an exception.
		__HandleScriptError()
		Critical, Off
		Return False
	}
}


; ..............................................................................
/****** ws4ahk/WS_Eval
* Description
*	Evaluates scripting code and returns the result.
* Usage
*	WS_Eval(ByRef ReturnValue, sScriptCode [, value1 [, value2 [,...]]])
* Parameters
*	* ReturnValue -- (ByRef) Variable to receive the return value.
*	* sScriptCode -- (String) Scripting code to execute.
*	* value1, value2, ... -- (Optional) Values to insert into the ScriptCode.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	There are two special codes that may be used within the ScriptCode: 
*	* %v - inserts the value
*	* %s - inserts the value wrapped in quotes, with special characters escaped.
*	These codes will be replaced with the value1, value2,... values.
*	Up to 10 values can be inserted into the ScriptCode.
*	
*	ReturnValue will hold the result of an evaluation. Most return types are
*	handled. Unhandled types are: Array, Currency, Date, VARIANT*, and the 
*	mysterious DECIMAL* type. You must convert these unhandled types to another 
*	type (usually string) before they can be returned. 
*	
*	If an expression results in an unhandled types, the function will return 
*	True (because the expression was evaluated), but ReturnValue will be set 
*	to "#Unhandled return type#".
*	
*	Note also that if the expression returns an object, the object should be 
*	released via the WS_ReleaseObject() function when it is no longer needed.
*
*	If WS_Initialize() has not been called, an error message will be displayed
*	and the program will exit.
* Related
*	WS_Exec, ScriptStr, VBStr, JStr
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	Code = 
	(
		Function Square(x)
			Square = x * x
		End Function
	)
	WS_Exec(Code)
	WS_Eval(Ret, "Square(6)")
	Msgbox % Ret
******
*/
WS_Eval(ByRef xReturn, sCode, arg1="`b`b", arg2="`b`b", arg3="`b`b", arg4="`b`b", arg5="`b`b"
                            , arg6="`b`b", arg7="`b`b", arg8="`b`b", arg9="`b`b", arg10="`b`b")
{
	global __iScriptControlObj__, __iScriptErrorObj__
	
	IfEqual __iScriptControlObj__,
	{
		Msgbox % "Windows Scripting has not been initialized!`nExiting application."
		ExitApp
	}
	
	; Merge the arguments into the code string
	iArg := 1
	iPos := 1
	Loop
	{
		val := arg%iArg%
		If (val = "`b`b")
			Break
			
		If (iPos := InStr(sCode, "%", True, iPos))
		{
			sNextChar := SubStr(sCode, iPos+1, 1)
			If (sNextChar == "v")
			{
				sCode := SubStr(sCode, 1, iPos-1) . val . SubStr(sCode, iPos+2)
				iArg++
				iPos += StrLen(val)
			}
			Else If (sNextChar == "s")
			{
				val = ScriptStr(val)
				sCode := SubStr(sCode, 1, iPos-1) . val . SubStr(sCode, iPos+2)
				iArg++
				iPos += StrLen(val)
			}
			Else
				iPos++
		}
		Else
			Break
	}

	; Run the code
	Critical, On ; For thread safty
	iErr := __IScriptControl_Eval(__iScriptControlObj__, sCode, varReturn)
	If (iErr = 0)
	{
		If (!__UnpackVARIANT(varReturn, xReturn))
			xReturn := "#Unhandled return type#"
		Critical, Off
		Return True
	}
	Else
	{
		; Probably an exception. Get the deatils.
		__HandleScriptError()
		Critical, Off
		Return False
	}
}


; ..............................................................................
/****ix* Internal Functions/__HandleScriptError
* Description
*	Sets ErrorLevel with the last ScriptError.Description.
* Usage
*	__HandleScriptError()
* Return Value
*	None ("").
* ErrorLevel
*	The ScriptError.Description text, or if there is no text, a default
*	automation error message with number.
* Remarks
*	Also clears the last script error.
* Related
*	WS_Exec, WS_Eval
******
*/
__HandleScriptError()
{
	global __iScriptErrorObj__

	sErrorDesc := __IScriptError_Description(__iScriptErrorObj__)
	IfEqual, sErrorDesc,
		sErrorDesc := "Automation error " __IScriptError_Number(__iScriptErrorObj__)
	__IScriptError_Clear(__iScriptErrorObj__)
	ErrorLevel := sErrorDesc
}


; ..............................................................................
/****** ws4ahk/ScriptStr
* Description
*	Wraps a string in quotes and escapes forbidden characters.
* Usage
*	ScriptStr(str)
* Parameters
*	* str -- (String) String to escape.
* Return Value
*	(String) Escaped string.
* ErrorLevel
*	Not changed.
* Remarks
*	Converts an Autohotkey string into a string usable in the scripting 
*	environment. Specifically, it escapes disallowed characters 
*	(e.g. quotes, carriage return) and, wraps the string in quotes.
*
*	This checks what scripting language is being used and calls the
*	appropriate language function (VBStr or JStr).
* Related
*	VBStr, JStr, WS_Exec, WS_Eval
* Example
	; see VBStr and JStr examples
******
*/
ScriptStr(s)
{
	global __sScriptLanguage__
	If (__sScriptLanguage__ == "VBScript")
		Return VBStr(s)
	Else If (__sScriptLanguage__ == "JScript")
		Return JStr(s)
	Else
		Return
}


; ..............................................................................
/****** ws4ahk/VBStr
* Description
*	Wraps a string in quotes and escapes disallowed characters for use in VBScript.
* Usage
*	VBStr(str)
* Parameters
*	* str -- (String) String to escape.
* Return Value
*	(String) Escaped string.
* ErrorLevel
*	Not changed.
* Remarks
*	Converts an Autohotkey string into a string usable in the scripting 
*	environment. Specifically, it escapes disallowed characters 
*	(e.g. quotes, carriage return) and, wraps the string in quotes.
* Related
*	ScriptStr, JStr, WS_Exec, WS_Eval
* Example
	#Include ws4ahk.ahk
	text := VBStr("this is a test")
	Msgbox % text
	; This is equivalent to
	; text := """this is a test"""
	

	text = 
	(
	Multi
	Line
	Text
	)
	text := VBStr(text)
	Msgbox % text
	; This is equivalent to
	; text := """Multi"" & vbLf & ""Line"" & vbLf & ""Text"""	
******
*/
VBStr(s)
{
	StringReplace, s, s, ", "", All
	StringReplace, s, s, `r, " & vbCr & ", All
	StringReplace, s, s, `n, " & vbLf & ", All
	StringReplace, s, s, `f, " & vbFormFeed & ", All
	StringReplace, s, s, `b, " & vbBack & ", All
	Return """" s """"
}


; ..............................................................................
/****** ws4ahk/JStr
* Description
*	Wraps a string in quotes and escapes disallowed characters for use in JScript.
* Usage
*	JStr(str)
* Parameters
*	* str -- (String) String to escape.
* Return Value
*	(String) Escaped string.
* ErrorLevel
*	Not changed.
* Remarks
*	Converts an Autohotkey string into a string usable in the scripting 
*	environment. Specifically, it escapes disallowed characters 
*	(e.g. quotes, carriage return) and, wraps the string in quotes.
* Related
*	ScriptStr, VBStr, WS_Exec, WS_Eval
* Example
	#Include ws4ahk.ahk
	text := JStr("this is a test")
	Msgbox % text
	; This is equivalent to
	; text := "\"this is a test\""
	

	text = 
	(
	Multi
	Line
	Text
	)
	text := JStr(text)
	Msgbox % text
	; This is equivalent to
	; text := "\"Multi\nLine\nText\""	
******
*/
JStr(s)
{
	StringReplace, s, s, \, \\, All
	StringReplace, s, s, ", \", All
	StringReplace, s, s, `r, \r, All
	StringReplace, s, s, `n, \n, All
	StringReplace, s, s, `f, \f, All
	StringReplace, s, s, `b, \b, All
	Return """" s """"
}


; ..............................................................................
/****** ws4ahk/WS_ErrMsg
* Description
*	Show a formatted Msgbox error.
* Usage
*	WS_ErrMsg(sSourceFile, iLine)
* Parameters
*	* sSourceFile -- (String) AHK source file name 
*	                      (use A_ScriptName, A_ScriptFullPath, or A_LineFile).
*	* iLine -- (String) Line in the source file (use A_LineNumber).
* Return Value
*	None ("").
* ErrorLevel
*	Not changed.
* Remarks
*	Shows a message box with formatted text with the current AHK script file,
*	line number, and current error in ErrorLevel.
* Related
*	WS_Exec, WS_Eval
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	If (!WS_Exec("Set x = ??"))
		WS_ErrMsg(A_ScriptName, A_LineNumber)
******
*/
WS_ErrMsg(sFile, iLine)
{
	Msgbox, , Windows Scripting Error
			, % "Scripting error on line " iLine " in file " sFile
			. "`nError: " ErrorLevel
}


; ..............................................................................
/****** ws4ahk/WS_AddObject
* Description
*	Adds a COM object to the scripting environment.
* Usage
*	WS_AddObject(pObject, sName [, blnGlobalMembers = False ] )
* Parameters
*	* pObject -- (Integer) Pointer to object to add.
*	* sName -- (String) The identifier that will be used in the  
*	                    scripting environment to refer to this object.
*	* blnGlobalMembers -- (Optional) (Boolean) Setting GlobalMembers to True will 
*	                      make all the members of the object global in the script.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	Adds an object created in AHK to the scripting environment. Setting
*	GlobalMembers to True will make all the members of the object global in
*	the script.
*	
*	This function is most useful after creating a COM control. The COM object
*	can be added to the scripting environment, and then the object can be
*	controlled via the script.
*	
*	If WS_Initialize() has not been called, an error message will be displayed
*	and the program will exit.
* Related
*	WS_ReleaseObject
* Example
	; see WS_CreateObject example
******
*/
WS_AddObject(pObject, sName, blnGlobalMembers = False)
{
	global IID_IDispatch, __iScriptControlObj__
	
	IfEqual __iScriptControlObj__,
	{
		Msgbox % "Windows Scripting has not been initialized!`nExiting application."
		ExitApp
	}
	
	Return __IScriptControl_AddObject(__iScriptControlObj__, sName, pObject
	                                                      , -blnGlobalMembers)
}


; ..............................................................................
/****** ws4ahk/WS_CreateObject
* Description
*	Creates a new instance of a COM object.
* Usage
*	WS_CreateObject(sProgIDorClassID [, sInterfaceID = IDispatch ] )
* Parameters
*	* sProgIDorClassID -- (String) Program ID (e.g. "Excel.Application") or 
*	                               Class ID (e.g. "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}").
*	* sInterfaceID -- (Optional) (String) Interface ID of the object to create
*	                                      (e.g. "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}").
* Return Value
*	* Success: (Integer) Pointer to the created object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	WS_ReleaseObject() should be called when the object is no longer needed.
* Related
*	WS_ReleaseObject
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	pSpObj := WS_CreateObject("SAPI.SpVoice")
	WS_AddObject(pSpObj, "oSpeak")
	WS_ReleaseObject(pSpObj)
	WS_Exec("oSpeak.Speak %s", "Hello world!")
******
*/
WS_CreateObject(sProgID_ClsId, sIId = "{00020400-0000-0000-C000-000000000046}")
{                                    ; ^ IDispatch                          
	global IID_IDispatch
	static CLSCTX_INPROC_SERVER   := 1
	static CLSCTX_INPROC_HANDLER  := 2
	static CLSCTX_LOCAL_SERVER    := 4
	static CLSCTX_INPROC_SERVER16 := 8
	static CLSCTX_REMOTE_SERVER   := 16
	
	If (InStr(sProgID_ClsId, "{")) ; Is it a CLSID?
	{
		If (!__CLSIDFromString(sProgID_ClsId, sbinClassId))
			Return ; unable to create binary class id
	}
	Else
	{
		If (!__CLSIDFromProgID(sProgID_ClsId, sbinClassId))
			Return ; unable to create binary class id
	}
	
	
	If (!__IIDFromString(sIId, sbinIId))
		Return ; unable to create binary interface id
	
	iErr := DllCall("ole32\CoCreateInstance"
					, "Str"  , sbinClassId
					, "UInt" , 0
					, "Int"  , CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER
					, "Str"  , sbinIId
					, "UInt*", ppv
					, "Int")
					
	If (__IsComError("CoCreateInstance", iErr))
		Return
	
	If (sIId = IID_IDispatch)
		Return __GetIDispatch(ppv)
	Else
		Return ppv
}


; ..............................................................................
/****** ws4ahk/WS_GetObject
* Description
*	Get the instance of an already existing COM object.
* Usage
*	WS_GetObject(sProgIDorClassID [, sInterfaceID = IDispatch ] )
* Parameters
*	* sProgIDorClassID -- (String) Program ID (e.g. "Excel.Application") or 
*	                               Class ID (e.g. "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}").
*	* sInterfaceID -- (Optional) (String) Interface ID
*	                                      (e.g. "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}").
* Return Value
*	* Success: (Integer) Pointer to the existing instance of the object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	WS_ReleaseObject() should be called when the object is no longer needed.
* Related
*	WS_ReleaseObject
* Example
	; This example works best if Microsoft Excel is installed.
	#Include ws4ahk.ahk
	WS_Initialize()
	pExcel := WS_GetObject("Excel.Application")
	
	If (pExcel = "")
	{
		Msgbox % "Excel isn't running."
	}
	Else
	{
		WS_AddObject(pExcel, "oExcel")
		WS_ReleaseObject(pExcel)
		WS_Eval(iWrkBks, "oExcel.Workbooks.Count")
		Msgbox % "You have " iWrkBks " workbook(s) open in Excel."
	}
******
*/
WS_GetObject(sProgID_ClsId, sIId = "{00020400-0000-0000-C000-000000000046}")
{                                 ; ^ IDispatch
	global IID_IDispatch
	
	; Get the binary form of class ID
	If (InStr(sProgID_ClsId, "{")) ; Is it a CLSID string?
	{
		If (!__CLSIDFromString(sProgID_ClsId, sbinClassId))
			Return ; unable to create binary class id
	}
	Else ; it's a Prog ID
	{
		If (!__CLSIDFromProgID(sProgID_ClsId, sbinClassId))
			Return ; unable to create binary class id
	}
	
	; Get the object
	iErr := DllCall("oleaut32\GetActiveObject"
				, "Str", sbinClassId
				, "UInt", 0
				, "UInt*", pIUnkn
				, "Int")
	
	; Failed?
	If (__IsComError("GetActiveObject", iErr))
		Return
	
	ppv := __IUnknown_QueryInterface(pIUnkn, sIId)
	
	__IUnknown_Release(pIUnkn)
	
	; Did QueryInterface fail?
	IfEqual, ppv,
		Return
	
	If (sIId = IID_IDispatch)
		Return __GetIDispatch(ppv)
	Else
		Return ppv
}


; ..............................................................................
/****** ws4ahk/WS_ReleaseObject
* Description
*	Frees a references to an object.
* Usage
*	WS_ReleaseObject(pObject)
* Parameters
*	* pObject - (Integer) Pointer to the object to be released.
* Return Value
*	* Success: (Integer) Number of remaining references to the object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0. 
*	* Failure: error description.
* Remarks
*	Has the same behavior as calling the internal function 
*	__IUnknown_Release(), but has a more accessible name.
* Related
*	WS_CreateObject
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	WS_Eval(oObj, "CreateObject(%s)", "Wscript.Shell")
	WS_ReleaseObject(oObj)
******
*/
WS_ReleaseObject(iObjPtr)
{
	Return __IUnknown_Release(iObjPtr)
}


; ## COM Controls ##############################################################
; These functions originally written by our resident COM guru Sean in
; the Autohotkey forums. They have ben expanded, commented, and renamed
; for easier reading. They have also been adjusted to use the WS4AHK COM
; API functions, and error checking has been added (eventually).

; These functions were small enough I figured they may as well just be included.
; Windows Scripting does not have to be initialized before using these functions.

; ..............................................................................
/****** ws4ahk/WS_InitComControls
* Description
*	Initializes COM controls.
* Usage
*	WS_InitComControls()
* Return Value
*	* Success: (Integer) nonzero.
*	* Failure: 0.
* ErrorLevel
*	Set to DllCall() result.
* Remarks
*	Must be called before calling WS_CreateComControlContainer.
*	
*	There is no harm in calling this function more than once.
* Related
*	WS_UninitComControls, WS_CreateComControlContainer
* Example
	; see WS_CreateComControlContainer example
******
*/
WS_InitComControls()
{
	; Check if atl dll has already been loaded.
	If !DllCall("GetModuleHandle", "Str", "atl")
		DllCall("LoadLibrary"    , "Str", "atl")
	; Initialize atl (it's ok to call this function more than once)
	Return DllCall("atl\AtlAxWinInit")
}


; ..............................................................................
/****** ws4ahk/WS_UninitComControls
* Description
*	Uninitializes COM controls.
* Usage
*	WS_UninitComControls()
* Return Value
*	None.
* ErrorLevel
*	Set to DllCall() result.
* Remarks
*	Unloads the atl library. Note: MSDN warns about a race condition that could
*	occur if two threads call this function at the same time.
* Related
*	WS_InitComControls
* Example
	; see WS_CreateComControlContainer example
******
*/
WS_UninitComControls()
{
	If hModule := DllCall("GetModuleHandle", "Str", "atl")
		DllCall("FreeLibrary", "UInt", hModule)
}


; ..............................................................................
/****** ws4ahk/WS_GetHWNDofComControl
* Description
*	Retrieves the hWnd (i.e. ahk_id) associated with a COM control object.
* Usage
*	WS_GetHWNDofComControl(pObject)
* Parameters
*	* pObject - (Integer) Pointer to a COM control object.
* Return Value
*	* Success: (Integer) The hWnd (i.e. ahk_id) of the control that hosts the COM object.
*	* COM related failure: None ("").
*	* Window related failure: 0 (NULL).
* ErrorLevel
*	* Success: 0
*	* COM related failure: COM error description.
*	* Window related failure: DllCall() result. 
* Remarks
*	
* Related
*	WS_GetComControlInHWND, WS_AttachComControlToHWND
* Example
******
*/
WS_GetHWNDofComControl(pComObject)
{ 
	static IID_IOleWindow := "{00000114-0000-0000-C000-000000000046}"
	pOleWin := __IUnknown_QueryInterface(pComObject, IID_IOleWindow)
	
	IfEqual pOleWin,
		Return False
	
	; IOleWindow::GetWindow()
	iErr := DllCall(__VTable(pOleWin, 3), "UInt", pOleWin, "UInt*", hWnd) 

	If (__IsComError("IOleWindow::GetWindow", iErr))
		Return
	
	__IUnknown_Release(pOleWin)
	
	Return DllCall("GetParent", "UInt", hWnd)
}


; ..............................................................................
/****** ws4ahk/WS_GetComControlInHWND
* Description
*	Retrieves the COM control object associated with a COM control.
* Usage
*	WS_GetComControlInHWND(hWnd)
* Parameters
*	* hWnd - (Integer) The handle (i.e. ahk_id) of a COM control.
* Return Value
*	* Success: (Integer) A pointer to an IDispach interface of the COM object 
*	                     in the HWND (i.e. ahk_id).
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	
* Related
*	WS_GetHWNDofComControl, WS_AttachComControlToHWND
* Example
	; see WS_CreateComControlContainer example
******
*/
WS_GetComControlInHWND(hWnd)
{
	global IID_IDispatch
	
	iErr := DllCall("atl\AtlAxGetControl"
						, "UInt", hWnd
						, "UInt*", pUnknown
						, "Int")

	If (__IsComError("AtlAxGetControl", iErr))
		Return

	pDispatch := __IUnknown_QueryInterface(pUnknown, IID_IDispatch)
	
	__IUnknown_Release(pUnknown)
	
	Return pDispatch
}


; ..............................................................................
/****** ws4ahk/WS_AttachComControlToHWND
* Description
*	Attaches a COM control object to an existing COM control container.
* Usage
*	WS_AttachComControlToHWND(pObject, hWnd)
* Parameters
*	* pObject - (Integer) Pointer to a COM control object.
*	* hWnd - (Integer) Handle to a control to attach the COM object to.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	Attaches a COM control object to a COM control container created using
*	CreateComControlContainer() function.
* Related
*	WS_GetComControlInHWND, WS_GetHWNDofComControl
* Example
    #Include ws4ahk.ahk
    WS_Initialize()
    WS_InitComControls()
	
    Gui, +LastFound
    Gui, Show, w800 h600 Center, DhtmlEdit Test
    hWnd := WinExist()
    COMhWnd := WS_CreateComControlContainer(hWnd, 10, 10, 780, 580)
	ppvDEdit := WS_CreateObject("DhtmlEdit.DhtmlEdit")
	WS_AttachComControlToHWND(ppvDEdit, COMhWnd) 
    WS_AddObject(ppvDEdit, "DHtmlControl")
    ; the scripting environment has the object, so we don't need it anymore
    WS_ReleaseObject(ppvDEdit)
    WS_Exec("DHtmlControl.LoadUrl %s", "http://www.autohotkey.com")
    Return
    GuiClose:
        Gui, %A_Gui%:Destroy
        WS_UninitComControls()
        WS_Uninitialize()
        ExitApp
******
*/
WS_AttachComControlToHWND(pdsp, hWnd)
{
	iErr := DllCall("atl\AtlAxAttachControl"
					, "UInt", pdsp
					, "UInt", hWnd
					, "UInt", 0
					, "Int")
	
	If (__IsComError("AtlAxAttachControl", iErr))
		Return False
	
	Return True
}


; ..............................................................................
/****** ws4ahk/WS_CreateComControlContainer
* Description
*	Create a control on a window that will host a COM control object.
* Usage
*	WS_CreateComControlContainer(hWnd, X, Y, Width, Height [, sProgIDorClassID ] )
* Parameters
*	*  hWnd - (Integer) The handle (i.e. ahk_id) of the window to add the COM control.
*	*  X - (Integer) X position of the control.
*	*  Y - (Integer) Y position of the control.
*	*  Width - (Integer) Width of the control.
*	*  Height - (Integer) Height of the control.
*	*  sProgIDorClassID - (Optional) (String) 
*	                      The Program ID (e.g. "DhtmlEdit.DhtmlEdit") or
*	                      Class ID of the object to create in the control.
* Return Value
*	* Success: (Integer) Handle (i.e. ahk_id) of the new control.
*	* Failure: 0 (NULL).
* ErrorLevel
*	Set to the DllCall() result.
* Remarks
*	WS_InitComControls must be called before calling this function.
*
*	If a valid sProgIDorClassID is supplied, a COM control object will
*	automatically be created in the COM control container. Use 
*	WS_GetComControlInHWND to retrieve the associated COM object.
*	
*	If sProgIDorClassID is not supplied, then the COM control container will not 
*	initially have an associated COM control. Use WS_AttachComControlToHWND() to
*	add a COM control object to the COM control container.
* Related
*	WS_InitComControls, 
*	WS_AttachComControlToHWND, WS_GetComControlInHWND, WS_GetHWNDofComControl
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	WS_InitComControls()
	Gui, +LastFound
	Gui, Show, w800 h600 Center, DhtmlEdit Test
	hWnd := WinExist()
	COMhWnd := WS_CreateComControlContainer(hWnd, 10, 10, 780, 580, "DhtmlEdit.DhtmlEdit")
	ppvDEdit := WS_GetComControlInHWND(COMhWnd)
	WS_AddObject(ppvDEdit, "DHtmlControl")
	; the scripting environment has the object, so we don't need it anymore
	WS_ReleaseObject(ppvDEdit)
	WS_Exec("DHtmlControl.LoadUrl %s", "http://www.autohotkey.com")
	Return
	GuiClose:
		Gui, %A_Gui%:Destroy
		WS_UninitComControls()
		WS_Uninitialize()
		ExitApp
******
*/
WS_CreateComControlContainer(hWnd, x, y, w, h, sName = "")
{
	static AtlAxWin := "AtlAxWin"
	pName := sName ? &sName : 0
	; 0x10000000|  0x40000000  |  0x04000000 
	; WS_VISIBLE|WS_CHILDWINDOW|WS_CLIPSIBLINGS
	Return DllCall("CreateWindowEx"
					, "UInt", 0x200
					, "UInt", &AtlAxWin
					, "UInt", pName
					, "UInt", 0x10000000|0x40000000|0x04000000
					, "Int" , x
					, "Int" , y
					, "Int" , w
					, "Int" , h
					, "UInt", hWnd
					, "UInt", 0
					, "UInt", 0
					, "UInt", 0
					, "UInt")
}


/****ih* /Internal Functions ***************************************************
* About
*	Windows Scripting for Autohotkey (stdlib) v0.11 beta
*	
*	Requires Autohotkey v1.0.47 or above.
*	
*	You shouldn't need to worry about these functions unless
*	you know what they're for.
********************************************************************************
*/


; ..............................................................................
/****if* Internal Functions/WS_CreateObjectFromDll
* Description
*	Create COM object directly from a DLL or OCX file.
* Usage
*	WS_CreateObjectFromDll(sDllPath, sClassID [, sInterfaceID = IDispatch ] )
* Parameters
*	* sDllPath -- (String) Path to the dll/ocx containing the COM object.
*	* sClassID -- (String) Class ID of the object to create
*	                       (e.g. "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}").
*	* sInterfaceID -- (Optional) (String) Interface ID of the object to create
*	                                     (e.g. "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}").
* Return Value
*	* Success: (Integer) Pointer to the created object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	To create a COM object, it must be registered with the system. This 
*	function skips the standard approach and creates objects directly from the
*	DLL or OCX file. Since this involves a bit of hackery, it is not
*	guaranteed to work. This has been tested with msscript.ocx and seems to
*	work without any problems.
*	
*	This is a port of the code from the very clever Elias on The Code Project.
*	http://www.codeproject.com/com/Emul_CoCreateInstance.asp
* Related
*	WS_CreateObject, WS_ReleaseObject
* Example
	#Include ws4ahk.ahk
	WS_Initialize()
	; Create the Microsoft Scripting Control directly from the DLL
	pScriptCtrl := WS_CreateObjectFromDll("C:\Windows\system32\msscript.ocx"
	                                     ,"{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}"
	                                     ,"{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}")
	WS_AddObject(pScriptCtrl, "oScript")
	WS_Exec("oScript.Language = %s", "VBScript")
	WS_Exec("oScript.ExecuteStatement %s", "Msgbox ""Did this blow your mind?""")
******
*/
WS_CreateObjectFromDll(sDll, sClsId, sIId = "{00020400-0000-0000-C000-000000000046}")
{                                          ; ^ IDispatch          
	global IID_IDispatch
	
	If (__CLSIDFromString(sClsId, sbinClsId) And __IIDFromString(sIId, sbinIId))
		ppv := __CreateInstanceFromDll(sDll, sbinClsId, sbinIId)
	Else
		Return ; failed to convert class id or interface id

	; Did __CreateInstanceFromDll fail?
	IfEqual, ppv,
		Return
		
	If (sIId = IID_IDispatch)
		Return __GetIDispatch(ppv)
	Else
		Return ppv
}


; ..............................................................................
/****ix* Internal Functions/__CreateInstanceFromDll
* Description
*	Manually creates an object by directly accessing the DLL/OCX file.
* Usage
*	__CreateInstanceFromDll(sDllPath, ByRef sBinaryClassID, ByRef sBinaryIId)
* Parameters
*	* sDllPath - (String) Path to the dll or ocx file.
*	* sBinaryClassID - (ByRef) (String) String holding the binary version of the Class ID
*	* sBinaryIId - (ByRef) (String) String holding the binary version of the Interface ID
* Return Value
*	* Success: (Integer) Pointer to created object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	This involves a bit of hackery, but it usually seems to work.
*	This isn't the recommended way of creating objects, so use at your own risk.
*	
*	This code is based on the amazing work by Elias (aka lallous) on CodeProject.
*	
*	http://www.codeproject.com/com/Emul_CoCreateInstance.asp
*	
*	Note that there is no need to free the library explicitly.
*	It should be automatically freed when CoUninitialize is called.
*	
*	Used in WS_CreateObjectFromDll
* Related
*	
******
*/
__CreateInstanceFromDll(sDll, ByRef sbinClassId, ByRef sbinIId)
{
	static IID_IClassFactory := "{00000001-0000-0000-C000-000000000046}"
	If (!__IIDFromString(IID_IClassFactory, sbinIID_IClassFactory))
		Return
	
	__ANSI2Unicode(sDll, wsDll)
	hDll := DllCall("ole32\CoLoadLibrary", "Str", wsDll, "Int", 1, "UInt")
	
	If (ErrorLevel <> 0)
	{
		__ComError(ErrorLevel, "CoLoadLibrary: Error calling dll function: " ErrorLevel)
		Return
	}
	
	If (hDll = 0)
	{
		__ComError("", "CoLoadLibrary: Library could not be loaded.")
		Return
	}

	iErr := DllCall(sDll . "\DllGetClassObject"
					,"Str" , sbinClassId
					,"Str" , sbinIID_IClassFactory
					,"UInt*", pIFactory
					,"Int")

	If (__IsComError("DllGetClassObject", iErr))
		Return
	
	iObjPtr := __IClassFactory_CreateInstance(pIFactory, 0, sbinIId)
	
	__IUnknown_Release(pIFactory)
	
	Return iObjPtr
}


; ..............................................................................
/****ix* Internal Functions/__GetIDispatch
* Description
*	Try to query a COM object for the 'most useful' interface.
* Usage
*	__GetIDispatch(pIDispatch [, iLocaleID = Default ] )
* Parameters
*	* pIDispatch -- (Integer) Pointer to an IDispach interface of an object.
*	* iLocaleID -- (Optional) (Integer) The Locale to use.
* Return Value
*	* Success: (Integer) Pointer to the 'most useful' interface.
*	* Failure: Value of pIDispatch argument.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	I don't quite understand the purpose of doing this, but Sean was doing it, 
*	and the code found here 
*	http://svn.python.org/projects/ctypes/trunk/comtypes/comtypes/client/__init__.py
*	was doing it, so I guess this can do it too.
* Related
*	WS_CreateObject, WS_GetObject, WS_CreateObjectFromDll
******
*/
__GetIDispatch(ppObj, LCID = 0)
{
	iTypeInfoCount := __IDispatch_GetTypeInfoCount(ppObj)
	If (iTypeInfoCount = 0 Or iTypeInfoCount = "")
		Return ppObj

	ppTypeInfo := __IDispatch_GetTypeInfo(ppObj, LCID)
	IfEqual, ppTypeInfo,
		Return ppObj
	
	; find the interface marked as default
	pattr := __ITypeInfo_GetTypeAttr(ppTypeInfo)
	IfEqual, pattr,
	{
		__IUnknown_Release(ppTypeInfo)                                    
		Return ppObj
	}
	
	pdisp := __IUnknown_QueryInterface(ppObj, pattr)
	
	sErr := ErrorLevel ; save ErrorLevel
	
	If (!__ITypeInfo_ReleaseTypeAttr(ppTypeInfo, pattr))
		ErrorLevel := sErr . "`n" . ErrorLevel  ; add to the error
		
	__IUnknown_Release(ppTypeInfo)
	
	IfEqual, pdisp,
		Return ppObj
	Else
	{
		__IUnknown_Release(ppObj)
		Return pdisp
	}	
}



; ..............................................................................
/****ix* Internal Functions/__CLSIDFromProgID
* Description
*	Looks up the binary Class ID of a Program ID.
* Usage
*	__CLSIDFromProgID(sProgID, ByRef BinaryClassID)
* Parameters
*	* sProgID -- (String) A Program ID (e.g. "Excel.Application") 
*	* BinaryClassID -- (ByRef) Variable to receive the binary version of the Class ID.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	The binary ClassId is stored in a string that must be passed ByRef because  
*	returning AHK strings 'by value' that contain binary data will be truncated
*	to the first 0x00 binary value.
* Related
*	__CLSIDFromString, __IIDFromString
******
*/
__CLSIDFromProgID(sProgId, ByRef sbinClassId)
{
	__ANSI2Unicode(sProgId, wsProgId)
	VarSetCapacity(sbinClassId, 16) ; 16 = sizeof(CLSID) 
	iErr := DllCall("ole32\CLSIDFromProgID"
					, "Str", wsProgId
					, "Str", sbinClassId
					, "Int")
					
	If (__IsComError("CLSIDFromProgID", iErr))
		Return False
	
	Return True
}


; ..............................................................................
/****ix* Internal Functions/__CLSIDFromString
* Description
*	Converts a string Class ID to a binary Class ID.
* Usage
*	__CLSIDFromString(sClassID, ByRef BinaryClassID)
* Parameters
*	* sClassID -- (String) A Class ID (e.g. "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}") 
*	* BinaryClassID -- (ByRef) Variable to receive the binary version of the Class ID.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	The binary ClassId is stored in a string that must be passed ByRef because  
*	returning AHK strings 'by value' that contain binary data will be truncated
*	to the first 0x00 binary value.
* Related
*	__CLSIDFromProgID, __IIDFromString
******
*/
__CLSIDFromString(sClassId, ByRef sbinClassId)
{
	__ANSI2Unicode(sClassId, wsClassId)
	VarSetCapacity(sbinClassId, 16) ; 16 = sizeof(CLSID) 
	iErr := DllCall("ole32\CLSIDFromString"
					, "Str", wsClassId
					, "Str", sbinClassId
					, "Int")

	If (__IsComError("CLSIDFromString", iErr))
		Return False
		
	Return True
}


; ..............................................................................
/****ix* Internal Functions/__IIDFromString
* Description
*	Converts a string Interface ID to a binary Interface ID.
* Usage
*	__IIDFromString(sIId, ByRef BinaryIId)
* Parameters
*	* sIId -- (String) An Interface ID (e.g. "{00000000-0000-0000-C000-000000000046}") 
*	* BinaryIId -- (ByRef) Variable to receive the binary version of the Interface ID.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	The binary Interface ID is stored in a string that must be passed ByRef   
*	because returning AHK strings 'by value' that contain binary data will be 
*	tuncated to the first 0x00 binary value.
*
*	Win API function IIDFromString() seems to be identical to Win API function 
*	CLSIDFromString(). I really don't see why there are two separate functions
*	to do this.
* Related
*	__CLSIDFromProgID, __CLSIDFromString
******
*/
__IIDFromString(sIId, ByRef sbinIId)
{
	__ANSI2Unicode(sIId, wsIId)
	VarSetCapacity(sbinIId, 16) ; 16 = sizeof(IID) 
	iErr := DllCall("ole32\IIDFromString"
					, "Str", wsIId
					, "Str", sbinIId
					, "Int")

	If (__IsComError("IIDFromString", iErr))
		Return False
	
	Return True
}


; ..............................................................................
/****ix* Internal Functions/__IsComError
* Description
*	Checks for error and sets ErrorLevel.
* Usage
*	__IsComError(sFunctionName, iHRESULT)
* Parameters
*	* sFunctionName -- (String) Name of the function for error description.
*	* iHRESULT -- (Integer) HRESULT value from a COM related Dll call.
* Return Value
*	(Boolean) True if there is an error. 
*	False if HRESULT indicates success (HRESULT >= 0).
* ErrorLevel
*	0 if previous DllCall was successful and HRESULT = 0 (S_OK).
*	Error description if was an error, or a success message.
* Remarks
*	Should be called right after a DllCall(). Checks ErrorLevel for DllCall()
*	error, then checks HRESULT for success or failure.
* Related
*	__ComError
******
*/
__IsComError(sFunction, iErr)
{
	If (ErrorLevel <> 0) ; error calling the function
	{
		__ComError(ErrorLevel, sFunction ": DllCall error " ErrorLevel)
		Return True
	}
	If (iErr = 0) ; S_OK
	{
		ErrorLevel := 0
		Return False
	}
	Else If ((iErr & 0x80000000) > 0) ; IS_ERROR()
	{
		__ComError(iErr, sFunction ": error " iErr)
		Return True
	}
	Else ; SUCCEEDED(), but not S_OK
	{
		__ComError(iErr, sFunction ": succeeded with result " iErr)
		Return False
	}
}

; ..............................................................................
/****ix* Internal Functions/__ComError
* Description
*	Sets ErrorLevel with an error.
* Usage
*	__ComError(?Error, sDescription)
* Parameters
*	* ?Error -- (Integer|"") Number of the error, or ""
*	* sDescription -- (String) Description of the error.
* Return Value
*	None.
* ErrorLevel
*	Set with a formatted error message.
* Remarks
*	
* Related
*	__IsComError
******
*/
__ComError(iErr, sErrDesc) 
{
	If (iErr = "")
		ErrorLevel := sErrDesc
	Else
		ErrorLevel := "[" iErr "] " sErrDesc
}


; ..............................................................................
/****ix* Internal Functions/__ANSI2Unicode
* Description
*	Converts an ANSI string to its UTF16 equivalent.
* Usage
*	__ANSI2Unicode(sAnsiString, ByRef Utf16String)
* Parameters
*	* sAnsiString -- (String) ANSI string to convert.
*	* Utf16String -- (ByRef) Variable to get the UTF16 string.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	DllCall() result.
* Remarks
*	Returned string must be ByRef because passing AHK strings 'by value' that 
*	contain binary data will be truncated to the first 0x00 binary value. 
* Related
*	__Unicode2ANSI
******
*/
__ANSI2Unicode(sAnsi, ByRef sUtf16)
{
	iSize := DllCall("MultiByteToWideChar"
   					, "UInt", 0  ; from CP_ACP (ANSI)
					, "UInt", 0  ; no flags
					, "UInt" , &sAnsi
					, "Int" , -1 ; until NULL
					, "UInt", 0  ; NULL
					, "Int" , 0)
					
	If (ErrorLevel <> 0)
		Return False
		
	If (iSize < 1)
		Return False
	
	VarSetCapacity(sUtf16, (iSize+1) * 2, 0)
   
	iSize := DllCall("MultiByteToWideChar"
   					, "UInt", 0  ; from CP_ACP (ANSI)
					, "UInt", 0  ; no flags
					, "UInt" , &sAnsi
					, "Int" , -1 ; until NULL
					, "UInt" , &sUtf16
					, "Int" , iSize)
					
	If (ErrorLevel <> 0)
		Return False
		
	If (iSize < 1)
		Return False
	Else
		Return True
}


; ..............................................................................
/****ix* Internal Functions/__Unicode2ANSI
* Description
*	Converts a UTF16 string to its ANSI equivalent.
* Usage
*	__Unicode2ANSI(psUtf16, ByRef Ansi)
* Parameters
*	* psUtf16 -- (Integer) Pointer to a UTF16 string.
*	* Ansi -- (ByRef) Variable to receive the ANSI string.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	Not changed.
* Remarks
*	
* Related
*	__ANSI2Unicode
******
*/
__Unicode2ANSI(psUtf16, ByRef sAnsi)
{
	If (psUtf16 = 0)
		Return False
	
	ErrLvl := ErrorLevel ; save ErrorLevel
	
	iSize := DllCall("WideCharToMultiByte"
					, "UInt", 0  ; to CP_API (ANSI)
					, "UInt", 0  ; no flags
					, "UInt", psUtf16
					, "Int", -1  ; until NULL
					, "UInt", 0  ; NULL
					, "Int",  0  ; Just find length
					, "UInt", 0  ; NULL
					, "UInt", 0) ; NULL
					
	If (ErrorLevel <> 0 Or iSize < 1)
	{
		ErrorLevel := ErrLvl ; restore ErrorLevel
		Return False
	}
					
	VarSetCapacity(sAnsi, iSize+1)
	iSize := DllCall("WideCharToMultiByte"
					, "UInt", 0  ; to CP_API (ANSI)
					, "UInt", 0  ; no flags
					, "UInt", psUtf16
					, "Int", -1  ; until NULL
					, "Str", sAnsi
					, "Int", iSize
					, "UInt", 0  ; NULL
					, "UInt", 0) ; NULL
					
	If (ErrorLevel <> 0 Or iSize < 1)
	{
		ErrorLevel := ErrLvl ; restore ErrorLevel
		Return False
	}
	Else
	{
		ErrorLevel := ErrLvl ; restore ErrorLevel
		Return True
	}
}


; ..............................................................................
/****ix* Internal Functions/__VTable
* Description
*	Get pointer to the function at the specified vtable index.
* Usage
*	__VTable(pVTable, iIndex)
* Parameters
*	* pVTable -- (Integer) Pointer to the object's vtable.
*	* iIndex -- (Integer) Index of the function pointer to retrieve.
* Return Value
*	(Integer) Pointer to the desired function.
* ErrorLevel
*	Not changed.
* Remarks
*	
* Related
*	
******
*/
__VTable(ppv, idx)
{
	Return NumGet(NumGet(ppv+0) + 4*idx)
}


; ..............................................................................
/****ix* Internal Functions/__StringToBSTR
* Description
*	Converts a string to a BSTR.
* Usage
*	__StringToBSTR(sAnsi)
* Parameters
*	* sAnsi -- (String) ANSI string to turn into a BSTR.
* Return Value
*	* Success: Pointer to the BSTR containing the string.
*	* DllCall failure: None ("").
*	* SysAllocString failure: 0 (NULL).
* ErrorLevel
*	DllCall() result.
* Remarks
*	Converts a normal ANSI string to Unicode, then creates a BSTR with it.
*	The resulting BSTR should be freed with the __FreeBSTR function.
* Related
*	__FreeBSTR
******
*/
__StringToBSTR(sAnsi)
{
	__ANSI2Unicode(sAnsi, sUnicode)
	Return DllCall("oleaut32\SysAllocString", "Str", sUnicode, "UInt")
}


; ..............................................................................
/****ix* Internal Functions/__FreeBSTR
* Description
*	Frees a BSTR.
* Usage
*	__FreeBSTR(pBSTR)
* Parameters
*	* pBSTR -- (Integer) Pointer to a BSTR.
* Return Value
*	DllCall() result.
* ErrorLevel
*	Not changed.
* Remarks
*	
* Related
*	__StringToBSTR
******
*/
__FreeBSTR(iBstrPtr)
{
	ErrLvl := ErrorLevel ; save ErrorLevel
	DllCall("oleaut32\SysFreeString", "UInt", iBstrPtr)
	Ret := ErrorLevel ; get the DllCall() error result
	ErrorLevel := ErrLvl ; restore ErrorLevel
	Return Ret ; return DllCall() error result
}


; ..............................................................................
/****ix* Internal Functions/__UnpackVARIANT
* Description
*	Converts a VARIANT structure to a normal AHK variable.
* Usage
*	__UnpackVARIANT(ByRef sVARIANT, ByRef RetVal)
* Parameters
*	* sVARIANT -- (ByRef) (String) String containing a VARIANT structure.
*	* RetVal -- (ByRef) Variable to receive the unpacked value.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	Not changed.
* Remarks
*	Most return types are handled. Unhandled types are:
*	* Array
*	* Currency
*	* Date
*	* VARIANT*
*	* DECIMAL*
* Related
*	
******
*/
__UnpackVARIANT(ByRef VARIANT, ByRef xReturn)
{
	static VT_BYREF := 0x4000
	vt := NumGet(VARIANT, 0, "UShort")
	pdata := &VARIANT + 8
	
	; VT_BSTR
	If (vt = 8)
	{
		__Unicode2ANSI(NumGet(pdata+0), xReturn)
		__VariantClear(VARIANT)
		Return True
	}
	Else If (vt = 8|VT_BYREF)
	{
		__Unicode2ANSI(NumGet(NumGet(pdata+0)), xReturn)
		__VariantClear(VARIANT)
		Return True
	}
	; VT_EMPTY
	Else If (vt = 0)
	{
		xReturn := ""
		Return True
	}
	; VT_UI1
	Else If (vt = 17)
	{
		xReturn := NumGet(pdata+0, 0,"UChar")
		Return True
	}
	Else If (vt = 17|VT_BYREF)
	{
		xReturn := NumGet(NumGet(pdata+0), 0,"UChar")
		Return True
	}
	; VT_I2
	Else If (vt = 2)
	{
		xReturn := NumGet(pdata+0, 0, "Short")
		Return True
	}
	Else If (vt = 2|VT_BYREF)
	{
		xReturn := NumGet(NumGet(pdata+0), 0,"Short")
		Return True
	}
	; VT_I4
	Else If (vt = 3)
	{
		xReturn := NumGet(pdata+0, 0,"Int")
		Return True
	}
	Else If (vt = 3|VT_BYREF)
	{
		xReturn := NumGet(NumGet(pdata+0), 0,"Int")
		Return True
	}
	; VT_R4
	Else If (vt = 4)
	{
		xReturn := NumGet(pdata+0, 0,"Float")
		Return True
	}
	Else If (vt = 4|VT_BYREF)
	{
		xReturn := NumGet(NumGet(pdata+0), 0,"Float")
		Return True
	}
	; VT_R8
	Else If (vt = 5)
	{
		xReturn := NumGet(pdata+0, 0,"Double")
		Return True
	}
	Else If (vt = 5|VT_BYREF)
	{
		xReturn := NumGet(NumGet(pdata+0), 0,"Double")
		Return True
	}
	; VT_BOOL
	Else If (vt = 11)
	{
		xVal := NumGet(pdata+0, 0,"Short")
		xReturn := -xVal ; fix -1 = true
		Return True
	}
	Else If (vt = 11|VT_BYREF)
	{
		xVal := NumGet(NumGet(pdata+0), 0,"Short")
		xReturn := -xVal ; fix -1 = true
		Return True
	}
	; VT_ERROR
	Else If (vt = 10)
	{
		xReturn := NumGet(pdata+0, 0,"UInt")
		Return True
	}
	Else If (vt = 10|VT_BYREF)
	{
		xReturn := NumGet(NumGet(pdata+0), 0,"UInt")
		Return True
	}
	; VT_DISPATCH or VT_UNKNOWN
	Else If ((vt = 9) Or (vt = 13))
	{
		xVal := NumGet(pdata+0, 0,"UInt")
		If (xVal = 0)
			xReturn := ""
		Else
			xReturn := xVal
		Return True
	}
	Else If ((vt = 9|VT_BYREF) Or (vt = 13|VT_BYREF))
	{
		xVal := NumGet(NumGet(pdata+0), 0,"UInt")
		If (xVal = 0)
			xReturn := ""
		Else
			xReturn := xVal
		Return True
	}
	
	; Unhandled VARIANT types:
	; Array, Currency, Date, VARIANT*, and DECIMAL*
	__VariantClear(VARIANT)
	Return False
}


; ..............................................................................
/****ix* Internal Functions/__VariantClear
* Description
*	Releases references and clears the contents of a VARIANT structure.
* Usage
*	__VariantClear(ByRef sVARIANT)
* Parameters
*	* sVARIANT -- (ByRef) (String) String containing a VARIANT structure.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	
* Related
*	
******
*/
__VariantClear(ByRef VAR)
{
	iErr := DllCall("oleaut32\VariantClear", "Str", VAR, "Int")
	
	If (__IsComError("VariantClear", iErr))
		Return False

	Return True
}



; ## IScriptControl ############################################################
/****ih* Interfaces/IScriptControl
* VTable
*	 0   call_QueryInterface    - Returns a pointer to a specified interface on an object to which a client currently holds an interface pointer 
*	 1   call_AddRef            - Increments the reference count for an interface on an object
*	 2   call_Release           - Decrements the reference count for the calling interface on a object
*	 3   call_GetTypeInfoCount  - Retrieves the number of type information interfaces that an object provides (either 0 or 1)
*	 4   call_GetTypeInfo       - Retrieves the type information for an object
*	 5   call_GetIDsOfNames     - Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs
*	 6   call_Invoke            - Provides access to properties and methods exposed by an object.
*	 7 * get_Language           - Language engine to use
*	 8 * put_Language           - Language engine to use
*	 9   get_State              - State of the control
*	10   put_State              - State of the control
*	11 * put_SitehWnd           - hWnd used as a parent for displaying UI
*	12 * get_SitehWnd           - hWnd used as a parent for displaying UI
*	13   get_Timeout            - Length of time in milliseconds that a script can execute before being considered hung
*	14   put_Timeout            - Length of time in milliseconds that a script can execute before being considered hung
*	15 * get_AllowUI            - Enable or disable display of the UI
*	16 * put_AllowUI            - Enable or disable display of the UI
*	17   get_UseSafeSubset      - Force script to execute in safe mode and disallow potentially harmful actions
*	18   put_UseSafeSubset      - Force script to execute in safe mode and disallow potentially harmful actions
*	19   get_Modules            - Collection of modules for the ScriptControl
*	20 * get_Error              - The last error reported by the scripting engine
*	21   get_CodeObject         - Object exposed by the scripting engine that contains methods and properties defined in the code added to the global module
*	22   get_Procedures         - Collection of procedures that are defined in the global module
*	23   call__AboutBox         - 
*	24 * call_AddObject         - Add an object to the global namespace of the scripting engine
*	25   call_Reset             - Reset the scripting engine to a newly created state
*	26   call_AddCode           - Add code to the global module
*	27 * call_Eval              - Evaluate an expression within the context of the global module
*	28 * call_ExecuteStatement  - Execute a statement within the context of the global module
*	29   call_Run               - Call a procedure defined in the global module
*	(only the * members are implemented)
******
*/


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_Language
* Description
*	Gets/Sets the language engine to use.
* Usage
*	__IScriptControl_Language(pIScriptControl [, sLanguage] )
* Parameters
*	*  pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
*	*  sLanguage -- (Optional) (String) Scripting language to use 
*	                                    ("VBScript" or "JScript")
* Return Value
*	If Language argument is supplied:
*	* (Boolean) True on success, False on failure.
*	If Language argument is NOT supplied:
*	* Success: (String) current language being used.
*	* Failure: 0.
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	Changing the scripting language seems to reset the environment.
*
*	If language has not been set, returns "".
*
*	It is possible that languages other than VBScript and JScript could be
*	set/returned. These alternate scripting languages would have be properly
*	registered with the system (but I've never seen any language that can
*	do that).
* Related
*	IScriptControl
******
*/
__IScriptControl_Language(ppvScriptControl, sLanguage="`b")
{
	If (sLanguage = "`b")
	{	; Get Language
		iErr := DllCall(__VTable(ppvScriptControl, 7), "UInt", ppvScriptControl
					, "UInt*", ibstrLang
					, "Int")
		
		If (__IsComError("IScriptControl::Language get", iErr))
			Return 0
			
		If (!__Unicode2ANSI(ibstrLang, sLanguage))
			sLanguage := 0 ; failed to change to ANSI
			
		__FreeBSTR(ibstrLang)
		
		Return sLanguage
	}
	Else
	{	; Put Language
		bstrLang := __StringToBSTR(sLanguage)
		iErr := DllCall(__VTable(ppvScriptControl, 8), "UInt", ppvScriptControl
					, "UInt", bstrLang
					, "Int")
					
		If (__IsComError("IScriptControl::Language set", iErr))
			blnIsSuccess := False
		Else
			blnIsSuccess := True

		__FreeBSTR(bstrLang)
		
		Return blnIsSuccess
	}
}


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_SitehWnd
* Description
*	Get/set hWnd used as a parent for displaying UI.
* Usage
*	__IScriptControl_SitehWnd(pIScriptControl [, iWindowHandle] )
* Parameters
*	*  pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
*	*  iWindowHandle -- (Optional) (Integer) The hWnd (i.e. ahk_id) to be used
*	                                         as the parent for displaying UI.
* Return Value
*	If WindowHandle argument is supplied:
*	* (Boolean) True on success, False on failure.
*	If WindowHandle argument is NOT supplied:
*	* Success: (Integer) Current hWnd being used.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description.
*	* Failure: error description.
* Remarks
*	
* Related
*	IScriptControl
******
*/
__IScriptControl_SitehWnd(ppvScriptControl, iWindowHandle="`b")
{
	If (iWindowHandle = "`b")
	{	; Get SitehWnd
		iErr := DllCall(__VTable(ppvScriptControl, 12), "UInt", ppvScriptControl
						, "UInt*", iWindowHandle
						, "Int")
						
		If (__IsComError("IScriptControl::SitehWnd get", iErr))
			Return
		
		Return iWindowHandle
	}
	Else
	{	; Put SitehWnd
		iErr := DllCall(__VTable(ppvScriptControl, 11), "UInt", ppvScriptControl
						, "UInt", iWindowHandle
						, "Int")
						
		If (__IsComError("IScriptControl::SitehWnd put", iErr))
			Return False
		
		Return True
	}
}


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_AllowUI
* Description
*	Gets/sets if the display of the UI is enabled or disabled.
* Usage
*	__IScriptControl_AllowUI(pIScriptControl [, blnAllow] )
* Parameters
*	*  pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
*	*  Allow -- (Optional) (Boolean) Sets if the display of the UI is enabled or disabled.
* Return Value
*	If blnAllow argument is provided:
*	* (Boolean) True on success, False on failure.
*	If blnAllow argument is NOT provided:
*	* Success: (Boolean) If the display of the UI is enabled.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*
* Related
*	IScriptControl
******
*/
__IScriptControl_AllowUI(ppvScriptControl, iAllow="`b")
{
	If (iAllow = "`b")
	{	; Get AllowUI
		iErr := DllCall(__VTable(ppvScriptControl, 15), "UInt", ppvScriptControl
						, "Short*", iAllow
						, "Int")
		If (__IsComError("IScriptControl::AllowUI get", iErr))
			Return
		
		Return -iAllow ; negitive fixes the COM 'True' to normal bool convention
	}
	Else
	{   ; Put AllowUI
		iErr := DllCall(__VTable(ppvScriptControl, 16), "UInt", ppvScriptControl
						, "Short", iAllow
						, "Int")
		If (__IsComError("IScriptControl::AllowUI put", iErr))
			Return False
		
		Return True
	}
}


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_Error
* Description
*	Gets the error object used by the scripting engine.
* Usage
*	__IScriptControl_Error(pIScriptControl)
* Parameters
*	*  pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
* Return Value
*	* Success: (Integer) Pointer to the IScriptError interface of a ScriptError object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*
* Related
*	IScriptControl
******
*/
__IScriptControl_Error(ppvScriptControl)
{
	iErr := DllCall(__VTable(ppvScriptControl, 20), "UInt", ppvScriptControl
					, "UInt*", ppvScriptError
					, "Int")
					
	If (__IsComError("IScriptControl::Error", iErr))
		Return
	
	Return ppvScriptError
}


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_AddObject
* Description
*	Add an object to the global namespace of the scripting engine.
* Usage
*	__IScriptControl_AddObject(pIScriptControl, sObjName, pIObjDispatch, blnAddMembers)
* Parameters
*	* pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
*	* sObjName -- (String) Name that the object will have in the scripting environment.
*	* pIObjDispatch -- (Integer) Pointer to an IDispach interface of the object to add.
*	* blnAddMembers -- (Boolean) Make the members of the object global.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	If blnAddMembers is True, all members of the object become publicly
*	accessable identifiers.
* Related
*	IScriptControl
******
*/
__IScriptControl_AddObject(ppvScriptControl, sName, pObjectDispatch, blnAddMembers)
{
	bstrName := __StringToBSTR(sName)
	iErr := DllCall(__VTable(ppvScriptControl, 24), "UInt", ppvScriptControl
				, "UInt", bstrName
				, "UInt", pObjectDispatch
				, "Short", blnAddMembers
				, "Int")

	If (__IsComError("IScriptControl::AddObject", iErr))
		blnIsSuccess := False
	Else
		blnIsSuccess := True
	
	__FreeBSTR(bstrName)
	
	Return blnIsSuccess
}


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_Eval
* Description
*	Evaluate an expression within the context of the global module.
* Usage
*	__IScriptControl_Eval(pIScriptControl, sExpression, ByRef VarRet)
* Parameters
*	* pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
*	* sExpression -- (String) Scripting code to evaluate.
*	* VarRet -- (ByRef) The variable to receive the VARIANT structure
*	                    returned from the evaluation.
* Return Value
*	* Success: The HRESULT of the Eval() call.
*	* Failure: None (""), only if DllCall() fails.
* ErrorLevel
*	* Success: 0. 
*	* Failure: error description with DllCall() error number.
* Remarks
*	On success, VarRet will be set with the evaluation result.
*	
*	On failure, it will be a 16 byte empty string.
*
*	Like __IScriptControl_ExecuteStatement, this function does not process the
*	HRESULT. The HRESULT is returned so that further handling can be done.
* Related
*	IScriptControl, __IScriptControl_ExecuteStatement
******
*/
__IScriptControl_Eval(ppvScriptControl, sExpression, ByRef VarRet)
{
	bstrExpression := __StringToBSTR(sExpression)
	
	; Initialize the VARIANT structure to return
	VarSetCapacity(VarRet, 16) ; sizeof(VARIANT) = 16
	DllCall("oleaut32\VariantInit", "Str", VarRet)
	
	iErr := DllCall(__VTable(ppvScriptControl, 27), "UInt", ppvScriptControl
				, "UInt", bstrExpression
				, "Str" , VarRet
				, "Int")
				
	iErrLvl := ErrorLevel ; save ErrorLevel
	
	__FreeBSTR(bstrExpression)
	
	If (iErrLvl <> 0)
	{
		__ComError(iErrLvl, "IScriptControl::Eval: DllCall error " iErrLvl)
		Return
	}
	Else
	{
		; Return the error code to check for exceptions
		Return iErr
	}
}


; ..............................................................................
/****iI* IScriptControl/__IScriptControl_ExecuteStatement
* Description
*	Execute a statement within the context of the global module.
* Usage
*	__IScriptControl_ExecuteStatement(pIScriptControl, sCode)
* Parameters
*	* pIScriptControl -- (Integer) Pointer to an IScriptControl interface of an object.
*	* sCode -- (String) Scripting code to execute.
* Return Value
*	* Success: The HRESULT of the ExecuteStatement() call.
*	* Failure: None (""), only if DllCall() fails.
* ErrorLevel
*	* Success: 0. 
*	* Failure: error description with DllCall() error number.
* Remarks
*	Like __IScriptControl_Eval, this function does not process the HRESULT.
*	The HRESULT is returned so that further handling can be done.
* Related
*	IScriptControl, __IScriptControl_Eval
******
*/
__IScriptControl_ExecuteStatement(ppvScriptControl, sStatement)
{
	bstrStatement := __StringToBSTR(sStatement)
	iErr := DllCall(__VTable(ppvScriptControl, 28), "UInt", ppvScriptControl
				, "UInt", bstrStatement
				, "Int")
				
	iErrLvl := ErrorLevel ; save ErrorLevel
	
	__FreeBSTR(bstrStatement)
	
	If (iErrLvl <> 0)
	{
		__ComError(iErrLvl, "IScriptControl::ExecuteStatement: DllCall error " iErrLvl)
		Return
	}
	Else
	{
		; Return the error code to check for exceptions
		Return iErr
	}
}

; ## IScriptError ##############################################################
/****ih* Interfaces/IScriptError
* VTable
*	 0   call_QueryInterface    - Returns a pointer to a specified interface on an object to which a client currently holds an interface pointer
*	 1   call_AddRef            - Increments the reference count for an interface on an object
*	 2   call_Release           - Decrements the reference count for the calling interface on a object
*	 3   call_GetTypeInfoCount  - Retrieves the number of type information interfaces that an object provides (either 0 or 1)
*	 4   call_GetTypeInfo       - Retrieves the type information for an object
*	 5   call_GetIDsOfNames     - Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs
*	 6   call_Invoke            - Provides access to properties and methods exposed by an object.
*	 7 * get_Number             - Error number
*	 8   get_Source             - Source of the error
*	 9 * get_Description        - Friendly description of error
*	10   get_HelpFile           - File in which help for the error can be found
*	11   get_HelpContext        - Context ID for the topic with information on the error
*	12   get_Text               - Line of source code on which the error occurred
*	13   get_Line               - Source code line number where the error occurred
*	14   get_Column             - Source code column position where the error occurred
*	15 * call_Clear             - Clear the script error
*	(only the * members are implemented)
******
*/

; ..............................................................................
/****iI* IScriptError/__IScriptError_Number
* Description
*	Get the last error number.
* Usage
*	__IScriptError_Number(pIScriptError)
* Parameters
*	*  pIScriptError -- (Integer) Pointer to an IScriptError interface of an object.
* Return Value
*	* Success: (Integer) Error number.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*
* Related
*	IScriptError
******
*/
__IScriptError_Number(ppvScriptError)
{
	iErr := DllCall(__VTable(ppvScriptError, 7), "UInt", ppvScriptError
				, "Int*", iNumber
				, "Int")
	
	If (__IsComError("IScriptError::Number", iErr))
		Return
	
	Return iNumber
}


; ..............................................................................
/****iI* IScriptError/__IScriptError_Description
* Description
*	Get a friendly description of the last error.
* Usage
*	__IScriptError_Description(pIScriptError)
* Parameters
*	*  pIScriptError -- (Integer) Pointer to an IScriptError interface of an object.
* Return Value
*	* Success: (String) Error description, or "" if there is no error description.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*
* Related
*	IScriptError
******
*/
__IScriptError_Description(ppvScriptError)
{
	iErr := DllCall(__VTable(ppvScriptError, 9), "UInt", ppvScriptError
				, "UInt*", bstrDescription
				, "Int")
				
	If (__IsComError("IScriptError::Description", iErr))
		Return

	; The BSTR pointer is sometimes NULL
	If (bstrDescription = 0)
	{
		Return
	}
	Else
	{
		__Unicode2ANSI(bstrDescription, sAnsi)
		__FreeBSTR(bstrDescription)
		Return sAnsi
	}
}


; ..............................................................................
/****iI* IScriptError/__IScriptError_Clear
* Description
*	Clear the script error.
* Usage
*	__IScriptError_Clear(pIScriptError)
* Parameters
*	*  pIScriptError -- (Integer) Pointer to an IScriptError interface of an object.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*
* Related
*	IScriptError
******
*/
__IScriptError_Clear(ppvScriptError)
{
	iErr := DllCall(__VTable(ppvScriptError, 15), "UInt", ppvScriptError
				, "Int")
				
	If (__IsComError("IScriptError::Clear", iErr))
		Return False
	
	Return True
}

; ## IClassFactory #############################################################
/****ih* Interfaces/IClassFactory
* VTable
*	0   call_QueryInterface    Returns a pointer to a specified interface on an object to which a client currently holds an interface pointer
*	1   call_AddRef            Increments the reference count for an interface on an object
*	2   call_Release           Decrements the reference count for the calling interface on a object
*	3 * call_CreateInstance    Creates an uninitialized object.
*	4   call_LockServer        Locks object application open in memory.
*	(only the * members are implemented)
******
*/


; ..............................................................................
/****iI* IClassFactory/__IClassFactory_CreateInstance
* Description
*	Creates an uninitialized object.
* Usage
*	__IClassFactory_CreateInstance(pIClassFactory)
* Parameters
*	*  pIClassFactory -- (Integer) Pointer to an IClassFactory interface of an object.
* Return Value
*	* Success: (Integer) Pointer to new object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	Used in __CreateInstanceFromDll() function.
* Related
*	IClassFactory
******
*/
__IClassFactory_CreateInstance(ppvIClassFactory, pUnkOuter, ByRef riid)
{
	iErr := DllCall(__VTable(ppvIClassFactory, 3), "UInt", ppvIClassFactory
					, "UInt",  pUnkOuter
					, "Str",   riid
					, "Uint*", ppvObject
					, "Int")
	
	If (__IsComError("IClassFactory::CreateInstance", iErr))
		Return
	
	Return ppvObject
}

; ## ITypeInfo #################################################################
/****ih* Interfaces/ITypeInfo
* VTable
*	 0   call_QueryInterface        Returns pointers to supported interfaces.
*	 1   call_AddRef                Increments reference count.
*	 2   call_Release               Decrements reference count.
*	 3 * call_GetTypeAttr           Retrieves a TYPEATTR structure that contains the attributes of the type description.
*	 4   call_GetTypeComp           Retrieves the ITypeComp interface for the type description, which enables a client compiler to bind to the type description's members.
*	 5   call_GetFuncDesc           Retrieves the FUNCDESC structure that contains information about a specified function.
*	 6   call_GetVarDesc            Retrieves a VARDESC structure that describes the specified variable.
*	 7   call_GetNames              Retrieves the variable with the specified member ID (or the name of the property or method and its parameters) that correspond to the specified function ID.
*	 8   call_GetRefTypeOfImplType  If a type description describes a COM class, it retrieves the type description of the implemented interface types. For an interface, GetRefTypeOfImplType returns the type information for inherited interfaces, if any exist.
*	 9   call_GetImplTypeFlags      Retrieves the IMPLTYPEFLAGS enumeration for one implemented interface or base interface in a type description.
*	10   call_GetIDsOfNames         Maps between member names and member IDs, and parameter names and parameter IDs.
*	11   call_Invoke                Invokes a method, or accesses a property of an object, that implements the interface described by the type description.
*	12   call_GetDocumentation      Retrieves the documentation string, the complete Help file name and path, and the context ID for the Help topic for a specified type description.
*	13   call_GetDllEntry           Retrieves a description or specification of an entry point for a function in a DLL.
*	14   call_GetRefTypeInfo        If a type description references other type descriptions, it retrieves the referenced type descriptions.
*	15   call_AddressOfMember       Retrieves the addresses of static functions or variables, such as those defined in a DLL.
*	16   call_CreateInstance        Creates a new instance of a type that describes a component object class (coclass).
*	17   call_GetMops               Retrieves marshaling information.
*	18   call_GetContainingTypeLib  Retrieves the containing type library and the index of the type description within that type library.
*	19 * call_ReleaseTypeAttr       Releases a TYPEATTR previously returned by GetTypeAttr.
*	20   call_ReleaseFuncDesc       Releases a FUNCDESC previously returned by GetFuncDesc.
*	21   call_ReleaseVarDesc        Releases a VARDESC previously returned by GetVarDesc.
*	(only the * members are implemented)
******
*/


; ..............................................................................
/****iI* ITypeInfo/__ITypeInfo_GetTypeAttr
* Description
*	Retrieves a TYPEATTR structure that contains the attributes of the type description.
* Usage
*	__ITypeInfo_GetTypeAttr(pITypeInfo)
* Parameters
*	*  pITypeInfo -- (Integer) Pointer to an ITypeInfo interface of an object.
* Return Value
*	* Success: (Integer) Pointer to a TYPEATTR structure.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	TYPEATTR pointer should be freed via a call to __ITypeInfo_ReleaseTypeAttr.
*
*	Used by __GetIDispatch
* Related
*	ITypeInfo
******
*/
__ITypeInfo_GetTypeAttr(ppTypeInfo) 
{
	iErr := DllCall(__VTable(ppTypeInfo, 3), "UInt", ppTypeInfo
					, "UInt*", pTypeAttr
					, "Int")
					
	If (__IsComError("ITypeInfo::GetTypeAttr", iErr))
		Return
	
	Return pTypeAttr
}


; ..............................................................................
/****iI* ITypeInfo/__ITypeInfo_ReleaseTypeAttr
* Description
*	Releases a TYPEATTR previously returned by __ITypeInfo_GetTypeAttr.
* Usage
*	__ITypeInfo_ReleaseTypeAttr(pITypeInfo, pTypeAttr)
* Parameters
*	*  pITypeInfo -- (Integer) Pointer to an ITypeInfo interface of an object.
*	*  pTypeAttr -- (Integer) Pointer to a TYPEATTR structure.
* Return Value
*	(Boolean) True on success, False on failure.
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	Used by __GetIDispatch
* Related
*	ITypeInfo
******
*/
__ITypeInfo_ReleaseTypeAttr(ppTypeInfo, pTypeAttr)
{
	iErr := DllCall(__VTable(ppTypeInfo, 19), "UInt", ppTypeInfo
					, "UInt" , pTypeAttr
					, "Int")
					
	If (__IsComError("ITypeInfo::ReleaseTypeAttr", iErr))
		Return False
		
	Return True
}

; ## IDispatch #################################################################
/****ih* Interfaces/IDispatch
* VTable
*	0   call_QueryInterface    Returns pointers to supported interfaces.
*	1   call_AddRef            Increments reference count.
*	2   call_Release           Decrements reference count.
*	3 * call_GetTypeInfoCount  Retrieves the number of type information interfaces that an object provides (either 0 or 1).
*	4 * call_GetTypeInfo       Gets the type information for an object.
*	5   call_GetIDsOfNames     Maps a single member and an optional set of argument names to a corresponding set of integer DISPIDs.
*	6   call_Invoke            Provides access to properties and methods exposed by an object.
*	(only the * members are implemented)
******
*/


; ..............................................................................
/****iI* IDispatch/__IDispatch_GetTypeInfoCount
* Description
*	Retrieves the number of type information interfaces that an object provides (either 0 or 1).
* Usage
*	__IDispatch_GetTypeInfoCount(pIDispatch)
* Parameters
*	*  pIDispatch -- (Integer) Pointer to an IDispatch interface of an object.
* Return Value
*	* Success: (Integer) Number of type information interfaces that an object provides (0 or 1).
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	Used by __GetIDispatch
* Related
*	IDispatch
******
*/
__IDispatch_GetTypeInfoCount(ppDispatch)
{
	iErr := DllCall(__VTable(ppDispatch, 3), "UInt", ppDispatch
					, "UInt*", iTypeInfoCount
					, "Int")
	
	If (__IsComError("IDispatch::GetTypeInfoCount", iErr))
		Return
					
	Return iTypeInfoCount
}


; ..............................................................................
/****iI* IDispatch/__IDispatch_GetTypeInfo
* Description
*	Gets the type information for an object.
* Usage
*	__IDispatch_GetTypeInfo(pIDispatch [, iLocaleID ] )
* Parameters
*	*  pIDispatch -- (Integer) Pointer to an IDispatch interface of an object.
*	*  iLocaleID -- (Optional) (Integer) Locale ID to use.
* Return Value
*	* Success: (Integer) Pointer to ITypeInfo interface of the object
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*	Used by __GetIDispatch
* Related
*	IDispatch
******
*/
__IDispatch_GetTypeInfo(ppDispatch, LCID = 0)
{
	iErr := DllCall(__VTable(ppDispatch, 4), "UInt", ppDispatch
					, "UInt" , 0   ; iTInfo
					, "UInt" , LCID
					, "UInt*", ppTypeInfo
					, "Int")
					
	If (__IsComError("IDispatch::GetTypeInfo", iErr))
		Return
					
	Return ppTypeInfo
}

; ## IUnknown ##################################################################
/****ih* Interfaces/IUnknown
* VTable
*	0 * call_QueryInterface    Returns pointers to supported interfaces.
*	1 * call_AddRef            Increments reference count.
*	2 * call_Release           Decrements reference count.
*	(only the * members are implemented)
******
*/

; ..............................................................................
/****iI* IUnknown/__IUnknown_QueryInterface
* Description
*	Returns pointers to supported interfaces.
* Usage
*	__IUnknown_QueryInterface(pIUnknown, ?IId)
* Parameters
*	*  pIUnknown -- (Integer) Pointer to an IUnknown interface of an object.
*	*  ?IId -- (String) ANSI String (e.g. "{00000000-0000-0000-C000-000000000046}")
*	          or (Integer) pointer to the binary form of the Interface ID to query for.
* Return Value
*	* Success: (Integer) Pointer to new requested interface of the object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0, or non-critical error description. 
*	* Failure: error description.
* Remarks
*
* Related
*	IUnknown
******
*/
__IUnknown_QueryInterface(ppv, iid)
{
	; Is it a pointer to a binary IID? (check if it's a number)
	If iid is Integer
	{
		iErr := DllCall(__VTable(ppv,0), "UInt", ppv
					, "UInt" , iid
					, "UInt*", ppvNewInterface
					, "Int")
	}
	Else ; otherwise we assume it is a class id string
	{
		If (!__IIDFromString(iid, biniid))
			Return 
	
		iErr := DllCall(__VTable(ppv,0), "UInt", ppv
					, "Str"  , biniid
					, "UInt*", ppvNewInterface
					, "Int")
	}
	
	If (__IsComError("IUnknown::QueryInterface", iErr))
		Return

	Return ppvNewInterface
}


; ..............................................................................
/****iI* IUnknown/__IUnknown_AddRef
* Description
*	Increments reference count.
* Usage
*	__IUnknown_AddRef(pIUnknown)
* Parameters
*	*  pIUnknown -- (Integer) Pointer to an IUnknown interface of an object.
* Return Value
*	* Success: (Integer) Number of references to the object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: 0. 
*	* Failure: error description.
* Remarks
*
* Related
*	IUnknown
******
*/
__IUnknown_AddRef(ppv)
{
	iCount := DllCall(__VTable(ppv,1), "UInt", ppv, "Int")
	If (ErrorLevel <> 0)
	{
		__ComError(ErrorLevel, "IUnknown::AddRef: DllCall error " ErrorLevel)
		Return
	}
	Return iCount
}


; ..............................................................................
/****iI* IUnknown/__IUnknown_Release
* Description
*	Decrements reference count.
* Usage
*	__IUnknown_Release(pIUnknown)
* Parameters
*	*  pIUnknown -- (Integer) Pointer to an IUnknown interface of an object.
* Return Value
*	* Success: (Integer) Number of remaining references to the object.
*	* Failure: None ("").
* ErrorLevel
*	* Success: Not changed.
*	* Failure: If ErrorLevel = 0, ErrorLevel is replaced with error description,
*	           otherwise ErrorLevel is appended with a newline and error description.
* Remarks
*
* Related
*	IUnknown
******
*/
__IUnknown_Release(ppv)
{
	ErrLvl := ErrorLevel ; save ErrorLevel
	iCount := DllCall(__VTable(ppv,2), "UInt", ppv, "Int")
	
	If (ErrorLevel <> 0) ; If failed DllCall()
	{
		__ComError(ErrorLevel, "IUnknown::Release: DllCall error " ErrorLevel)
		
		; Append the old ErrorLevel if there was already an error
		If (ErrLvl <> 0)
			ErrorLevel := ErrLvl . "`n" . ErrorLevel
			
		Return
	}
	Else
	{
		ErrorLevel := ErrLvl ; just restore ErrorLevel
	}
	Return iCount
}

