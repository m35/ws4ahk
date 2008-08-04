;
; AutoHotkey Version: 1.0.44+
; Language:       English
; Platform:       Win9x/NT
; Author:         Michael Sabin <msabin@alpine-la.com>
;
; Script Function:
;	Template AutoHotkey script.
;
#NoEnv

#Include ws4ahk.ahk


	;WS_Initializetest()
	
	;Test_InitAndLang()
	Test_CreateLangDel()

	;Test_CoInit()
	
	;Test_Init()
	
	;Test_comErr()

    WS_Initialize()
	;----------------------------
	
		;Test_SetLanguage()
		
		;Test_GetErrObj()
		
		;Test_MakeScriptObj()

		;Test_BSTR()
		
		;Test_ID()
		
		;Test_Exec()
		
		;Test_Eval()
	
	;----------------------------
	WS_Uninitialize()
	
	Return
	
	
Test_comErr()
{
	Loop, 9000
	{
		Random, rndnm
		__WS_IsComError("String", rndnm)
	}
}
	
Test_InitAndLang()
{
	ProgId_ScriptControl := "MSScriptControl.ScriptControl"
	IID_ScriptControl    := "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}"
	
	Msgbox % "InitAndLang test begin"
	
	Loop, 9000
	{
	
		iErr := DllCall("ole32\CoInitialize", "UInt", 0, "Int")
		
		If (iErr <> 0 Or ErrorLevel <> 0)
		{
			Msgbox % "Failed to init COM " ErrorLevel
			Return False
		}
		
		__WS_iScriptControlObj__ := WS_CreateObject(ProgId_ScriptControl, IID_ScriptControl)
			
		If (__WS_iScriptControlObj__ = "")
		{
			Msgbox % "Failed to create object " ErrorLevel
			ExitApp
		}

		If (!__WS_IScriptControl_Language(__WS_iScriptControlObj__, "VBScript"))
		{	; Failed to set language
			Msgbox % "Failed to set language " ErrorLevel
			ExitApp
		}
			
		iCount := __WS_IUnknown_Release(__WS_iScriptControlObj__)
		;Msgbox % iCount
		
		Sleep 100
		
		DllCall("ole32\CoUninitialize")
		;Msgbox % "End of loop " __WS_iScriptControlObj__
	
	}
	
	Msgbox % "InitAndLang test end"
}
	
Test_CreateLangDel()
{
	ProgId_ScriptControl := "MSScriptControl.ScriptControl"
	IID_ScriptControl    := "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}"
	
	iErr := DllCall("ole32\CoInitialize", "UInt", 0, "Int")
	
	If (iErr <> 0 Or ErrorLevel <> 0)
	{
		Msgbox % "Failed to init COM " ErrorLevel
		Return False
	}
	
	Msgbox % "CreateLangDel test begin"
	
	Loop, 9000
	{
	
		__WS_iScriptControlObj__ := WS_CreateObject(ProgId_ScriptControl, IID_ScriptControl)
			
		If (__WS_iScriptControlObj__ = "")
		{
			Msgbox % "Failed to create object " ErrorLevel
			ExitApp
		}

		If (!__WS_IScriptControl_Language(__WS_iScriptControlObj__, "VBScript"))
		{	; Failed to set language
			Msgbox % "Failed to set language " ErrorLevel
			ExitApp
		}
			
		iCount := __WS_IUnknown_Release(__WS_iScriptControlObj__)
		;Msgbox % iCount
		
	}
	
	DllCall("ole32\CoUninitialize")
	;Msgbox % "End of loop " __WS_iScriptControlObj__

	Msgbox % "CreateLangDel test end"
}
	
WS_Initializetest(sLanguage = "VBScript", sMSScriptOCX="")
{
	;global __WS_iScriptControlObj__, __WS_iScriptErrorObj__
	
	ProgId_ScriptControl := "MSScriptControl.ScriptControl"
	CLSID_ScriptControl  := "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}"
	IID_ScriptControl    := "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}"

	Msgbox % "WS_Initializetest start"
	Loop, 9000
	{
		/*
		If (__WS_iScriptControlObj__ <> "") {
			Msgbox % "Already set"
			Return True ; Windows Scripting has already been initialized
		}
		*/
					   
		; Init COM
		iErr := DllCall("ole32\CoInitialize", "UInt", 0, "Int")
		
		;If (__WS_IsComError("CoInitialize", iErr))
		;	Return False
		
		; Create Scripting Control
		;If (sMSScriptOCX="")
			__WS_iScriptControlObj__ := WS_CreateObject(ProgId_ScriptControl, IID_ScriptControl)
		;Else
		;	__WS_iScriptControlObj__ := WS_CreateObjectFromDll(sMSScriptOCX, CLSID_ScriptControl, IID_ScriptControl)
			
		;IfEqual, __WS_iScriptControlObj__,
		;{
		;	WS_Uninitialize()
		;	Return False
		;}
		
			
			; Set the language
			__WS_IScriptControl_Language(__WS_iScriptControlObj__, "VBScript")
			/*
			If (!__WS_IScriptControl_Language(__WS_iScriptControlObj__, sLanguage))
			{	; Failed to set language
				WS_Uninitialize()
				Return False
			}
			*/
			
			/*
			; Get Error object
			__WS_iScriptErrorObj__ := __WS_IScriptControl_Error(__WS_iScriptControlObj__)
			IfEqual, __WS_iScriptErrorObj__,
			{	; Failed to get error object
				WS_Uninitialize()
				Return False
			}
			__WS_IUnknown_Release(__WS_iScriptErrorObj__)
			*/
			
		__WS_IUnknown_Release(__WS_iScriptControlObj__)
		
		DllCall("ole32\CoUninitialize")
		;Msgbox % "End of loop " __WS_iScriptControlObj__
		Sleep 100
		
	}
	Msgbox % "WS_Initializetest end"
}
	
	
	
Test_CoInit()
{
	Msgbox % "CoInit test start"
	
	Loop, 9000
	{
		iErr := DllCall("ole32\CoInitialize", "UInt", 0, "Int")
	
		If (__WS_IsComError("CoInitialize", iErr))
		{
			Msgbox % "Error CoInitialize: " ErrorLevel
			ExitApp
		}
		
		DllCall("ole32\CoUninitialize")
		If (ErrorLevel <> 0)
		{
			Msgbox % "Error CoUninitialize: " ErrorLevel
			ExitApp
		}
	}
	
	Msgbox % "CoInit test end"
}

Test_Init()
{
	Msgbox % "Init test start"

	If (False)
	{
		Loop, 3000
		{
			If (!WS_Initialize())
			{
				Msgbox % ErrorLevel
				ExitApp
			}
		}
	
		Loop, 3000
		{
			If (WS_Uninitialize() <> "")
			{
				Msgbox % ErrorLevel
				ExitApp
			}
		}
	}
	

	; There appears to be a memory leak here somewhere, but I don't see where.
	Loop, 5000
	{
		If (!WS_Initialize())
		{
			Msgbox % ErrorLevel
			ExitApp
		}
		If (WS_Uninitialize() <> "")
		{
			Msgbox % ErrorLevel
			ExitApp
		}
		;Sleep 100
	}
	
	Msgbox % "Init test ok"	
}

	
Test_GetErrObj()
{
	global __WS_iScriptControlObj__
	
	Msgbox % "Get error obj test start"
	
	Loop, 100000
	{
		pErrObj := __WS_IScriptControl_Error(__WS_iScriptControlObj__)
		If (!pErrObj)
		{
			Msgbox % "Failed to get error object: " ErrorLevel
			ExitApp
		}
		WS_ReleaseObject(pErrObj)
	}
	
	Msgbox % "Get error obj test done"
}


Test_SetLanguage()
{
	global __WS_iScriptControlObj__
	
	Msgbox % "Set language test start"
	
	Loop, 100000
	{
		If (!__WS_IScriptControl_Language(__WS_iScriptControlObj__, "VBScript"))
		{
			Msgbox % "Failed to set language: " ErrorLevel
			ExitApp
		}
	}
	
	Msgbox % "Set language test done"
}


Test_MakeScriptObj()
{
	static ProgId_ScriptControl := "MSScriptControl.ScriptControl"
	static CLSID_ScriptControl  := "{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}"
	static IID_ScriptControl    := "{0E59F1D3-1FBE-11D0-8FF2-00A0D10038BC}"
	
	Msgbox % "MakeScriptObj test start"
	
	Loop, 100000
	{
		pScriptCtrl := WS_CreateObject(ProgId_ScriptControl, IID_ScriptControl)
		
		If (!pScriptCtrl)
		{
			Msgbox % "Failed to MakeScriptObj: " ErrorLevel
			ExitApp
		}
		WS_ReleaseObject(pScriptCtrl)
	}
	
	Msgbox % "MakeScriptObj test done"
}

Test_BSTR()
{
	Msgbox % "BSTR test start"
	
	Loop, 9000
	{
	
		pBSTR := __WS_StringToBSTR("Happy day")
		
		bln1 := __WS_Unicode2ANSI(pBSTR, sAnsi)
		
		bln2 := __WS_FreeBSTR(pBSTR)
		
		If (bln1 = False Or bln2 <> 0 Or sAnsi <> "Happy day")
		{
			Msgbox % "Failure: " bln " " sAnsi
			ExitApp
		}
	}
	
	Msgbox % "BSTR test OK"
}


Test_ID()
{
	Msgbox % "ID conversion test start"
	
	bln := __WS_CLSIDFromProgID("Nothing.Error", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __WS_CLSIDFromProgID("SAPI.SpVoice", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __WS_CLSIDFromProgID("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __WS_CLSIDFromString("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __WS_CLSIDFromString("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __WS_CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038B}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __WS_CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC", sbinCls)
	Msgbox % bln " " ErrorLevel
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromProgID("Nothing.Error", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromProgID("SAPI.SpVoice", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromString("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromString("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038B}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	Loop, 9000
	{
		bln := __WS_IIDFromString("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_IIDFromString("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_IIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_IIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038B}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __WS_IIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}

	
	Msgbox % "ID conversion test ok"
}

Test_Exec()
{
	Msgbox % "Exec test start"
	
	bln := WS_Exec("x = 10000")
	Msgbox % bln " " ErrorLevel
	bln := WS_Exec("Sqrt")
	Msgbox % bln " " ErrorLevel
	
	Loop, 9000
	{
		bln := WS_Exec("x = 10000")
		If (!bln)
		{
			Msgbox % "Should have worked " bln " " ErrorLevel
		}
	}
	
	Loop, 9000
	{
		bln := WS_Exec("Sqrt")
		If (bln)
		{
			Msgbox % "Shouldn't have worked " bln " " ErrorLevel
		}
	}
	
	
	Loop, 9000
	{
		bln := WS_Exec("ReDim x(10) : x(1) = 10")
		If (!bln)
		{
			Msgbox % "Should have worked " bln " " ErrorLevel
		}
	}
	Msgbox % "Exec test ok"
	
}

Test_Eval()
{
	Msgbox % "Test_Eval start"
	
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "1"))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "1.5123"))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "9999999.9999999"))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, """abcasdfasdlfhjawlekjfhawoieufhadlkfhlas"""))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "True"))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "False"))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, ""))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "non_exist_var"))
		{
			Msgbox % "Failed " ErrorLevel
			ExitApp
		}
	}
	;Msgbox % RetVal
	
	WS_Exec("ReDim x(10)")
		
	Loop, 9000
	{
		If (!WS_Eval(RetVal, "x"))
		{
			Msgbox % "Failed " RetVal " " ErrorLevel
			ExitApp
		}
	}
	Msgbox % ErrorLevel
	
	Msgbox % "Test_Eval ok"
}


Test_Args()
{
	
}
