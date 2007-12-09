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


	;Test_Init()
	
	
    WS_Initialize()
	
	;Test_BSTR()
	
	;Test_ID()
	
	;Test_Exec()
	
	Test_Eval()
	
	WS_Uninitialize()

	

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
	}
	
	Msgbox % "Init test ok"	
}

	
Test_BSTR()
{
	Msgbox % "BSTR test start"
	
	Loop, 9000
	{
	
		pBSTR := __StringToBSTR("Happy day")
		
		bln1 := __Unicode2ANSI(pBSTR, sAnsi)
		
		bln2 := __FreeBSTR(pBSTR)
		
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
	
	bln := __CLSIDFromProgID("Nothing.Error", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __CLSIDFromProgID("SAPI.SpVoice", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __CLSIDFromProgID("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __CLSIDFromString("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __CLSIDFromString("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038B}", sbinCls)
	Msgbox % bln " " ErrorLevel
	bln := __CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC", sbinCls)
	Msgbox % bln " " ErrorLevel
	
	Loop, 9000
	{
		bln := __CLSIDFromProgID("Nothing.Error", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __CLSIDFromProgID("SAPI.SpVoice", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __CLSIDFromString("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __CLSIDFromString("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038B}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __CLSIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}

	;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	Loop, 9000
	{
		bln := __IIDFromString("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __IIDFromString("{0E59F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (!bln)
		{
			Msgbox % "Should have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __IIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __IIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038B}", sbinCls)
		If (bln)
		{
			Msgbox % "Shouldn't have worked " ErrorLevel
			ExitApp
		}
	}
	
	Loop, 9000
	{
		bln := __IIDFromString("{0EG9F1D5-1FBE-11D0-8FF2-00A0D10038BC", sbinCls)
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
