/*
IEControl.ahk
*/

/* Supported Functions
IE_Add(hWnd, x, y, w, h)
IE_Move(pwb, x, y, w, h)
IE_LoadURL(pwb, u)
IE_LoadHTML(pwb, h)
IE_GoBack(pwb)
IE_GoForward(pwb)
IE_GoHome(pwb)
IE_GoSearch(pwb)
IE_Refresh(pwb)
IE_Stop(pwb)
; Removed: IE_Document(pwb)
IE_GetTitle(pwb)
IE_GetUrl(pwb)
IE_Busy(pwb)
IE_Quit(pwb)            ; iexplore.exe only
IE_hWnd(pwb)            ; iexplore.exe only
IE_FullName(pwb)         ; iexplore.exe only
IE_GetStatusText(pwb)         ; iexplore.exe only
IE_SetStatusText(pwb, sText = "")   ; iexplore.exe only
IE_ReadyState(pwb)
IE_Open(pwb)
IE_New(pwb)
IE_Save(pwb)
IE_SaveAs(pwb)
IE_Print(pwb)
IE_PrintPreview(pwb)
IE_PageSetup(pwb)
IE_Properties(pwb)
IE_Cut(pwb)
IE_Copy(pwb)
IE_Paste(pwb)
IE_SelectAll(pwb)
IE_Find(pwb)
IE_DoFontSize(pwb, s)
IE_InternetOptions(pwb)
IE_ViewSource(pwb)
IE_AddToFavorites(pwb)
IE_MakeDesktopShortcut(pwb)
IE_SendEMail(pwb)
__CGID_MSHTML(pwb, nCmd, nOpt = 0)
GetWebControl()
UrlHistoryEnum()
UrlHistoryClear()
*/

/*
Type Library description: "Microsoft Internet Controls"
File:   shdocvw.dll
ProgId: "Shell.Explorer"
*/


#Include ws4ahk.ahk

IE_Add(hWnd, x, y, w, h)
{
   InitComControls()
   Return GetComControlInHWND( CreateComControlContainer(hWnd, x, y, w, h, "Shell.Explorer") )
}

IE_Move(pwb, x, y, w, h)
{
   WinMove, % "ahk_id " . GetHWNDofComControl(pwb), , x, y, w, h
}

IE_LoadURL(pwb, u)
{
	If (!WS_Exec(pwb ".Navigate %s", u))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_LoadHTML(pwb, h)
{
	If (!WS_Exec(pwb ".Navigate %s", "about:" . h))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_GoBack(pwb)
{
	If (!WS_Exec(pwb ".GoBack"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_GoForward(pwb)
{
	If (!WS_Exec(pwb ".GoForward"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_GoHome(pwb)
{
	If (!WS_Exec(pwb ".GoHome"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_GoSearch(pwb)
{
	If (!WS_Exec(pwb ".GoSearch"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Refresh(pwb)
{
	If (!WS_Exec(pwb ".Refresh"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Stop(pwb)
{
	If (!WS_Exec(pwb ".Stop"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_GetTitle(pwb)
{
	If (!WS_Eval(pTitle, pwb ".LocationName"))
		Msgbox % A_LineFile ": " ErrorLevel
   Return sTitle
}

IE_GetUrl(pwb)
{
	If (!WS_Eval(sUrl, pwb ".LocationURL"))
		Msgbox % A_LineFile ": " ErrorLevel
   Return sUrl
}

IE_Busy(pwb)
{
	If (!WS_Eval(bBusy, pwb ".Busy"))
		Msgbox % A_LineFile ": " ErrorLevel
   Return bBusy
}

IE_Quit(pwb)            ; iexplore.exe only
{
	If (!WS_Exec(pwb ".Quit"))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_hWnd(pwb)            ; iexplore.exe only
{
	If (!WS_Eval(hIE, pwb ".HWND"))
		Msgbox % A_LineFile ": " ErrorLevel
   Return hIE
}

IE_FullName(pwb)         ; iexplore.exe only
{
	If (!WS_Eval(sFile, pwb ".FullName"))
		Msgbox % A_LineFile ": " ErrorLevel
   Return sFile
}

IE_GetStatusText(pwb)         ; iexplore.exe only
{
	If (!WS_Eval(sText, pwb ".StatusText"))
		Msgbox % A_LineFile ": " ErrorLevel
   Return sText
}

IE_SetStatusText(pwb, sText = "")   ; iexplore.exe only
{
	If (!WS_Exec(pwb ".StatusText = %s", sText))
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_ReadyState(pwb)
{
/*
   READYSTATE_UNINITIALIZED = 0      ; Default initialization state.
   READYSTATE_LOADING       = 1      ; Object is currently loading its properties.
   READYSTATE_LOADED        = 2      ; Object has been initialized.
   READYSTATE_INTERACTIVE   = 3      ; Object is interactive, but not all of its data is available.
   READYSTATE_COMPLETE      = 4      ; Object has received all of its data.
*/
	If (!WS_Eval(nReady, pwb ".ReadyState"))
		Msgbox % A_LineFile ": " ErrorLevel
	Return nReady
}

IE_Open(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 1, 0")) ; OLECMDID_OPEN
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_New(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 2, 0")) ; OLECMDID_NEW
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Save(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 3, 0")) ; OLECMDID_SAVE
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_SaveAs(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 4, 0")) ; OLECMDID_SAVEAS
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Print(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 6, 0")) ; OLECMDID_PRINT
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_PrintPreview(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 7, 0")) ; OLECMDID_PRINTPREVIEW
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_PageSetup(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 8, 0")) ; OLECMDID_PAGESETUP
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Properties(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 10, 0")) ; OLECMDID_PROPERTIES
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Cut(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 11, 0")) ; OLECMDID_CUT
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Copy(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 12, 0")) ; OLECMDID_COPY
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Paste(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 13, 0")) ; OLECMDID_PASTE
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_SelectAll(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 17, 0")) ; OLECMDID_SELECTALL
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_Find(pwb)
{
	If (!WS_Exec(pwb ".ExecWB 32, 0")) ; OLECMDID_FIND
		Msgbox % A_LineFile ": " ErrorLevel
}

IE_DoFontSize(pwb, s)
{
/*
   s = 4   ; Largest
   s = 3   ; Larger
   s = 2   ; Medium
   s = 1   ; Smaller
   s = 0   ; Smallest
*/
	If (!WS_Exec(pwb ".ExecWB 19, 2, " s)) ; OLECMDID_ZOOM
		Msgbox % A_LineFile ": " ErrorLevel
}

; ==============================================================================
; ==============================================================================
; The remaining functions use features unavailable to the scripting language
; so they must be implemented with raw COM calls.

IE_InternetOptions(pwb)
{
   __CGID_MSHTML(pwb, 2135)
}

; ..............................................................................

IE_ViewSource(pwb)
{
   __CGID_MSHTML(pwb, 2139)
}

; ..............................................................................

IE_AddToFavorites(pwb)
{
   __CGID_MSHTML(pwb, 2261)
}

; ..............................................................................

IE_MakeDesktopShortcut(pwb)
{
   __CGID_MSHTML(pwb, 2266)
}

; ..............................................................................

IE_SendEMail(pwb)
{
   __CGID_MSHTML(pwb, 2288)
}

; ..............................................................................


GetWebControl(hIESvr)
{
	; This code based on "How to get IHTMLDocument2 from a HWND"
	; http://support.microsoft.com/kb/q249232/
   __IIDFromString("{332C4425-26CB-11D0-B483-00C04FD90119}", IID_IHTMLDocument2)
   DllCall("SendMessageTimeout"
   		, "UInt",  hIESvr ; Handle to the window that will receive the message.
		, "UInt",  DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
		, "Int",   0    ; wParam
		, "Int",   0    ; lParam
		, "UInt",  2    ; [SMTO_ABORTIFHUNG] Don't wait if the receiving thread "hangs."
		, "UInt",  1000 ; duration, in milliseconds, of time-out period
		, "UIntP", lResult) ; Receives the result of the message processing.
		
   DllCall("oleacc\ObjectFromLresult"
   		, "UInt", lResult
		, "Str", IID_IHTMLDocument2
		, "Int", 0
		, "UIntP", phd)
   
   static SID_SWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"
   static IID_IWebBrowser2   := "{D30C1661-CDAF-11D0-8A3E-00C04FC9E26E}"
   pwb := IServiceProvider_QueryService(phd, SID_SWebBrowserApp, IID_IWebBrowser2)
   IUnknown_Release(phd)
   Return pwb
}

; ..............................................................................

UrlHistoryEnum()
{
	; Unforunately IID_IUrlHistoryStg doesn't speak IDispatch
	; so it's not usable within VBScript/JScript :(
	; Some details about this function
	; http://www.codeguru.com/cpp/i-n/ieprogram/displayinginformation/article.php/c13353/
	static CLSID_CUrlHistory  := "{3C374A40-BAE4-11CF-BF7D-00AA006946EE}"
	static IID_IUrlHistoryStg := "{3C374A41-BAE4-11CF-BF7D-00AA006946EE}"
	puh := CreateObject(CLSID_CUrlHistory, IID_IUrlHistoryStg)
	IfEqual puh,
	{
		Msgbox % ErrorLevel
		Return
	}
	iErr := DllCall(__VTable(puh, 7), "Uint", puh ; puh->EnumUrls(IEnumSTATURL &peu)
			, "UintP", peu
			, "UInt") 
	If (iErr <> 0 Or ErrorLevel <> 0)
	{
		Msgbox % "CUrlHistory::EnumUrls() error: ErrorLevel=" ErrorLevel "  iErr=" iErr
		Return
	}
	VarSetCapacity(var, 40, 0) ; sizeof(STATURL) = 40
	NumPut(VarSetCapacity(var), var) ; var.cbSize = 40
   
	Loop
	{
		iErr := DllCall(__VTable(peu, 3), "Uint", peu ; peu->Next(1, &var, NULL)
						, "Uint", 1
						, "Uint", &var
						, "Uint", 0
						, "UInt")
		If (iErr <> 0)
			Break
		pUrl   := NumGet(var, 4) ; pUrl   = var->pwcsUrl
		pTitle := NumGet(var, 8) ; pTitle = var->pwcsTitle
		If (pUrl <> 0)
		{
			__Unicode2Ansi(pUrl  , sUrl  )
			DllCall("ole32\CoTaskMemFree", "UInt", pUrl)
		}
		If (pTitle <> 0)
		{
			__Unicode2Ansi(pTitle, sTitle)
			DllCall("ole32\CoTaskMemFree", "UInt", pTitle)
		}
		sHistory .= sUrl . "|" . sTitle . "`n"
	}
	
	ReleaseObject(peu)
	ReleaseObject(puh)
	Return sHistory
}

; ..............................................................................

UrlHistoryClear()
{
	; Unforunately IID_IUrlHistoryStg2 doesn't speak IID_IDispatch
	; so it's not usable within VBScript/JScript :(
	static CLSID_CUrlHistory   := "{3C374A40-BAE4-11CF-BF7D-00AA006946EE}"
	static IID_IUrlHistoryStg2 := "{AFA0DC11-C313-11D0-831A-00C04FD5AE38}"
	puh := CreateObject(CLSID_CUrlHistory, IID_IUrlHistoryStg2)
	iErr := DllCall(__VTable(puh, 9), "Uint", puh) ; puh->ClearHistory()
	If (iErr <> 0 Or ErrorLevel <> 0)
	{
		Msgbox % "CUrlHistory::ClearHistory() error: ErrorLevel=" ErrorLevel "  iErr=" iErr
		Return
	}
	ReleaseObject(puh)
}


; ##############################################################################

__CGID_MSHTML(pwb, nCmd, nOpt = 0)
{
   ; Unforunately IID_IOleCommandTarget doesn't speak IDispatch
   ; so it's not usable within VBScript/JScript :(
   static CGID_MSHTML           := "{DE4BA900-59CA-11CF-9592-444553540000}"
   static IID_IOleCommandTarget := "{B722BCCB-4E68-101B-A2BC-00AA00404770}"
   __CLSIDFromString(CGID_MSHTML, sbinCGID_MSHTML)
   pct := IUnknown_QueryInterface(pwb, IID_IOleCommandTarget)
   DllCall(__VTable(pct, 4), "Uint", pct     ; pct->Exec(...)
   					, "str", sbinCGID_MSHTML ; Pointer to command group
					, "Uint", nCmd           ; Identifier of command to execute
					, "Uint", nOpt           ; Options for executing the command
					, "Uint", 0              ; Pointer to input arguments
					, "Uint", 0)             ; Pointer to command output
   IUnknown_Release(pct)
}


; ## IServiceProvider ##########################################################

IServiceProvider_QueryService(ppv, SID, IID)
{
   __IIDFromString(SID, binSID)
   __IIDFromString(IID, binIID)
   static IID_IServiceProvider := "{6D5140C1-7436-11CE-8034-00AA006009FA}"
   psp := IUnknown_QueryInterface(ppv, IID_IServiceProvider)
   DllCall(__VTable(psp,3), "Uint", psp
   			, "str", binSID ; GUID identifying the service
			, "str", binIID ; IID identifying an interface provided by that service
			, "UintP", ppvService) ; address of a pointer to receive the requested interface
   IUnknown_Release(psp)
   Return ppvService
}

