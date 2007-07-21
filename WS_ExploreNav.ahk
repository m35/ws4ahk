#Include EasyScript.ahk

#m::ShellNavigate(A_MyDocuments,1)
#p::ShellNavigate(A_ProgramFiles,1)
#w::ShellNavigate(A_WinDir,1)

WS_Initialize()

ShellNavigate(sPath, nFlags=0, hWnd=0)
{
	static SID_STopLevelBrowser := "{4C96BE40-915C-11CF-99D3-00AA004AE837}"
	static IID_IShellBrowser    := "{000214E2-0000-0000-C000-000000000046}"
	
	If (!WS_Exec("Set psw = CreateObject(%s).Windows", "Shell.Application"))
		Msgbox % A_LineFile ": " ErrorLevel

	If hWnd || (hWnd := WinExist("ahk_class CabinetWClass")) || (hWnd := WinExist("ahk_class ExploreWClass"))
	{
		sFindWindow =
				(
				For Each win In psw
					If win.hWnd = `%v Then
						Set pwb = win 
						Exit For
					End If
				Next
				)
		If (!WS_Exec(sFindWindow, hWnd))
			Msgbox % A_LineFile ": " ErrorLevel
		If (!WS_Eval(pwb, "pwb"))
			Msgbox % A_LineFile ": " ErrorLevel
		If psb := IServiceProvider_QueryService(pwb, SID_STopLevelBrowser, IID_IShellBrowser)
			BrowseObject(psb, sPath, nFlags)
	}
	Else pwb := SHGetIDispatchForFolder(sPath)
}

IServiceProvider_QueryService(ppv, SID, IID)
{
   __IIDFromString(SID, binSID)
   __IIDFromString(IID, binIID)
   static IID_IServiceProvider := "{6D5140C1-7436-11CE-8034-00AA006009FA}"
   psp := IUnknown_QueryInterface(ppv, IID_IServiceProvider)
   DllCall(__VTable(psp,3), "Uint", psp
   			, "str", binSID ; GUID identifying the service
			, "str", binIID ; IID identifying an interface provided by that service
			, "UintP", ppv) ; address of a pointer to receive the requested interface
   IUnknown_Release(psp)
   Return ppv
}

BrowseObject(psb, sPath, nFlags = 0)
{
/*
   SBSP_DEFBROWSER           = 0x0000
   SBSP_SAMEBROWSER          = 0x0001
   SBSP_NEWBROWSER           = 0x0002
   SBSP_DEFMODE              = 0x0000
   SBSP_OPENMODE             = 0x0010
   SBSP_EXPLOREMODE          = 0x0020
   SBSP_HELPMODE             = 0x0040
   SBSP_NOTRANSFERHIST       = 0x0080
   SBSP_ABSOLUTE             = 0x0000
   SBSP_RELATIVE             = 0x1000
   SBSP_PARENT               = 0x2000
   SBSP_NAVIGATEBACK         = 0x4000
   SBSP_NAVIGATEFORWARD      = 0x8000
   SBSP_ALLOW_AUTONAVIGATE   = 0x10000
   SBSP_CALLERUNTRUSTED      = 0x00800000
   SBSP_TRUSTFIRSTDOWNLOAD   = 0x01000000
   SBSP_UNTRUSTEDFORDOWNLOAD = 0x02000000
   SBSP_NOAUTOSELECT         = 0x04000000
   SBSP_WRITENOHISTORY       = 0x08000000
   SBSP_TRUSTEDFORACTIVEX    = 0x10000000
   SBSP_REDIRECT             = 0x40000000
   SBSP_INITIATEDBYHLINKFRAME= 0x80000000
*/
	If sPath Is Integer
		pidl := SHGetFolderLocation(sPath)
	Else
		pidl := SHParseDisplayName(sPath)
	hResult := DllCall(__VTable(psb, 11), "Uint", psb, "Uint", pidl, "Uint", nFlags)
	DllCall("ole32\CoTaskMemFree", "UInt", pidl)
	Return hResult
}

SHGetIDispatchForFolder(sPath)
{
	If sPath Is Integer
		pidl := SHGetFolderLocation(sPath)
	Else
		pidl := SHParseDisplayName(sPath)
	DllCall("shdocvw\SHGetIDispatchForFolder", "Uint", pidl, "UintP", pwb)
	DllCall("ole32\CoTaskMemFree", "UInt", pidl)
	Return pwb
}

SHParseDisplayName(sPath)
{
	__Ansi2Unicode(sPath, wPath)
	DllCall("shell32\SHParseDisplayName", "Str", wPath, "Uint", 0, "UintP", pidl, "Uint", 0, "Uint", 0)
	Return pidl
}

SHGetFolderLocation(CSIDL)
{
   DllCall("shell32\SHGetFolderLocation", "Uint", 0, "int", CSIDL, "Uint", 0, "Uint", 0, "UintP", pidl)
   Return pidl
}
