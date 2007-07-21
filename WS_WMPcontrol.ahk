#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.

#Include EasyScript.ahk
#SingleInstance Force

sUrl := A_WinDir . "\clock.avi"   ; Specify the media file.

GoSub, GuiStart

Gui, +Resize +LastFound
Gui, Show, w800 h600 Center, WMP
hWnd := WinExist()

vbwmp := "oWMP"
If (!WS_Exec("Set %v = CreateObject(%s)", vbwmp, "WMPlayer.OCX"))
	Msgbox % A_LineFile ": " ErrorLevel
If (!WS_Eval(pwmp, vbwmp))
	Msgbox % A_LineFile ": " ErrorLevel
AttachComControlToHWND(pwmp, hWnd)

OpenURL(vbwmp, sUrl)
;Invoke(pwmp, "Play") ; Play is not a member of WMP, and not needed to play
Return

GuiStart:
	WS_Initialize("VBScript")
	InitComControls()
Return

GuiClose:
	Gui, %A_Gui%:Destroy
	ReleaseObject(pwmp)
	UninitComControls()
	WS_Uninitialize()
ExitApp

OpenURL(vbwmp, sUrl)
{
   WS_Exec(vbwmp ".URL = %s", sUrl)
}

