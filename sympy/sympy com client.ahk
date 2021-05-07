#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0
ListLines Off

sympyComServer:=ComObjCreate("Python.SimplifyExpr")
;or
; pythonComServer:=ComObjCreate("{4530C817-6C66-46C8-8FB0-E606970A8DF6}") ;use your own CLSID

; clipboard:=sympyComServer.parExprN("1+3*7")
clipboard:=sympyComServer.parExprN("1/3 + 1/2")

$#s::
clipboard:=sympyComServer.parExprN(clipboard)
return

f3::Exitapp