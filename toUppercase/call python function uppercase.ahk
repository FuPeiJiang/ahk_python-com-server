#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0
ListLines Off

pythonComServer:=ComObjCreate("Python.stringUppercaser")
;or
; pythonComServer:=ComObjCreate("{C70F3BF7-2947-4F87-B31E-9F5B8B13D24F}") ;use your own CLSID

MsgBox % pythonComServer.toUppercase("hello world")
Exitapp

f3::Exitapp