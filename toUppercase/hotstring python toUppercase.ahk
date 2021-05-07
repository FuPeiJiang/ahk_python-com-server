#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, force
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetBatchLines, -1
#KeyHistory 0
ListLines Off
#Persistent
#MaxThreadsPerHotkey 4

pythonComServer:=ComObjCreate("Python.stringUppercaser")
; OR
; pythonComServer:=ComObjCreate("{C70F3BF7-2947-4F87-B31E-9F5B8B13D24F}") ;use your own CLSID


; * do not wait for string to end
; C case sensitive
:*:hello world::

savedHotstring:=A_ThisHotkey

;theActualHotstring=savedHotstring[second colon:end of string]
theActualHotstring:=SubStr(savedHotstring, InStr(savedHotstring, ":",, 2) + 1)
send, % pythonComServer.toUppercase(theActualHotstring)


return



f3::Exitapp