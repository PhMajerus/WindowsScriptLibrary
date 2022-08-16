'
' Create a new GUID.
'

Option Explicit

Function CreateGUID
	Dim ScrTL
	Set ScrTL = CreateObject("Scriptlet.TypeLib")
	CreateGUID = ScrTL.GUID
End Function
