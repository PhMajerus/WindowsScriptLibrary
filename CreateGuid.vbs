'
' Generates a new GUID.
' 
' Dependencies:
' - Windows Script Component Runtime
'   (part of the base Windows Script platform)
'

Option Explicit

Function CreateGuid
	Dim ScrTL
	Set ScrTL = CreateObject("Scriptlet.TypeLib")
	CreateGuid = Left(ScrTL.GUID, 38)
End Function
