'
' Write a Binary buffer (byte array / VT_ARRAY|VT_UI1) to a file path, using
' only components included with Windows.
' 
' Dependencies:
' - Microsoft ActiveX Data Objects ("ADODB.Stream")
'

Option Explicit

Sub BinaryToFile (Binary, Filepath)
	Dim Stream
	Set Stream = CreateObject("ADODB.Stream")
	Stream.Type = 1 ' adTypeBinary
	Stream.Open
	Stream.Write Binary
	Stream.SaveToFile Filepath
End Sub
