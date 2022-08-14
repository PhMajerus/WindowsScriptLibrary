'
' Create a Binary buffer (byte array / VT_ARRAY|VT_UI1) from a file path,
' using only components included with Windows.
' 
' Dependencies:
' - Microsoft ActiveX Data Objects ("ADODB.Stream")
'

Option Explicit

Function FileToBinary (Filepath)
	Dim Stream
	Set Stream = CreateObject("ADODB.Stream")
	Stream.Type = 1 ' adTypeBinary
	Stream.Open
	Stream.LoadFromFile Filepath
	FileToBinary = Stream.Read()
End Function
