'
' Convert a COM binary buffer (bytes array / VT_ARRAY|VT_UI1) contents to a
' COM stream object (IStream), using only components included with Windows.
' 
' Dependencies:
' - Microsoft ActiveX Data Objects ("ADODB.Stream")
'

Option Explicit

Function BinaryToStream (Binary)
	Dim ADOStm
	Set ADOStm = CreateObject("ADODB.Stream")
	ADOStm.Type = 1 ' adTypeBinary
	ADOStm.Open
	ADOStm.Write Binary
	ADOStm.Position = 0
	Set BinaryToStream = ADOStm
End Function
