'
' Convert a COM binary buffer (bytes array / VT_ARRAY|VT_UI1) to a string,
' using the specified charset (use "utf-8" for UTF-8 encoding).
' 
' See HKEY_CLASSES_ROOT\MIME\Database\Charset for a list of supported charsets on your system.
' Warning: csISOLatin1, ISO_8859-1, ... are mapped to Windows 1252, not proper ISO/IEC 8859-1.
' 
' Dependencies:
' - Microsoft ActiveX Data Objects ("ADODB.Stream")
' 
' For example, to explicitely convert using MS-DOS US codepage 437:
' S = BinaryToString(FileToBinary("axsh.ans"),"437")
' And if you want to handle SUB as an end-of-file like CP/M and MS-DOS does:
' With New RegExp: .Pattern = "\x1A[\s\S]*": S = .Replace(S,""): End With
'

Option Explicit

Function BinaryToString (Data, Charset)
	Dim ADOStm
	Set ADOStm = CreateObject("ADODB.Stream")
	With New Try: On Error Resume Next
		ADOStm.Charset = Charset
	.Catch: On Error GoTo 0
		If .Number = 3001 Then ' Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
			Err.Raise 5, , "Unknown charset requested: "& Charset &"."
		Else
			.RaiseAgain
		End If
	End With
	ADOStm.Type = 1 ' adTypeBinary
	ADOStm.Open
	ADOStm.Write Data
	ADOStm.Position = 0
	ADOStm.Type = 2 ' adTypeText
	With New Try: On Error Resume Next
		BinaryToString = ADOStm.ReadText()
	.Catch: On Error GoTo 0
		If .Number = -2147024809 Then ' The parameter is incorrect.
			Err.Raise 5, , "The data to be decoded contains an invalid character for charset "& Charset &"."
		Else
			.RaiseAgain
		End If
	End With
End Function
