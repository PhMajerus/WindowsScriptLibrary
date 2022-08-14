'
' Convert a string to a COM binary buffer (bytes array / VT_ARRAY|VT_UI1),
' using the specified charset (use "utf-8" for UTF-8 encoding).
' 
' See HKEY_CLASSES_ROOT\MIME\Database\Charset for a list of supported charsets on your system.
' 
' Dependencies:
' - Microsoft ActiveX Data Objects ("ADODB.Stream")
'

Option Explicit

Function StringToBinary (Text, Charset)
	Dim TxtStm, CharsetNormalized
	Set TxtStm = CreateObject("ADODB.Stream")
	TxtStm.Open
	TxtStm.Type = 2 ' adTypeText
	With New Try: On Error Resume Next
		TxtStm.Charset = Charset
	.Catch: On Error GoTo 0
		If .Number = 3001 Then ' Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
			Err.Raise 5, , "Unknown charset requested: "& Charset
		Else
			.RaiseAgain
		End If
	End With
	CharsetNormalized = TxtStm.Charset ' retrieve standardized charset name
	TxtStm.WriteText Text
	
	' switch stream to binary mode and return its contents
	TxtStm.Position = 0
	TxtStm.Type = 1 ' adTypeBinary
	If LCase(CharsetNormalized) = "utf-8" Then
		TxtStm.Position = 3 ' skip UTF-8 BOM
	ElseIf CharsetNormalized = "Unicode" Then
		' This also handles "UTF-16" and "utf-16", they automatically get changed to "Unicode".
		TxtStm.Position = 2 ' skip UTF-16 BOM
	End If ' No BOM is added for UTF-16BE
	StringToBinary = TxtStm.Read()
End Function
