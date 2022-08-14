'
' Wraps an existing file in a stream object to access it as an
' IStream/ISequentialStream object.
' 
' This function requires the Microsoft Windows Image Mastering API v2.0
' (IMAPIv2), included in Windows Vista, 2008 and later, and installable on
' XP and 2003.
' It is much more efficient than using the ADODB.Stream object that many
' developers use for this purpose because it wraps the file in a stream from
' its original storage instead of loading it all into RAM.
' 
' Dependencies:
' - Microsoft Windows Image Mastering API v2.0 (IMAPIv2)
'

Option Explicit

Function FileToStream (Path)
	Dim IIM
	Set IIM = CreateObject("IMAPI2FS.MsftIsoImageManager")
	IIM.SetPath Path
	Set FileToStream = IIM.Stream
End Function
