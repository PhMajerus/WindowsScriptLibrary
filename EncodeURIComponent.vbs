'
' Encodes a Uniform Resource Identifier (URI) component by replacing each
' instance of certain characters by one, two, three, or four escape sequences
' representing the UTF-8 encoding of the character.
' 
' This is a form of percent-encoding limiting the character set to ASCII
' characters allowed as part of a URI field. This means it can be used as
' a querystring value as it also escapes the ;,/?:@&=+$# characters.
' It does not provide the same result as IIS ASP Server.URLEncode, which
' encodes all non-alphanumeric characters (including -_.!~*'() characters).
' 
' Note the function code is exactly the same as for EncodeURI, but with the
' extra characters to encode removed from the sNotEscaped string.
' 
' Dependencies:
' None, pure VBScript implementation.
' 
' - Philippe Majerus, July 2018
'

Option Explicit

Function EncodeURIComponent (URI)
	' Unreserved marks, lists the characters that are allowed as it in a URI field (querystring value)
	Const sNotEscaped = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_.!~*'()"
	
	Dim L, I, C, C2, B1, B2, B3, B4
	EncodeURIComponent = vbNullString
	L = Len(URI)
	I = 1
	While I <= L
		C = Mid(URI, I, 1)
		If InStr(sNotEscaped, C) <> 0 Then
			' Character is allowed, append it without escaping it
			EncodeURIComponent = EncodeURIComponent & C
		Else
			C = AscW(C)
			
			' Combine any UTF-16 surrogate pair
			If C >= &hD800 And C <= &hDBFF Then
				I = I + 1
				C2 = AscW(Mid(URI, I, 1))
				If Not (C2 >= &hDC00 And C2 <= &hDFFF) Then
					' High surrogate without low surrogate
					Err.Raise 5, , "The URI to be encoded contains an invalid character"
				End If
				C = &h10000 + ((C - &hD800) * &h400) + (C2 - &hDC00)
			ElseIf C >= &hDC00 And C <= &hDFFF Then
				' Low surrogate without high surrogate before it
				Err.Raise 5, , "The URI to be encoded contains an invalid character"
			ElseIf C < 0 Then
				' AscW returns a signed int, fix it to a positive long.
				C = &h10000 + C
			End If
			
			' Encode code point
			If C < &h80 Then
				B1 = C ' 0xxxxxxx
				B1 = "%" & Right("0" & Hex(B1), 2)
				EncodeURIComponent = EncodeURIComponent & B1
			ElseIf C < &h800 Then
				B1 = (C \ &h40)   Or &hC0 ' 110xxxxx
				B2 = (C And &h3F) Or &h80 ' 10xxxxxx
				B1 = "%" & Right("0" & Hex(B1), 2)
				B2 = "%" & Right("0" & Hex(B2), 2)
				EncodeURIComponent = EncodeURIComponent & (B1 & B2)
			ElseIf C < &h10000 Then
				B1 =  (C \ &h1000)         Or &hE0 ' 1110xxxx
				B2 = ((C \ &h40) And &h3F) Or &h80 ' 10xxxxxx
				B3 = ( C         And &h3F) Or &h80 ' 10xxxxxx
				B1 = "%" & Right("0" & Hex(B1), 2)
				B2 = "%" & Right("0" & Hex(B2), 2)
				B3 = "%" & Right("0" & Hex(B3), 2)
				EncodeURIComponent = EncodeURIComponent & (B1 & B2 & B3)
			Else
				B1 =  (C \ &h40000)          Or &hF0 ' 11110xxx
				B2 = ((C \ &h1000) And &h3F) Or &h80 ' 10xxxxxx
				B3 = ((C \ &h40)   And &h3F) Or &h80 ' 10xxxxxx
				B4 = ( C           And &h3F) Or &h80 ' 10xxxxxx
				B1 = "%" & Right("0" & Hex(B1), 2)
				B2 = "%" & Right("0" & Hex(B2), 2)
				B3 = "%" & Right("0" & Hex(B3), 2)
				B4 = "%" & Right("0" & Hex(B4), 2)
				EncodeURIComponent = EncodeURIComponent & (B1 & B2 & B3 & B4)
			End If
		End If
		' loop
		I = I + 1
	WEnd
End Function
