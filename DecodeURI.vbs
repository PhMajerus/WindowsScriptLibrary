'
' Decodes a Uniform Resource Identifier (URI) previously created by EncodeURI
' or by a similar routine.
' Note it does not decode all percent-encoded characters by design, use
' DecodeURIComponent instead for a generic decoding function.
' 
' This function's code isn't very elegant but allows parsing without using
' any subfunction while limiting code redundancy.
' 
' This is almost the same code as DecodeURIComponent, except that it
' explicitely does not decode escape sequences that could not have been
' introduced by EncodeURI. This means percent-encoded ;,/?:@&=+$# characters
' are not decoded and left as it.
' 
' Dependencies:
' None, pure VBScript implementation.
' 
' - Philippe Majerus, July 2018, updated January 2019
'

Option Explicit

Function DecodeURI (EncodedURI)
	Const sNotDecoded = ";,/?:@&=+$#"
	
	Const sInvalidEncoding = "The URI to be decoded is not a valid encoding"
	Dim L, I, C, H, B1, B2, B3, B4
	L = Len(EncodedURI)
	DecodeURI = vbNullString
	I = 1
	While I <= L
		' Get next character
		C = Mid(EncodedURI, I, 1)
		I = I + 1
		
		If C = "%" Then
			' Character is escaped, parse 1, 2, 3 or 4 bytes sequence
			C = vbNullString ' flag to check for error when we're done with sequence
			
			If (I+1 <= L) Then ' checks "%xx" length
				' Get and check first byte
				H = "&h" & Mid(EncodedURI, I, 2)
				If Not IsNumeric(H) Then Err.Raise 5, , sInvalidEncoding
				B1 = CInt(H)
				I = I + 2
				
				If (B1 And &h80) = 0 Then
					' One byte character
					B1 = B1 And &h7F ' 0xxxxxxx
					C = ChrW(B1)
					' Now, handle special case of characters excluded from being decoded
					If InStr(sNotDecoded, C) <> 0 Then
						' Revert to percent-encoded form
						C = Mid(EncodedURI, I-3, 3)
					End If
				ElseIf (I+2 <= L) Then ' checks "%xx" length
					' Get and check second byte
					H = "&h" & Mid(EncodedURI, I+1, 2)
					If Not IsNumeric(H) Then Err.Raise 5, , sInvalidEncoding
					B2 = CInt(H)
					If (Mid(EncodedURI, I, 1) = "%") And ((B2 And &hC0) = &h80) Then
						I = I + 3
						B2 = B2 And &h3F ' 10xxxxxx
						
						If (B1 And &hE0) = &hC0 Then
							' Two bytes character
							B1 = B1 And &h1F ' 110xxxxx
							C = ChrW((B1 * &h40) + B2)
						ElseIf (I+2 <= L) Then ' checks "%xx" length
							' Get and check third byte
							H = "&h" & Mid(EncodedURI, I+1, 2)
							If Not IsNumeric(H) Then Err.Raise 5, , sInvalidEncoding
							B3 = CInt(H)
							If (Mid(EncodedURI, I, 1) = "%") And ((B3 And &hC0) = &h80) Then
								I = I + 3
								B3 = B3 And &h3F ' 10xxxxxx
								
								If (B1 And &hF0) = &hE0 Then
									' Three bytes character
									B1 = B1 And &h0F ' 1110xxxx
									C = ChrW((B1 * &h1000) + (B2 * &h40) + B3)
								ElseIf (I+2 <= L) Then ' checks "%xx" length
									' Get and check forth byte
									H = "&h" & Mid(EncodedURI, I+1, 2)
									If Not IsNumeric(H) Then Err.Raise 5, , sInvalidEncoding
									B4 = CInt(H)
									If (Mid(EncodedURI, I, 1) = "%") And ((B4 And &hC0) = &h80) Then
										I = I + 3
										B4 = B4 And &h3F ' 10xxxxxx
										
										If (B1 And &hF8) = &hF0 Then
											' Four bytes character
											B1 = B1 And &h07 ' 11110xxx
											' Character from supplementary plane, requires UTF-16 encoding
											C = (B1 * &h40000) + (B2 * &h1000) + (B3 * &h40) + B4
											C = C - &h010000
											C = ChrW(&hD800 + (C \ &h400)) & ChrW(&hDC00 + (C And &h3FF))
										End If
									End If
								End If
							End If
						End If
					End If
				End If
			End If
		End If
		
		If C = vbNullString Then
			Err.Raise 5, , sInvalidEncoding
		End If
		DecodeURI = DecodeURI & C
	WEnd
End Function
