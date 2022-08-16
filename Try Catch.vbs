'****************************************************************************
'* "Try Catch.vbs"  (VBScript Include)
'* 
'* Defines a structured exception handling helper class for VBScript.
'* 
'* A big issue in VBScript is that On Error GoTo 0 also clears the Err
'* object, which makes it difficult to handle some errors and propagate
'* unhandled ones, since you'll have to copy all the fields of the Err
'* object before the "On Error GoTo 0" statement, and use these to raise
'* the error again.
'* 
'* This helper class is designed to make it easier to handle errors in an
'* exception-handling pattern in VBScript by palliating these issues.
'* It is typically used as follows:
'* With New Try: On Error Resume Next
'*   ' statements to try, beware contrary to typical try - catch blocks,
'*   ' an error here will not interrupt and exit to the catch block, all
'*   ' lines are executed before handling checks if any error occured.
'*   ' A "Do ... Loop While False" could be used with one or several
'*   ' "If Err.Number <> 0 Then Exit Do" statements for similar results.
'* .Catch: On Error GoTo 0
'*   ' Handle any error, typically by testing .Number for errors we want to
'*   ' recover from. The .RaiseAgain method can be used to raise the error
'*   ' again as it, and propagate (bubble) it up the call stack.
'*   ' Note code here will run even if no error occured, with .Number = 0.
'*   Select Case .Number ' or a simpler "If .Number = 5 Then" could be used.
'*     Case 5 ' some error we want to handle
'*       ' error handling statements
'*     Case Else
'*       ' Raise the error again if we're not recovering from it.
'*       ' This is safe to call even if .Number = 0, which won't re-raise.
'*       .RaiseAgain
'*   End Select
'* End With
'* ' Code here will only get executed if no error was raised after the
'* ' ".Catch: On Error GoTo 0" line above.
'* 
'****************************************************************************


Option Explicit

Class Try
	Private m_Number
	Private m_Source
	Private m_Description
	Private m_HelpFile
	Private m_HelpContext
	
	' Capture the error state
	Public Sub Catch
		m_Number = Err.Number
		m_Source = Err.Source
		m_Description = Err.Description
		m_HelpFile = Err.HelpFile
		m_HelpContext = Err.HelpContext
	End Sub
	
	' Propagate (raise again / rethrow) the error previously captured, if any
	Public Sub RaiseAgain
		If m_Number <> 0 Then
			Err.Raise m_Number, m_Source, m_Description, m_HelpFile, m_HelpContext
			' Note re-raised errors will appear as originating from the line above,
			' see the call stack to find the real error origin.
		End If
	End Sub
	
	' Property accessors
	
	Public Property Get Number
		Number = m_Number
	End Property
	
	Public Property Get Source
		Source = m_Source
	End Property
	
	Public Property Get Description
		Description = m_Description
	End Property
	
	Public Property Get HelpFile
		HelpFile = m_HelpFile
	End Property
	
	Public Property Get HelpContext
		HelpContext = m_HelpContext
	End Property
	
End Class
