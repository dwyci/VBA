Attribute VB_Name = "Module1"
Sub messagebox()

'Basic message box pop-up window.
'MsgBox "Hi there!"

'MsgBox "Hi again!", vbYesNoCancel + vbCritical, "My Title"

'Get the button that was clicked.
returnValue = MsgBox("Should we continue?", vbYesNo + vbCritical, "Important Notice")

'MsgBox returnValue

'Check which button was clicked
If returnValue = vbYes Then
    'Code to run if user clicks Yes button.
    MsgBox "Yes"
    
ElseIf returnValue = vbNo Then
    'Code to run if user clicks No button.
    
    'Exit the macro
    'Exit Sub
    
    MsgBox "No"
    
End If


End Sub

Sub GetUserInput()

End Sub
'useless version
'InputBox "hi there"

'get user inpt
inputValue = InputBox("hi there", "My Title", "Default Value")

'display value inpt by user
'MsgBox inputvalue

If inputValue = vbNullString Then

    'nothing submitt or cancelled button pressed
    
    MsgBox "No input"
    
    
Else
    'user submitted a value

     MsgBox inputValue
     
     'Input Validation
     
     
     
End If

Sub GetUserInputValidation()

Dim userInputRange As Range



'ignore error in the macro
On Error Resume Next
'check if a range was input
Set userInputRange = Application.InputBox(Prompt:="Select a Cell", Type:=8)
userInputRange.Value = "Hi"

'put hi into user selected cell
MsgBox "hi"
If userInputRange Is Nothing Then
    MsgBox "no range entered"
    Exit Sub
    End If


'inputValue = Application.InputBox("Please enter a number.", "Enter a number", , , , , , 1)
'inputValue = Application.InputBox("Prompt:=", Title:="Enter a Number", Type:=1)

'output the user input
'MsgBox inputValue



On Error GoTo 0
End Sub

