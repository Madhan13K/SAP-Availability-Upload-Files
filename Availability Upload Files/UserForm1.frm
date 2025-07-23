VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConfigureParameters
   Caption         =   "UserForm1"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8430.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ConfigureParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private selectedDomain As String
Private authenticationUrl As String
Private clientId As String
Private clientSecret As String
Private isFormSubmitted As Boolean ' Flag to check if user submitted

' Function to show the form and return all values in one call
Public Function GetUserInputs(ByRef domain As String, ByRef authUrl As String, ByRef id As String, ByRef secret As String) As Boolean
    Me.Caption = "Enter Your Details"
    
    ' Show the form
    isFormSubmitted = False
    Me.Show
    
    ' Check if user clicked Submit or closed the form
    If isFormSubmitted Then
        domain = selectedDomain
        authUrl = authenticationUrl
        id = clientId
        secret = clientSecret
        GetUserInputs = True ' Return True if successfully submitted
    Else
        GetUserInputs = False ' Return False if user canceled
    End If
End Function

' Submit button click event (CommandButton1)
Private Sub CommandButton1_Click()
    ' Validate inputs
    If Me.OptionButton1.Value Then
        selectedDomain = "eu10"
    ElseIf Me.OptionButton2.Value Then
        selectedDomain = "us10"
    ElseIf Me.OptionButton3.Value Then
        selectedDomain = "eu11"
    Else
        MsgBox "Please select a domain.", vbExclamation, "Selection Required"
        Exit Sub
    End If

    authenticationUrl = Trim(Me.TextBox1.Value)
    clientId = Trim(Me.TextBox2.Value)
    clientSecret = Trim(Me.TextBox3.Value)

    ' Check if any field is empty
    If authenticationUrl = "" Or clientId = "" Or clientSecret = "" Then
        MsgBox "All fields are required. Please fill in all details.", vbExclamation, "Missing Information"
        Exit Sub
    End If

    ' Set the flag to indicate successful submission
    isFormSubmitted = True

    ' Close the form
    Me.Hide
End Sub

Private Sub Label2_Click()

End Sub
