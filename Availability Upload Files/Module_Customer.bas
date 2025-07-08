Attribute VB_Name = "Module3"
Option Explicit

Public Function PostWorkforceAvailability(ByVal startDate As String, _
                                        ByVal workAssignmentExternalId As String, _
                                        ByVal UUID As String, _
                                        ByVal plannedWorkingHours As String, _
                                        ByVal nonWorkingHours As String, _
                                        ByVal token As String, _
                                        ByVal host As String, _
                                        ByVal enviroment As String, _
                                        ByVal domain As String _
                                        ) As String
    Dim URL As String
    Dim BOUNDARY As String
    URL = "https://" & host & "." & enviroment & "." & domain & "/WorkforceAvailabilityService/v1/$batch"
    BOUNDARY = "batch_123"
    
    Dim http As Object
    Dim requestBody As String
    
    requestBody = "--batch_123" & vbLf & _
    "Content-Type: application/http" & vbLf & _
    "Content-Transfer-Encoding:binary" & vbLf & _
    vbLf & _
    "POST WorkforceAvailability HTTP/1.1" & vbLf & _
    "Content-Type: application/json" & vbLf & _
    vbLf & _
    "{" & vbLf & _
    "  ""workAssignmentID"": """ & workAssignmentExternalId & """," & vbLf & _
    "  ""availabilityDate"": """ & startDate & """," & vbLf & _
    "  ""workforcePerson_ID"": """ & UUID & """," & vbLf & _
    "  ""normalWorkingTime"": """ & plannedWorkingHours & """," & vbLf & _
    "  ""availabilitySupplements"": [" & vbLf & _
    "    {" & vbLf & _
    "      ""contribution"": """ & nonWorkingHours & """," & vbLf & _
    "      ""absenceApprovalStatus"": ""APPROVED""" & vbLf & _
    "    }" & vbLf & _
    "  ]," & vbLf & _
    "  ""availabilityIntervals"": []" & vbLf & _
    "}" & vbLf & _
    vbLf & _
    "--batch_123--"
    Debug.Print "Body:" & requestBody
    ' Create and configure HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", URL, False
    
    ' Set headers
    http.setRequestHeader "Content-Type", "multipart/mixed;boundary=" & BOUNDARY
    http.setRequestHeader "Authorization", "Bearer " & token
    
    ' Send request
    On Error Resume Next
    http.Send requestBody
    
    ' Error handling
    If Err.Number <> 0 Then
        PostWorkforceAvailability = "Error: " & Err.Description
        Exit Function
    End If
    On Error GoTo 0
    

    If InStr(1, http.responseText, "HTTP/1.1 201 Created", vbTextCompare) > 0 Then
        PostWorkforceAvailability = " Availability created"
    Else
        PostWorkforceAvailability = http.responseText
    End If
End Function
Function GetUUID(workForcePersonExternalId As String, token As String, host As String, enviroment As String, domain As String) As String
    Dim http As Object
    Dim URL As String
    Dim requestBody As String
    Dim responseText As String
    Dim jsonResponse As Object
    Dim id As String
    
    ' API Endpoint
    URL = "https://" & host & "." & enviroment & "." & domain & "/ProjectExperienceService/v1/$batch"
    Debug.Print URL
    ' Authentication Token (Replace with a valid token)

requestBody = "--request-separator" & vbLf & _
                 "Content-Type: application/http" & vbLf & _
                 "Content-Transfer-Encoding:binary" & vbLf & _
                 vbLf & _
                 "GET Profiles?$filter=workforcePersonExternalID%20eq%20'" & workForcePersonExternalId & "' HTTP/1.1" & vbLf & _
                 vbLf & _
                 "--request-separator--"
    
    
    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Open POST request
    http.Open "POST", URL, False
    
    ' Set headers
    http.setRequestHeader "Content-Type", "multipart/mixed;boundary=request-separator"
    http.setRequestHeader "Authorization", "Bearer " & token
    
    
    ' Send request
    On Error Resume Next
    http.Send requestBody
    
    
    If http.Status = 400 Then
        Debug.Print "Full Response Text:"
        Debug.Print http.responseText
    End If
    
    On Error GoTo 0
    
    ' Get response
    If http.Status = 200 Then
        responseText = http.responseText
        Debug.Print "Response: " & responseText
        
        ' Find ID in response text
        Dim idStart As Long
        Dim idEnd As Long
        
        idStart = InStr(responseText, """ID"":""") ' Find start of ID
        If idStart > 0 Then
            idStart = idStart + 6 ' Length of """ID"":"""
            idEnd = idStart + 35 ' UUID is always 36 characters long
            
            ' Extract the UUID
            id = Mid(responseText, idStart, 36)
            
            ' Verify it looks like a UUID (basic check)
            If Len(id) = 36 And InStr(id, "-") > 0 Then
                GetUUID = id
                Exit Function
            End If
        End If
        
        Debug.Print "Could not find valid UUID in response"
    Else
        Debug.Print "HTTP Error: " & http.Status & " - " & http.statusText
        Debug.Print "Response: " & http.responseText
    End If
    
    GetUUID = "Error: UUID not found"
End Function
Function Base64Encode(text As String) As String
    Dim arr() As Byte
    Dim objXML As Object, objNode As Object
    
    ' Convert text to byte array
    arr = StrConv(text, vbFromUnicode)
    
    ' Encode to Base64
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arr
    Base64Encode = Replace(objNode.text, vbLf, "")

    ' Cleanup
    Set objNode = Nothing
    Set objXML = Nothing
End Function

Function GetAccessToken(URL As String, username As String, password As String) As String
    Dim http As Object
    Dim body As String, responseText As String
    Dim accessToken As String
    Dim startPos As Integer, endPos As Integer

    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Set up request body
    body = "grant_type=client_credentials"

    ' Open HTTP request (POST)
    http.Open "POST", URL, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
    
    ' Send request
    http.Send body

    ' Check response status
    If http.Status = 200 Then
        responseText = http.responseText

        ' Find the position of "access_token":" in the response
        startPos = InStr(responseText, """access_token"":""")
        If startPos > 0 Then
            startPos = startPos + Len("""access_token"":""") ' Move past "access_token":"

            ' Find the next quote that ends the token
            endPos = InStr(startPos, responseText, """")
            If endPos > 0 Then
                accessToken = Mid(responseText, startPos, endPos - startPos)
            End If
        End If
    Else
        MsgBox "Error: " & http.Status & " - " & http.statusText, vbCritical
        accessToken = ""
    End If
    
    ' Return access token
    GetAccessToken = accessToken
End Function


Function FormatDate(dateValue As Variant) As String
    If IsDate(dateValue) Then
        FormatDate = Format(dateValue, "yyyy-mm-dd")
    Else
        FormatDate = ""
    End If
End Function

Function FormatTime(timeValue As Variant) As String
    If IsNumeric(timeValue) Then
        FormatTime = Format(timeValue, "00") & ":00"
    Else
        FormatTime = ""
    End If
End Function

Public Function GetUserInputs(ByRef domain As String, ByRef authUrl As String, ByRef id As String, ByRef secret As String) As Boolean
    Dim ws As Worksheet
    Dim domainInput As String
    Dim selectedDomain As String
    Dim authenticationUrl As String
    Dim clientId As String
    Dim clientSecret As String

    Set ws = ThisWorkbook.Sheets("InputForm") ' Change to your actual sheet name if different

    domainInput = LCase(Trim(ws.Range("B2").Value))
    authenticationUrl = Trim(ws.Range("B3").Value)
    clientId = Trim(ws.Range("B4").Value)
    clientSecret = Trim(ws.Range("B5").Value)

    ' Validate domain selection
    Select Case domainInput
        Case "eu10"
            selectedDomain = "eu10.hana.ondemand.com"
        Case "us10"
            selectedDomain = "us10.hana.ondemand.com"
        Case "eu11"
            selectedDomain = "eu11.hana.ondemand.com"
        Case Else
            MsgBox "Please select a valid domain: eu10, us10, or eu11.", vbExclamation, "Domain Required"
            GetUserInputs = False
            Exit Function
    End Select

    ' Validate inputs
    If authenticationUrl = "" Or clientId = "" Or clientSecret = "" Then
        MsgBox "All fields are required. Please fill in all details.", vbExclamation, "Missing Information"
        GetUserInputs = False
        Exit Function
    End If

    ' Assign values to output variables
    domain = selectedDomain
    authUrl = authenticationUrl
    id = clientId
    secret = clientSecret
    GetUserInputs = True
End Function

Public Sub PostAvailability()
    Dim ws As Worksheet
    Dim workForcePersonExternalId As String
    Dim UUID As String
    Dim startDate As String
    Dim plannedWorkingHours As String
    Dim nonWorkingHours As String
    Dim startDateValue As String
    Dim plannedWorkingHoursValue As String
    Dim nonWorkingHoursValue As String
    Dim host As String
    Dim enviroment As String
    Dim domain As String
    Set ws = ThisWorkbook.ActiveSheet
    ' Host are always same
    host = "resource-management-api-projectscloud"
    ' Enviroment are always same
    enviroment = "cfapps"
    ' domain AuthUrl ClientId ClientSecrate needs input
    Dim authUrl As String
    Dim clientId As String
    Dim clientSecret As String
    
   Dim success As Boolean

    ' Call the UserForm to get all values at once
    success = UserForm1.GetUserInputs(domain, authUrl, clientId, clientSecret)

    ' Check if user submitted successfully
    If success Then
        MsgBox "You selected the following details:" & vbCrLf & vbCrLf & _
               "Domain: " & domain & vbCrLf & _
               "Authentication URL: " & authUrl & vbCrLf & _
               "Client ID: " & clientId & vbCrLf & _
               "Client Secret: " & clientSecret, _
               vbInformation, "User Input Summary"
    Else
        MsgBox "Operation cancelled. No data was submitted.", vbExclamation, "Cancelled"
    End If
    
    'For the token
    Dim token As String
    
    token = GetAccessToken(authUrl, clientId, clientSecret)

    
    ' Start from row 2 (assuming row 1 has headers)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow
        Dim response As String
        workForcePersonExternalId = ws.Cells(i, "B").Value
        UUID = GetUUID(workForcePersonExternalId, token, host, enviroment, domain)
        startDateValue = ws.Cells(i, "H").Value
        startDate = FormatDate(startDateValue)
        plannedWorkingHoursValue = ws.Cells(i, "I").Value
        plannedWorkingHours = FormatTime(plannedWorkingHoursValue)
        nonWorkingHoursValue = ws.Cells(i, "J").Value
        nonWorkingHours = FormatTime(nonWorkingHoursValue)
        response = PostWorkforceAvailability(startDate, ws.Cells(i, "G").Value, UUID, plannedWorkingHours, nonWorkingHours, token, host, enviroment, domain)
        
        ws.Cells(1, "L").Value = "Response"
        ws.Cells(i, "L").Value = response
    Next i
End Sub


