

' Put this at the top of your regular code module (not in a function)
Private Type AffirmationSettings
    ToneStyle As String
    Length As String
End Type


Private Sub UserForm_Initialize()
    ' Set default values
    OptionFormal.Value = True
    OptionShort.Value = True
    Cancelled = True
End Sub


' Then the function
Private Function ShowAffirmationForm() As AffirmationSettings
    Dim frm As New AffirmationSettingsForm
    frm.Show
    
    Dim settings As AffirmationSettings
    If Not frm.Cancelled Then
        settings.ToneStyle = frm.ToneStyle
        settings.Length = frm.Length
    End If
    
    Unload frm
    ShowAffirmationForm = settings
End Function
' Function to generate prompt based on settings
Private Function GeneratePrompt(settings As AffirmationSettings) As String
    Dim prompt As String
    prompt = "You are an email editor. Generate an affirmation response to the email in a " & _
             settings.ToneStyle & " tone, making it " & settings.Length & ". "
    
    Select Case settings.ToneStyle
        Case "formal"
            prompt = prompt & "Use professional and respectful language. "
        Case "casual"
            prompt = prompt & "Use friendly, conversational language. "
        Case "humorous"
            prompt = prompt & "Include appropriate humor while maintaining positivity. "
    End Select
    
    Select Case settings.Length
        Case "short"
            prompt = prompt & "Keep the response concise and brief. "
        Case "long"
            prompt = prompt & "Provide a detailed and elaborate response. "
    End Select
    
    GeneratePrompt = prompt
End Function
' Modified main function
Public Sub GenerateAffirmation()
    On Error GoTo ErrorHandler
    
    ' Validate configuration on startup
    If Len(GetApiKey()) = 0 Then
        MsgBox "API Key not configured properly", vbCritical
        Exit Sub
    End If
    
    Dim objItem As Object
    Set objItem = Application.ActiveInspector.CurrentItem
    
    ' Check if we're in compose mode
    If objItem.Class = olMail Then
        ' Store the original content
        Dim emailBody As String
        emailBody = objItem.Body
        
        ' Get user preferences through form
        Dim settings As AffirmationSettings
        settings = ShowAffirmationForm()
        
        ' Generate custom prompt
        Dim promptText As String
        promptText = GeneratePrompt(settings)
        
        ' Get response from OpenAI
        Dim formalEmail As String
        formalEmail = GetFormalVersionFromOpenAI(emailBody, promptText)
        
        ' Insert at the beginning if we got a response
        If Len(formalEmail) > 0 Then
            ' Create the new content with simple HTML formatting
            Dim newContent As String
            newContent = "<div style='font-family: Calibri; font-size: 11pt;'>" & _
                        Replace(formalEmail, vbCrLf, "<br>") & _
                        "<br><br>" & _
                        "<hr>" & _
                        "<br></div>"
            
            ' Set the HTMLBody with combined content
            objItem.htmlBody = newContent & objItem.htmlBody
        End If
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub
' Helper function to format the response
Private Function CreateFormattedResponse(response As String) As String
    CreateFormattedResponse = "<div style='font-family: Calibri; font-size: 11pt;'>" & _
                            Replace(response, vbCrLf, "<br>") & _
                            "<br><br>" & _
                            "<hr>" & _
                            "<br></div>"
End Function

' Function to call OpenAI API
Private Function GetFormalVersionFromOpenAI(originalText As String, promptText As String) As String
    On Error GoTo ErrorHandler
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Clean and escape the input text properly
    originalText = Replace(originalText, """", "\""")    ' Escape quotes
    originalText = Replace(originalText, vbCrLf, "\n")   ' Handle line breaks
    originalText = Replace(originalText, vbCr, "\n")     ' Handle carriage returns
    originalText = Replace(originalText, vbLf, "\n")     ' Handle line feeds
    
    ' Prepare API request
    xmlhttp.Open "POST", GetApiEndpoint(), False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer " & Trim(GetApiKey())
    
    ' Create properly formatted JSON request
    Dim requestBody As String
    requestBody = "{" & _
        """model"": ""gpt-3.5-turbo""," & _
        """messages"": [" & _
            "{" & _
                """role"": ""system""," & _
                """content"": """ & promptText & """" & _
            "}," & _
            "{" & _
                """role"": ""user""," & _
                """content"": """ & originalText & """" & _
            "}" & _
        "]," & _
        """temperature"": 0.7," & _
        """max_tokens"": 2000" & _
    "}"
    ' Debug output - check the request
    Debug.Print "Request Body: " & requestBody
    
    ' Send request
    xmlhttp.Send requestBody
    
    ' Enhanced response handling
    If xmlhttp.Status = 200 Then
        Dim responseText As String
        responseText = xmlhttp.responseText
        Debug.Print "Response: " & responseText
        
        ' Better JSON parsing
        Dim startPos As Long, endPos As Long
        startPos = InStr(responseText, """content"": """)
        If startPos > 0 Then
            startPos = startPos + 12  ' Length of """content"": """
            endPos = InStr(startPos, responseText, """")
            If endPos > 0 Then
                GetFormalVersionFromOpenAI = Mid(responseText, startPos, endPos - startPos)
                ' Unescape special characters
                GetFormalVersionFromOpenAI = Replace(GetFormalVersionFromOpenAI, "\n", vbNewLine)
                GetFormalVersionFromOpenAI = Replace(GetFormalVersionFromOpenAI, "\""", """")
            End If
        End If
    Else
        Dim errorMsg As String
        errorMsg = "API request failed with status: " & xmlhttp.Status & vbNewLine & _
                  "Response: " & xmlhttp.responseText
        Debug.Print errorMsg
        MsgBox errorMsg, vbCritical
        GetFormalVersionFromOpenAI = ""
    End If
    Exit Function
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    MsgBox "Error calling OpenAI API: " & Err.Description, vbCritical
    GetFormalVersionFromOpenAI = ""
End Function
