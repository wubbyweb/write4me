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

Public Sub GenerateAffirmation()
    On Error GoTo ErrorHandler

    ' Validate configuration on startup
    If Not HasValidApiConfig() Then
        MsgBox "API Configuration not found or invalid", vbCritical
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

        ' Get response from AI
        Dim aiResponse As String
        aiResponse = GetAIResponse(emailBody, promptText)

        ' Insert at the beginning if we got a response
        If Len(aiResponse) > 0 Then
            ' Create the new content with simple HTML formatting
            Dim newContent As String
            newContent = CreateFormattedResponse(aiResponse)

            ' Set the HTMLBody with combined content
            objItem.htmlBody = newContent & objItem.htmlBody
        End If
    End If

    Exit Sub
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

Private Function CreateFormattedResponse(response As String) As String
    CreateFormattedResponse = "<div style='font-family: Calibri; font-size: 11pt;'>" & _
                            Replace(response, vbCrLf, "<br>") & _
                            "<br><br>" & _
                            "<hr>" & _
                            "<br></div>"
End Function

Private Function CleanTextForJson(text As String) As String
    Dim cleanText As String
    cleanText = text

    ' Escape backslashes first
    cleanText = Replace(cleanText, "\", "\\")

    ' Escape quotes
    cleanText = Replace(cleanText, """", "\""")

    ' Handle different types of line breaks
    cleanText = Replace(cleanText, vbCrLf, "\n")
    cleanText = Replace(cleanText, vbCr, "\n")
    cleanText = Replace(cleanText, vbLf, "\n")

    ' Escape tabs
    cleanText = Replace(cleanText, vbTab, "\t")

    ' Handle other special characters
    cleanText = Replace(cleanText, "/", "\/")

    ' Remove any control characters
    Dim i As Long
    Dim char As String
    Dim result As String

    For i = 1 To Len(cleanText)
        char = Mid(cleanText, i, 1)
        If AscW(char) >= 32 Or char = "\n" Or char = "\t" Then
            result = result & char
        End If
    Next i

    CleanTextForJson = result
End Function

Private Function GetAIResponse(originalText As String, promptText As String) As String
    On Error GoTo ErrorHandler
    
    Dim apiConfig As apiConfig
    apiConfig = GetActiveApiConfig()
    
    If Len(apiConfig.ApiKey) = 0 Then
        MsgBox "No valid API configuration found", vbCritical
        Exit Function
    End If
    
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Clean and escape the input text properly
    originalText = CleanTextForJson(originalText)
    promptText = CleanTextForJson(promptText)
    
    ' Prepare API request
    xmlhttp.Open "POST", apiConfig.ApiEndpoint, False
    
    ' Set headers based on API type
    Select Case apiConfig.ApiType
        Case "openai"
            xmlhttp.setRequestHeader "Content-Type", "application/json"
            xmlhttp.setRequestHeader "Authorization", "Bearer " & apiConfig.ApiKey
            
            ' Create OpenAI formatted JSON request
            Dim openAiRequest As String
            openAiRequest = "{" & _
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
            
            xmlhttp.Send openAiRequest
            
        Case "anthropic"
            xmlhttp.setRequestHeader "Content-Type", "application/json"
            xmlhttp.setRequestHeader "x-api-key", apiConfig.ApiKey
            xmlhttp.setRequestHeader "anthropic-version", "2023-06-01"
            
            ' Create Anthropic formatted JSON request
            Dim anthropicRequest As String
            anthropicRequest = "{" & _
                """model"": ""claude-3-5-sonnet-20241022""," & _
                """max_tokens"": 2000," & _
                """messages"": [" & _
                    "{" & _
                        """role"": ""user""," & _
                        """content"": ""System: " & promptText & "\n\nUser: " & originalText & """" & _
                    "}" & _
                "]" & _
            "}"
            
            xmlhttp.Send anthropicRequest
    End Select
    
    ' Enhanced response handling
    If xmlhttp.Status = 200 Then
        Dim responseText As String
        responseText = xmlhttp.responseText
        Debug.Print "Response: " & responseText
        
        ' Parse response based on API type
        Select Case apiConfig.ApiType
            Case "openai"
                Dim openAiStart As Long, openAiEnd As Long
                openAiStart = InStr(responseText, """content"": """)
                If openAiStart > 0 Then
                    openAiStart = openAiStart + 12
                    openAiEnd = InStr(openAiStart, responseText, """")
                    If openAiEnd > 0 Then
                        GetAIResponse = Mid(responseText, openAiStart, openAiEnd - openAiStart)
                    End If
                End If
                
            Case "anthropic"
                Dim anthropicStart As Long, anthropicEnd As Long
                anthropicStart = InStr(responseText, """text"":""")
                If anthropicStart > 0 Then
                    anthropicStart = anthropicStart + 11
                    anthropicEnd = InStr(anthropicStart, responseText, """")
                    If anthropicEnd > 0 Then
                        GetAIResponse = Mid(responseText, anthropicStart, anthropicEnd - anthropicStart)
                    End If
                End If
        End Select
        
        ' Unescape special characters
        GetAIResponse = Replace(GetAIResponse, "\n", vbNewLine)
        GetAIResponse = Replace(GetAIResponse, "\""", """")
    Else
        Dim errorMsg As String
        errorMsg = "API request failed with status: " & xmlhttp.Status & vbNewLine & _
                  "Response: " & xmlhttp.responseText
        Debug.Print errorMsg
        MsgBox errorMsg, vbCritical
        GetAIResponse = ""
    End If
    Exit Function
    
ErrorHandler:
    Debug.Print "Error: " & Err.Description
    MsgBox "Error calling API: " & Err.Description, vbCritical
    GetAIResponse = ""
End Function

