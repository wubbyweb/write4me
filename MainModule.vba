' Main Module
Option Explicit

Private globalRibbon As IRibbonUI

' Ribbon callback for customUI.onLoad
Public Sub Ribbon_Load(ribbon As IRibbonUI)
    Set globalRibbon = ribbon
End Sub

' Main function that will be triggered by the button
Public Sub RewriteEmail()
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
        Dim emailBody As String
        emailBody = objItem.Body
        
        Dim promptText As String
        promptText = "You are an email editor. Rewrite the following email with no spelling or grammar errors and make it more formal while maintaining the core message."
        ' Get formal version from OpenAI
        Dim formalEmail As String
        formalEmail = GetFormalVersionFromOpenAI(emailBody, promptText)

        ' Update email body if we got a response
        If Len(formalEmail) > 0 Then
              objItem.Body = formalEmail & vbCrLf & vbCrLf & "--------------------------------------------------------------------------------" & vbCrLf & emailBody
        End If
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' Main function that will be triggered by the button
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
        
        Dim promptText As String
        promptText = "You are an email editor. Read the content and generate an affirmation response to the email, keep the response short."
        
        ' Get formal version from OpenAI
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

