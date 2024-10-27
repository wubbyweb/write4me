Option Explicit

' Constants for OpenAI API
Private Const API_KEY As String = "your-api-key-here"
Private Const API_ENDPOINT As String = "https://api.openai.com/v1/chat/completions"

' Ribbon callback for customUI.onLoad
Public Sub Ribbon_Load(ribbon As IRibbonUI)
    Set globalRibbon = ribbon
End Sub

' Main function that will be triggered by the button
Public Sub RewriteEmail()
    On Error GoTo ErrorHandler
    
    Dim objItem As Object
    Set objItem = Application.ActiveInspector.CurrentItem
    
    ' Check if we're in compose mode
    If objItem.Class = olMail Then
        Dim emailBody As String
        emailBody = objItem.Body
        
        ' Get formal version from OpenAI
        Dim formalEmail As String
        formalEmail = GetFormalVersionFromOpenAI(emailBody)
        
        ' Update email body
        objItem.Body = formalEmail
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

' Function to call OpenAI API
Private Function GetFormalVersionFromOpenAI(originalText As String) As String
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Prepare API request
    xmlhttp.Open "POST", API_ENDPOINT, False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer " & API_KEY
    
    ' Prepare request body
    Dim requestBody As String
    requestBody = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""system"", ""content"": ""You are a professional email editor. Rewrite the following email in a formal, professional tone while maintaining the core message.""}, {""role"": ""user"", ""content"": """ & Replace(originalText, """", "\""") & """}]}"
    
    ' Send request
    xmlhttp.send requestBody
    
    ' Parse response
    Dim responseText As String
    responseText = xmlhttp.responseText
    
    ' Extract content from JSON response (basic parsing)
    Dim startPos As Long
    Dim endPos As Long
    startPos = InStr(responseText, """content"":""") + 11
    endPos = InStr(startPos, responseText, """")
    
    GetFormalVersionFromOpenAI = Mid(responseText, startPos, endPos - startPos)
End Function