' ConfigManager.bas module
Option Explicit

Private Type ConfigSettings
    ApiKey As String
    ApiEndpoint As String
End Type

Private Config As ConfigSettings

Private Function ReadConfigFile() As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim txtFile As Object
    Dim configPath As String
    
    ' Get config file path relative to the Excel file location
    configPath = ThisWorkbook.Path & "\config.ini"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(configPath) Then
        MsgBox "Configuration file not found at: " & configPath, vbCritical
        ReadConfigFile = False
        Exit Function
    End If
    
    Set txtFile = fso.OpenTextFile(configPath, 1) ' 1 = ForReading
    
    ' Read and parse config file
    While Not txtFile.AtEndOfStream
        Dim line As String
        line = txtFile.ReadLine
        
        If InStr(line, "=") > 0 Then
            Dim parts() As String
            parts = Split(line, "=")
            
            Select Case Trim(parts(0))
                Case "OPENAI_API_KEY"
                    Config.ApiKey = Trim(parts(1))
                Case "API_ENDPOINT"
                    Config.ApiEndpoint = Trim(parts(1))
            End Select
        End If
    Wend
    
    txtFile.Close
    
    ' Validate configuration
    If Len(Config.ApiKey) = 0 Or Len(Config.ApiEndpoint) = 0 Then
        MsgBox "Invalid configuration: Missing required settings", vbCritical
        ReadConfigFile = False
        Exit Function
    End If
    
    ReadConfigFile = True
    Exit Function

ErrorHandler:
    MsgBox "Error reading configuration: " & Err.Description, vbCritical
    ReadConfigFile = False
End Function

Public Function GetApiKey() As String
    If Len(Config.ApiKey) = 0 Then
        If Not ReadConfigFile() Then
            Exit Function
        End If
    End If
    GetApiKey = Config.ApiKey
End Function

Public Function GetApiEndpoint() As String
    If Len(Config.ApiEndpoint) = 0 Then
        If Not ReadConfigFile() Then
            Exit Function
        End If
    End If
    GetApiEndpoint = Config.ApiEndpoint
End Function

' Main module (write4me.vba)
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
        
        ' Get formal version from OpenAI
        Dim formalEmail As String
        formalEmail = GetFormalVersionFromOpenAI(emailBody)
        
        ' Update email body if we got a response
        If Len(formalEmail) > 0 Then
            objItem.Body = formalEmail
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' Function to call OpenAI API
Private Function GetFormalVersionFromOpenAI(originalText As String) As String
    On Error GoTo ErrorHandler
    
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Prepare API request
    xmlhttp.Open "POST", GetApiEndpoint(), False
    xmlhttp.setRequestHeader "Content-Type", "application/json"
    xmlhttp.setRequestHeader "Authorization", "Bearer " & GetApiKey()
    
    ' Prepare request body
    Dim requestBody As String
    requestBody = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""system"", ""content"": ""You are a professional email editor. Rewrite the following email in a formal, professional tone while maintaining the core message.""}, {""role"": ""user"", ""content"": """ & Replace(originalText, """", "\""") & """}]}"
    
    ' Send request
    xmlhttp.send requestBody
    
    ' Check for successful response
    If xmlhttp.Status = 200 Then
        ' Parse response
        Dim responseText As String
        responseText = xmlhttp.responseText
        
        ' Extract content from JSON response (basic parsing)
        Dim startPos As Long
        Dim endPos As Long
        startPos = InStr(responseText, """content"":""") + 11
        endPos = InStr(startPos, responseText, """")
        
        GetFormalVersionFromOpenAI = Mid(responseText, startPos, endPos - startPos)
    Else
        MsgBox "API request failed with status: " & xmlhttp.Status, vbCritical
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error calling OpenAI API: " & Err.Description, vbCritical
    GetFormalVersionFromOpenAI = ""
End Function