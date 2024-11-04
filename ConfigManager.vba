' ConfigManager.vba
Public Type apiConfig
    ApiKey As String
    ApiType As String  ' "openai" or "anthropic"
    ApiEndpoint As String
End Type

Private Function GetConfigValue(key As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim configPath As String
    configPath = "C:\Users\Raj\appdata\Roaming\Microsoft\Outlook\config.ini"
    
    ' Check if config file exists
    If Not fso.FileExists(configPath) Then
        GetConfigValue = ""
        Exit Function
    End If
    
    ' Read config file
    Dim fileNum As Integer
    fileNum = FreeFile
    
    Open configPath For Input As #fileNum
    
    Dim line As String
    Do Until EOF(fileNum)
        Line Input #fileNum, line
        
        ' Skip comments and empty lines
        If Left(Trim(line), 1) <> "#" And Len(Trim(line)) > 0 Then
            Dim parts() As String
            parts = Split(line, "=")
            
            If UBound(parts) = 1 Then
                If Trim(parts(0)) = key Then
                    GetConfigValue = Trim(parts(1))
                    Close #fileNum
                    Exit Function
                End If
            End If
        End If
    Loop
    
    Close #fileNum
    GetConfigValue = ""
End Function

Public Function GetActiveApiConfig() As apiConfig
    Dim config As apiConfig
    
    ' Try OpenAI first
    config.ApiKey = Trim(GetConfigValue("OPENAI_API_KEY"))
    If Len(config.ApiKey) > 0 Then
        config.ApiType = "openai"
        config.ApiEndpoint = GetConfigValue("OPENAI_API_ENDPOINT")
        If Len(config.ApiEndpoint) = 0 Then
            config.ApiEndpoint = "https://api.openai.com/v1/chat/completions"
        End If
        GetActiveApiConfig = config
        Exit Function
    End If
    
    ' Try Anthropic if OpenAI not configured
    config.ApiKey = Trim(GetConfigValue("ANTHROPIC_API_KEY"))
    If Len(config.ApiKey) > 0 Then
        config.ApiType = "anthropic"
        config.ApiEndpoint = GetConfigValue("ANTHROPIC_API_ENDPOINT")
        If Len(config.ApiEndpoint) = 0 Then
            config.ApiEndpoint = "https://api.anthropic.com/v1/messages"
        End If
        GetActiveApiConfig = config
        Exit Function
    End If
    
    ' No valid API keys found
    config.ApiKey = ""
    config.ApiType = ""
    config.ApiEndpoint = ""
    GetActiveApiConfig = config
End Function

' Helper function to check if we have a valid API configuration
Public Function HasValidApiConfig() As Boolean
    Dim config As apiConfig
    config = GetActiveApiConfig()
    HasValidApiConfig = Len(config.ApiKey) > 0
End Function
