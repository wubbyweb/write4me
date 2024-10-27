' ConfigManager Module
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
    configPath = "\config.ini"

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
