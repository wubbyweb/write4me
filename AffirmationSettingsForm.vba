' In a new UserForm called AffirmationSettingsForm
Option Explicit
Public ToneStyle As String
Public Length As String
Public Cancelled As Boolean

Private Sub cmdOK_Click()
    ' Get tone selection
    If OptionFormal.Value Then
        ToneStyle = "formal"
    ElseIf OptionCasual.Value Then
        ToneStyle = "casual"
    ElseIf OptionHumorous.Value Then
        ToneStyle = "humorous"
    End If
    
    ' Get length selection
    If OptionShort.Value Then
        Length = "short"
    ElseIf OptionLong.Value Then
        Length = "long"
    End If
    
    Cancelled = False
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Cancelled = True
    Me.Hide
End Sub
