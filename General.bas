Attribute VB_Name = "General"

Public Function GetProgramname() As String
'returns the Program Name
GetProgramname = GetSetting(App.Title, "Main", "PPName", "Code Editor")
End Function

Public Sub ShowError(INum As Long, sDescr As String)
'Show an error message to the user
MsgBox "Error: " & INum & vbCrLf & sDescr, vbCritical, GetProgramname
End Sub
Public Sub ShowInfo(sDescr As String)
'Show the user an info box
MsgBox sDescr, vbInformation, GetProgramname
End Sub


