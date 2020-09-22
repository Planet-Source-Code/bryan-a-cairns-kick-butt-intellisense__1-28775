Attribute VB_Name = "mod_Itsworkzone"
'Main Database functions
Public Function CheckFile(sFile As String) As Boolean
'Does a file exist TRUE / FALSE
On Error Resume Next
If sFile = "" Then
CheckFile = False
Exit Function
End If
Dim Iret
Iret = Dir(sFile)
If Iret > "" Then
CheckFile = True
Else
If Iret = "" Then
CheckFile = False
End If
End If

End Function
Public Function GetDatabaseFilename() As String
Dim sTMP As String
sTMP = GetSetting(App.Title, "Main", "Profile", App.Path & "\default.mdb")
GetDatabaseFilename = sTMP
End Function


Public Function DoesTableExist(sTable As String) As Boolean
On Error GoTo EH
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long
Dim bFound As Boolean
If CheckFile(GetDatabaseFilename) = False Then
MsgBox "Could not find main database!", vbCritical, "Error"
DoesTableExist = False
Exit Function
End If
bFound = False
Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)
    For I = 0 To Ndb.TableDefs.Count - 1
        If LCase(sTable) = LCase(Ndb.TableDefs(I).Name) Then
            bFound = True
            Exit For
        End If
    Next I
Ndb.Close
DoesTableExist = bFound
Screen.MousePointer = 0
Exit Function
EH:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical, "Does Table Exists Query"
Exit Function
End Function

Public Sub ClearTable(sTable As String)
On Error GoTo EH
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long

If DoesTableExist(sTable) = False Then
MsgBox sTable & " does not exist!", vbCritical, "Clear Table"
Exit Sub
End If

Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)
Set RECndb = Ndb.OpenRecordset(sTable, dbOpenDynaset)
If RECndb.BOF = True And RECndb.EOF = True Then
Else
    Do While Not RECndb.EOF
    RECndb.Delete
    RECndb.MoveNext
    Loop
End If
RECndb.Close
Ndb.Close
Screen.MousePointer = 0

Exit Sub
EH:
Screen.MousePointer = 0
MsgBox Err.Description, vbCritical, "Clear Table"
Exit Sub
End Sub

Public Sub WZRepairDatabase()

On Error Resume Next
Dim L As Long
Dim Ndb As Database
Dim TNdb As TableDef
Dim FLDnssc As Field
Dim RECndb As Recordset
Dim wrkDefault As Workspace
Dim I As Long

Dim bFound As Boolean
If CheckFile(GetDatabaseFilename) = False Then
MsgBox "Could not find main database!", vbCritical, "Error"
Exit Sub
End If

Screen.MousePointer = 11
Set Ndb = OpenDatabase(GetDatabaseFilename)

    For I = 0 To Ndb.TableDefs.Count - 1
        For L = 0 To Ndb.TableDefs(I).Fields.Count - 1
            Ndb.TableDefs(I).Fields(L).AllowZeroLength = True
            Ndb.TableDefs(I).Fields(L).Required = False
        Next L

    Next I
Ndb.Close

Screen.MousePointer = 0

End Sub

''''''''''''''''''''''''

