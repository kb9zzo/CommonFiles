Attribute VB_Name = "modBackend"
Option Compare Database
Option Explicit

Public Const g_Backend = 1

Public Function AttachDSNLessTable(LocalTable As String, RemoteTable As String, Location As Long) As Boolean
10    On Error GoTo Error_Handler

      Dim td As TableDef
      Dim strCon As String
      Dim db As DAO.Database

20    Set db = CurrentDb

30    On Error Resume Next
40    db.TableDefs.Delete LocalTable
              
50    On Error GoTo Error_Handler

60    strCon = "ODBC;" & GetLocation(Location) & ";Trusted_Connection=YES"

70    Set td = db.CreateTableDef(LocalTable, dbAttachSavePWD, RemoteTable, strCon)
80    db.TableDefs.Append td

90    AttachDSNLessTable = True

Exit_Procedure:
100       Set db = Nothing
110       Set td = Nothing
          
120       Exit Function

Error_Handler:
130       AttachDSNLessTable = False
          
140       If TempVars("RunMode") = "Test" Then
150           Call ErrorMessage(Err.Number, Err.Description, "modBackend: AttachDSNLessTable")
160       Else
170           TSCs_ReportUnexpectedError "AttachDSNLessTable", "modBackend", "Custom info"
180       End If
190       Resume Exit_Procedure
200       Resume
    
End Function

Public Function AddTable(TableName As String, RemoteTable As String, Location As Long) As Boolean
10    On Error GoTo Error_Handler

      Dim db As DAO.Database
      Dim td As TableDef
      Dim strCon As String
      Dim strAddQry As String

20    Set db = CurrentDb
30    strCon = "ODBC;" & GetLocation(Location) & "Trusted_Connection=YES"
40    Set td = db.CreateTableDef(TableName, dbAttachSavePWD, RemoteTable, strCon)

50    db.TableDefs.Append td

60    strAddQry = "INSERT INTO tblTableLocation (TableName, RemoteTableName, LocationID_fk) " & _
                  "VALUES ('" & TableName & "', '" & RemoteTable & "', " & Location & ")"
                  
70    db.Execute strAddQry, dbFailOnError

80    AddTable = True

Exit_Procedure:
90        Set db = Nothing
100       Set td = Nothing
110       Exit Function

Error_Handler:
120       AddTable = False
          
130       If TempVars("RunMode") = "Test" Then
140           Call ErrorMessage(Err.Number, Err.Description, "modBackend: AddTable")
150       Else
160           TSCs_ReportUnexpectedError "AddTable", "modBackend", "Custom info"
170       End If
180       Resume Exit_Procedure
190       Resume
          
End Function

Private Function GetLocation(LocationID As Long) As String
On Error GoTo Error_Handler

Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim strRst As String
Dim strLoc As String

strRst = "SELECT * FROM tblBE WHERE BE_ID_pk = " & LocationID

Set db = CurrentDb
Set rst = db.OpenRecordset(strRst, dbOpenDynaset)

With rst
    strLoc = "DRIVER=" & !Driver & ";SERVER=" & !Server & ";DATABASE=" & !DatabaseName & ";"
End With

GetLocation = strLoc

Exit_Procedure:
    On Error Resume Next
    rst.Close
    Set db = Nothing
    Set rst = Nothing
    
    Exit Function

Error_Handler:
    If TempVars("RunMode") = "Test" Then
        Call ErrorMessage(Err.Number, Err.Description, "modBackend: GetLocation")
    Else
        TSCs_ReportUnexpectedError "GetLocation", "modBackend", "Custom info"
    End If
    Resume Exit_Procedure
    Resume
    
End Function
