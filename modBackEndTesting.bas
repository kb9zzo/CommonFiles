Attribute VB_Name = "modBackEndTesting"
Option Compare Database
Option Explicit

Public Function Exist(strFile As String, _
                      Optional intAttrib As Integer = vbReadOnly Or _
                                                      vbHidden Or _
                                                      vbSystem) As Boolean
                                                      
On Error Resume Next
Exist = (Dir(PathName:=strFile, Attributes:=intAttrib) <> "")
           

End Function

Public Sub ReLink(ByVal strDBName As String, _
                  Optional ByVal strFolder As String = "")
    Dim intParam As Integer, intErrNo As Integer
    Dim strOldLink As String, strOldName As String
    Dim strNewLink As String, strMsg As String
    Dim varLinkAry As Variant
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
 
    Set db = CurrentDb()
    If strFolder = "" Then strFolder = CurrentProject.Path
    If Right(strFolder, 1) = "\" Then _
        strFolder = Left(strFolder, Len(strFolder) - 1)
    strNewLink = strFolder & "\" & strDBName
    For Each tdf In db.TableDefs
        With tdf
            If .Attributes And dbAttachedTable Then
                varLinkAry = Split(.Connect, ";")
                For intParam = LBound(varLinkAry) To UBound(varLinkAry)
                    If Left(varLinkAry(intParam), 9) = "DATABASE=" Then Exit For
                Next intParam
                strOldLink = Mid(varLinkAry(intParam), 10)
                If strOldLink <> strNewLink Then
                    strOldName = Split(strOldLink, _
                                       "\")(UBound(Split(strOldLink, "\")))
                    If strOldName = strDBName Then
                        varLinkAry(intParam) = "DATABASE=" & strNewLink
                        .Connect = Join(varLinkAry, ";")
                        On Error Resume Next
                        Call .RefreshLink
                        intErrNo = Err.Number
                        On Error GoTo 0
                        Select Case intErrNo
                        Case 3011, 3024, 3044, 3055, 7874
                            varLinkAry(intParam) = "DATABASE=" & strOldLink
                            .Connect = Join(varLinkAry, ";")
                            strMsg = "Database file (%F) not found.%L" & _
                                     "Unable to ReLink [%T]."
                            strMsg = Replace(strMsg, "%F", strNewLink)
                            strMsg = Replace(strMsg, "%L", vbCrLf)
                            strMsg = Replace(strMsg, "%T", .Name)
                            Call MsgBox(Prompt:=strMsg, _
                                        Buttons:=vbExclamation Or vbOKOnly, _
                                        Title:="ReLink")
                            If intErrNo = 3024 _
                            Or intErrNo = 3044 _
                            Or intErrNo = 3055 Then Exit For
                        Case Else
                            strMsg = "[%T] relinked to ""%F"""
                            strMsg = Replace(strMsg, "%T", .Name)
                            strMsg = Replace(strMsg, "%F", strNewLink)
                            Debug.Print strMsg
                        End Select
                    End If
                End If
            End If
        End With
    Next tdf
End Sub

Public Function Connected(strTblName As String) As Boolean
Dim db As DAO.Database
Dim rsTable As DAO.Recordset

On Error GoTo Connected_Err

Set db = CurrentDb
Set rsTable = db.TableDefs(strTblName).OpenRecordset()
Set rsTable = db.OpenRecordset(strDBName)

Connected = True

Connected_Exit:
    Set rsTable = Nothing
    Set db = Nothing
    Exit Function
 
Connected_Err:
    Connected = False
    Resume Connected_Exit
    
End Function

