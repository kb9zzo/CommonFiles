Attribute VB_Name = "modDatabaseVersion"
Option Compare Database
Option Explicit

Public Sub CreateProperty(PropName As String, intType As Integer, PropValue As String)
Dim db As DAO.Database
Dim p As Property
Set db = DBEngine(0)(0)
Set p = db.CreateProperty(PropName, intType, PropValue)
db.Properties.Append p
End Sub

Public Sub DeleteProperty(PropName As String)
Dim db As DAO.Database
Dim p As Property
Set db = DBEngine(0)(0)
Set p = db.Properties(PropName)
db.Properties.Delete p.Name
End Sub

Public Sub ChangeVersion(VersionNum As String)
DBEngine(0)(0).Properties("DatabaseVersion") = VersionNum
Debug.Print "New version is: " & DBEngine(0)(0).Properties("DatabaseVersion")

End Sub
