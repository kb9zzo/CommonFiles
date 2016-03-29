Attribute VB_Name = "modEmailDatabaseLink"
'---------------------------------------------------------------------------------------
' Module          : modEmailDatabaseLink
' Author          : Seth Schrock
' Date            : 4/24/2014
' Purpose         : This module contains the code necessary to add a link to an Outlook email that when clicked,
'                   it will take you to a specified form and to the requested record on that form.
' Instructions    :
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Function GetLink(Display As String, ScriptPath As String) As String
10    On Error GoTo Error_Handler

      Dim strLink As String

20    strLink = "<a href=""" & ScriptPath & """>" & Display & "</a>"
30    GetLink = strLink

Exit_Procedure:
40        Exit Function

Error_Handler:
50        If TempVars("RunMode") = "Test" Then
60            Call ErrorMessage(Err.Number, Err.Description, "modEmailDatabaseLink: GetLink")
70        Else
80            TSCs_ReportUnexpectedError "GetLink", "modEmailDatabaseLink", "Custom info"
90        End If
100       Resume Exit_Procedure
110       Resume
          
End Function

Public Function GetLinkInfo(FormName As String, PK As String, Optional Delimiter As String = "##") As String
10    On Error GoTo Error_Handler

      Dim strInfo As String

20    strInfo = "<font color=""white"">" & Delimiter & FormName & Delimiter & PK & "</font>"
30    GetLinkInfo = strInfo

Exit_Procedure:
40        Exit Function

Error_Handler:
50        If TempVars("RunMode") = "Test" Then
60            Call ErrorMessage(Err.Number, Err.Description, "modEmailDatabaseLink: GetLinkInfo")
70        Else
80            TSCs_ReportUnexpectedError "GetLinkInfo", "modEmailDatabaseLink", "Custom info"
90        End If
100       Resume Exit_Procedure
110       Resume
          
End Function

Public Function OpenEmailForm() As Boolean
Dim strEmail As String
Dim strForm As String
Dim strPK As String

strEmail = GetEmailText
strForm = GetFormName(strEmail)
strPK = GetPK(strEmail)

DoCmd.OpenForm FormName:=strForm, WhereCondition:="WireID=" & strPK

End Function

Private Function GetEmailText() As String

      Dim ol As Outlook.Application
      Dim olMail As Outlook.MailItem

10    On Error GoTo Error_Handler

      'Set ol = New Outlook.Application
20    Set ol = GetObject(, "Outlook.Application")

30    If TypeName(ol.ActiveWindow) = "Inspector" Then
40        Set olMail = ol.ActiveInspector.CurrentItem
50    Else
60        Set olMail = ol.ActiveExplorer.Selection(ol.ActiveExplorer.Selection.Count)
70    End If

80    GetEmailText = olMail.Body

Exit_Procedure:
90        Set ol = Nothing
100       Set olMail = Nothing
110       Exit Function

Error_Handler:
120       If TempVars("RunMode") = "Test" Then
130           Call ErrorMessage(Err.Number, Err.Description, "modEmailDatabaseLink: GetEmailText")
140       Else
150           TSCs_ReportUnexpectedError "GetEmailText", "modEmailDatabaseLink", "Custom info"
160       End If
170       Resume Exit_Procedure
180       Resume
          

End Function

Private Function GetFormName(strEmail As String, Optional Delimiter As String = "##") As String
10    On Error GoTo Error_Handler
      Dim strInfo() As String

20    strInfo = Split(strEmail, Delimiter)
30    GetFormName = strInfo(1)

Exit_Procedure:
40        Exit Function

Error_Handler:
50        If TempVars("RunMode") = "Test" Then
60            Call ErrorMessage(Err.Number, Err.Description, "modEmailDatabaseLink: GetFormName")
70        Else
80            TSCs_ReportUnexpectedError "GetFormName", "modEmailDatabaseLink", "Custom info"
90        End If
100       Resume Exit_Procedure
110       Resume

End Function

Private Function GetPK(strEmail As String, Optional Delimiter As String = "##") As String
10    On Error GoTo Error_Handler
      Dim strInfo() As String

20    strInfo = Split(strEmail, Delimiter)
30    GetPK = strInfo(2)

Exit_Procedure:
40        Exit Function

Error_Handler:
50        If TempVars("RunMode") = "Test" Then
60            Call ErrorMessage(Err.Number, Err.Description, "modEmailDatabaseLink: GetPK")
70        Else
80            TSCs_ReportUnexpectedError "GetPK", "modEmailDatabaseLink", "Custom info"
90        End If
100       Resume Exit_Procedure
110       Resume
          

End Function


