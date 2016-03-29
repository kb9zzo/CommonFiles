Attribute VB_Name = "modEmails"
Option Compare Database
Option Explicit

Public Sub SendEmail(ByVal SendTo As String, ByVal Message As String, ByVal Subject As String, Optional ByVal AttachmentFile As String)
      Dim objOutlook As Object
      Dim objMsg As Object
      Dim objAttach As Object

10    On Error GoTo Error_Handler

20    Set objOutlook = CreateObject("Outlook.Application")
30    Set objMsg = objOutlook.CreateItem(0)

40    With objMsg
50        .To = SendTo
          '.From = GetUserEmail
60        .Subject = Subject
70        .Body = Message
                  
80        If AttachmentFile <> "" Then
90            Set objAttach = .Attachments.Add(AttachmentFile)
100       End If
          
110       .Send
          
120   End With


Exit_Procedure:
130       Set objOutlook = Nothing
140       Set objMsg = Nothing
150       Set objAttach = Nothing
          
160       Exit Sub

Error_Handler:
170       If TempVars("Runmode") = "Test" Then
180           Call ErrorMessage(Err.Number, Err.Description, "modEmails: SendEmail")
190       Else
200           TSCs_ReportUnexpectedError "SendEmail", "modEmails", "Custom info"
210       End If
220       Resume Exit_Procedure
230       Resume
          

End Sub

Public Function SendCDOEmail(ByVal SendTo As String, ByVal Message As String, ByVal Subject As String, Optional ByVal Attach As String) As Boolean
10    On Error GoTo Error_Handler

      Dim cdoConfig As CDO.Configuration
      Dim cdoMsg As CDO.Message

20    Set cdoMsg = New CDO.Message
30    Set cdoConfig = New CDO.Configuration

40    With cdoConfig.Fields
50        .Item(cdoPrefix & "sendusing") = conCDOSendUsingPort
60        .Item(cdoPrefix & "smtpserver") = conStrSmtpServer
70        .Item(cdoPrefix & "smtpserverport") = conCdoSmtpServerPort
80        .Update
90    End With


100   Set cdoMsg = CreateObject("CDO.Message")
110   Set cdoMsg.Configuration = cdoConfig


120   With cdoMsg
130       .To = SendTo
140       .From = "DatabaseNotifications@fountaintrust.com"
150       .Subject = Subject
160       .TextBody = Message
          
170       If Attach <> "" Then
180           .AddAttachment Attach
190       End If
          
200       .Send
210   End With

220   SendCDOEmail = True

Exit_Procedure:
230       On Error Resume Next
240       Set cdoMsg = Nothing
250       Set cdoConfig = Nothing
          
260       Exit Function

Error_Handler:
270       SendCDOEmail = False
280       If TempVars("RunMode") = "Test" Then
290           Call ErrorMessage(Err.Number, Err.Description, "modEmails: SendCDOEmail")
300       Else
310           TSCs_ReportUnexpectedError "SendCDOEmail", "modEmails", "Custom info"
320       End If
330       Resume Exit_Procedure
340       Resume
          
End Function
