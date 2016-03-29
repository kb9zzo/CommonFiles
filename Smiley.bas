Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit

Public Sub TSCs_ReportUnexpectedError(Optional strProcName As String, Optional strModuleName As String, Optional strCustomInfo As String)
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Procedure : TSCs_ReportUnexpectedError
' Author    : AEC - Anders Ebro Christensen / TheSmileyCoder
' Date      : 2013-02-02
' Version   : 1.0
' Purpose   : Record as much info as possible about what went wrong and email it.
' Bugs?     : Email: SmileyCoderTools@gmail.com
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
   'Variables
      Dim lngErrNr As Long
      Dim lngLineNumber As Long
      Dim strErrDescription As String
      ReDim strScreenShots(0 To 0)
      Dim errX As DAO.Error
      Dim strCode As String
      
   'First record basic details on the error
        If Errors.Count > 1 Then
            lngErrNr = Err.Number
            For Each errX In DAO.Errors
                strErrDescription = strErrDescription & "; Error Number " & errX.Number & _
                                    ", Description " & errX.Description
            Next errX
        Else
            lngErrNr = Err.Number
            strErrDescription = Err.Description
        End If
      
        lngLineNumber = Erl
        
        strCode = ReadWholeLine(strModuleName, strProcName, lngLineNumber & "  ")
      
   'VERY IMPORTANT Do not start err handler before the above 3 lines.
   '               Otherwise the err object will be reset, and error information lost.
      On Error GoTo Err_Handler
   
   'Set hourglass
      DoCmd.Hourglass True


   'Compose custom message (Will be displayed to user eventually)
      mStrInfoForUser = "Error " & lngErrNr & " has occured"
         If strProcName <> "" Then
            mStrInfoForUser = mStrInfoForUser & " in procedure [" & strProcName & "]"
         End If
         If strModuleName <> "" Then
            mStrInfoForUser = mStrInfoForUser & " in module [" & strModuleName & "]"
         End If
         If lngLineNumber <> 0 Then 'If line numbers not used, ErrLine will be 0
            mStrInfoForUser = mStrInfoForUser & " on line [" & lngLineNumber & "] " & _
                    "on the following code:" & vbCrLf & vbCrLf & strCode & vbCrLf
         End If
      
      mStrInfoForUser = mStrInfoForUser & vbNewLine & strErrDescription
      
      If strCustomInfo <> "" Then
         mStrInfoForUser = mStrInfoForUser & vbNewLine & vbNewLine & strCustomInfo
      End If
      

      
      'Get a GUID for the error log(s)
         mTSC_GUID = TSCf_GetNewGUID()
      
      'Create a folder for storing screenshots and files
         mStrFolder = Environ("Temp") & "\TSC_ErrorReport\" & Format$(Date, "yyyy-mm-dd") & " " & mTSC_GUID & "\"
         TSCf_MakeDir mStrFolder
         
   
      'Open the table for logging
         Call TSCs_OpenLogTable
         
   'Log main info. Note that all these functions also log to the text file at the same time.
   'The text file is also a way to make sure error gots logged even if error 3048 occurs.
         mTSC_strLog = "Error ID:" & mTSC_GUID & vbNewLine
         TSCs_WriteToLog "mem_CustomMessage", "Error Message as presented to user", mStrInfoForUser
         TSCs_WriteToLog "lng_ErrNumber", "Error Number", lngErrNr
         TSCs_WriteToLog "lng_ErrLine", "Error Line (0 if N/A)", lngLineNumber
         TSCs_WriteToLog "tx_ErrDescription", "Error Description", strErrDescription
         TSCs_WriteToLog "tx_ErrorInModule", "Error in Module", strModuleName
         TSCs_WriteToLog "tx_ErrorInProcedure", "Error in Procedure", strProcName
         TSCs_WriteToLog "mem_SessionInfo", "Session Information", TSCf_CollectSessionInfo
         TSCs_WriteToLog "dt_DateTime", "Time of Error", Now()
         TSCs_WriteToLog "tx_UserName", "Windows User Name", TSCf_GetWindowsLoginName
         TSCs_WriteToLog "tx_ActiveForm", "Active Form", TSCf_getActiveFormName
         TSCs_WriteToLog "tx_ActiveControl", "Active Control", TSCf_getActiveControlName
         TSCs_WriteToLog "tx_ActiveControlsForm", "Active Control Parent Form", TSCf_getActiveControlsParentForm
         TSCs_WriteToLog "tx_ActiveDataSheet", "Active Data Sheet", TSCf_getActiveDatasheetName
         TSCs_WriteToLog "tx_ActiveReport", "Active Report", TSCf_getActiveReportName
         TSCs_WriteToLog "mem_OpenForms", "All Open Forms", TSCf_GetListOfOpenForms
         TSCs_WriteToLog "tx_AppVersion", "Application Version", TSCf_GetAppVersion
         TSCs_WriteToLog "tx_TSC_ErrRepVersion", "TheSmileyCoders Error Report Tool Version Info", gErrorReportToolVersion
         TSCs_WriteToLog "tx_AppVersion", "Application Version", TSCf_GetAppVersion
         TSCs_WriteToLog "lng_MinutesRunning", "Minutes application has been running", TSCf_MinutesAppHasBeenRunning
         TSCs_WriteToLog "lng_MinutesSinceLastBoot", "Hours since last Windows reboot", TSCf_HoursSinceLastWindowsBoot
         TSCs_WriteToLog "mem_DatabaseProperties", "List Of database properties", TSCf_CollectDatabaseProperties
         TSCs_WriteToLog "tx_Internetconnection", "Internet Connection", TSCf_CollectInternetConnectionInfo
         TSCs_WriteToLog "mem_EnvironmentInfo", "Environ Variables", TSCf_CollectEnvironmentInfo
         TSCs_WriteToLog "mem_BackendInfo", "Backend Connection Information", TSCf_CollectBackendInformation
         'TSCs_WriteToLog "mem_WMI", "WMI Info", TSCf_CollectWMIInfo 'Useless to most cases, so not included as default
     
   'Collect screenshot
      If gTakeScreenshot Then
         mStrScreenShots = TSCf_CaptureAllWindows(mTSC_GUID)
      End If
      
   'Close the log file (or it will be lost if an error occurs)
      Call TSCs_CloseLogTable
      
   'Open submittal form as dialog (code will pause here, but there is code in the load event of the form)
      If lngErrNr = 3048 And TSCf_TableAndQueryObjectsAvailable = 0 Then
         'The dreaded error "3048 - Cannot open any more databases"
         'This error prevents us from openening any more table or query objects.
         'Since access stores form information in tables, we cannot even open a unbound form.
         'Therefore we must stick to a simple msgbox
         MsgBox mStrInfoForUser, vbOKOnly, gAppName & ": Cannot open any more forms"
      Else
         DoCmd.OpenForm "TSC_ErrRep_frm_SubmitError", acNormal, , , , acDialog
      End If

      

Exit_Sub:
   'Cleanup
      Call TSCs_CloseLogTable
      DoCmd.Hourglass False
   'Leave
      On Error GoTo 0
      Exit Sub

Err_Handler:
   
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleUnexpectedError of Module vbTSC_ErrReporter"
   Call TSCs_CloseSubmitform 'Make sure form is closed/cleaned
   Resume Exit_Sub
   'Not used resume. Good for debugging
   Resume

End Sub

