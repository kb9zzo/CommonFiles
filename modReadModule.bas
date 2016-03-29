Attribute VB_Name = "modReadModule"
Option Compare Database
Option Explicit

Function ReadWholeLine(strModuleName As String, strProcName As String, strText As String) As String
 Dim mdl As Module, lngNumLines As Long
 Dim lngSLine As Long, lngSCol As Long
 Dim lngELine As Long, lngECol As Long
 Dim strTemp As String
 
 On Error GoTo Error_DeleteWholeLine
 
 DoCmd.OpenModule strModuleName
 Set mdl = Modules(strModuleName)
 
 lngSLine = ProcStartLine(strModuleName, strProcName)
 
 If mdl.Find(strText, lngSLine, lngSCol, lngELine, lngECol, True) Then
    lngNumLines = Abs(lngELine - lngSLine) + 1
    strTemp = LTrim$(mdl.Lines(lngSLine, lngNumLines))
    strTemp = RTrim$(strTemp)
    Do While Right(strTemp, 1) = "_"
        lngSLine = lngSLine + 1
        
        strTemp = strTemp & vbNewLine & _
                  Trim(mdl.Lines(lngSLine, 1))
                  
    Loop
    
    ReadWholeLine = strTemp
 Else
    ReadWholeLine = ""
 End If
  
Exit_DeleteWholeLine:
 Set mdl = Nothing
 Exit Function
 
Error_DeleteWholeLine:
 ReadWholeLine = ""
 Resume Exit_DeleteWholeLine
End Function

Function ProcStartLine(strModuleName As String, strProcName As String) As Long
Dim mdl As Module

On Error GoTo Error_Handler

Set mdl = Modules(strModuleName)

ProcStartLine = mdl.ProcStartLine(strProcName, vbext_pk_Proc)

Exit_Proc:
    Set mdl = Nothing
    Exit Function
    
Error_Handler:
    ProcStartLine = 0
    Resume Exit_Proc
End Function
