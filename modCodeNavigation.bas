Attribute VB_Name = "modCodeNavigation"
Option Compare Database
Option Explicit

'Requires the following references
'  Microsoft VBScript Regular Expressions 5.5
'  Microsoft Visual Basic for Applications Extensibility 5.3

Public Sub FindProcedure(strInput As String)
'---------------------------------------------------------------------------------------
' Procedure : FindProcedure
' Author    : sschrock
' Date      : 8/12/2015
' Purpose   : Navigate to a given procedure based on the error message provided
'             through Smiley's error hander
' Requirements: Microsoft VBScript Regular Expressions 5.5 reference
'---------------------------------------------------------------------------------------
'
    'Declare Variables
    Dim regex As New RegExp
    Dim colMatches As MatchCollection
    Dim strModule As String
    Dim strProcedure As String
    Dim lngStartLine As Long
             
    'Setup Regex Pattern
    '  Return string(s) between square brackets
    With regex
        .Global = True
        .Pattern = "\[([^\]]+)\]"
    End With
    
    'Perform matching and collect the matches
    Set colMatches = regex.Execute(strInput)
    
    'Get the first and second matches from the input string
    With colMatches
        strProcedure = .Item(0).submatches.Item(0)
        strModule = .Item(1).submatches.Item(0)
    End With

    
    With Application.VBE.ActiveVBProject.VBComponents(strModule).CodeModule
        'Show the proper module
        .codepane.Show
        
        'Go to the first line of the procedure
        lngStartLine = .ProcStartLine(strProcedure, vbext_pk_Proc)
        .codepane.SetSelection lngStartLine, 1, lngStartLine, 1
    End With
    
    'Clean up
    Set regex = Nothing
    Set colMatches = Nothing
    
End Sub

