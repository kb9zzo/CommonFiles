VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : clsMessage
' Author    : sschrock
' Date      : 11/7/2013
' Purpose   : This class module will be passed a series of text strings and then will
'             return the collection as a single string will proper comma placements
'             as well as any additional words needed to make it grammatically correct.
'
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private intUpper As Integer
Private intLower As Integer
Private arrText() As String



Private Sub class_initialize()

intUpper = 0
ReDim arrText(0) As String

End Sub

Public Sub AddText(Text As String)

intUpper = intUpper + 1
ReDim Preserve arrText(intUpper) As String
arrText(intUpper) = Text

End Sub

Public Function Message(Optional addPeriod As Boolean = False) As String
Dim i As Integer
Dim strMsg As String

Select Case intUpper
    Case 1
        strMsg = arrText(1)
        
    Case 2
        strMsg = arrText(1) & " and " & arrText(2)
        
    Case Else
        For i = 1 To intUpper
            Select Case i
                Case 1
                    strMsg = arrText(i)
                    
                Case Is <= (intUpper - 1)
                    strMsg = strMsg & ", " & arrText(i)
                    
                Case Is = intUpper
                    strMsg = strMsg & " and " & arrText(i)
            
            End Select
        Next

End Select

If addPeriod Then
    strMsg = strMsg & "."
End If

Message = strMsg

End Function

Public Property Get ItemCount() As Integer
    ItemCount = intUpper
End Property
