VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private ctlCollection As New collection
Private lngKey As Long

Public Sub AddControl(Control As Control, ControlName As String)
    ctlCollection.Add Control, ControlName
End Sub

Public Sub Permissions(Level As Integer)
Dim ctl As Control

Select Case Level
    Case 1 'Hidden
        For Each ctl In ctlCollection
            ctl.Properties("Visible") = False
        Next ctl
        
    Case 2 'View
        For Each ctl In ctlCollection
            ctl.Properties("Visible") = True
            ctl.Properties("Enabled") = False
        Next ctl
        
    Case 3 'Edit
        For Each ctl In ctlCollection
            ctl.Properties("Visible") = True
            ctl.Properties("Enabled") = True
        Next ctl
    
End Select


'Clean up
Set ctl = Nothing

End Sub

Public Property Let Key(value As Long)
    lngKey = value
End Property

Public Property Get Key() As Long
    Key = lngKey
End Property
