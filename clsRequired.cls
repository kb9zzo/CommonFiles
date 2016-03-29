VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRequired"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'**************************************************************************************************
'Author         : Seth Schrock
'Date           : 10/30/13
'Purpose        : Set the fore color of the labels connected to controls that are set
'                 to be required to red if they haven't been populated and back to
'                 black when they have been populated.
'
'Instructions  1: When calling the the CheckControls procedure, frm is the calling form's
'                 name.
'
'              2: TxtCboLstChkOpt is a flag variable that allows you to choose
'                 which types of controls to check if they are required.  The calling code just
'                 need to add up the control types listed as properties of this class.
'                 For example, to check textboxes and checkboxes, you would just pass the value
'                 .txt + .chk    Order doesn't matter.
'
'              3: The tag parameter allows you to specify what the tag property value is
'                 that specifies that control as being required.  By default the value
'                 is "req".
'
'**************************************************************************************************


'**************************************************************************************************
'SETUP CONSTANTS
'
'Error Color:
Const errColor As Long = 255
'
'Populated Color:
Const okColor As Long = 855309
'
'Control OK Border Color
Const okBorderColor As Long = 10921638
'
'Control Type Values
Const iTxt = 16
Const iCbo = 8
Const iLst = 4
Const iChk = 2
Const iOpt = 1
'
'**************************************************************************************************



'Stored Values
Dim blnCompleted As Boolean
Dim strMesssage As String

Private msg As New clsMessage


Private Sub Class_Initialize()
blnCompleted = True

End Sub

Public Sub CheckControls(frm As String, TxtCboLstChkOpt As Long, Optional tag As String = "req")
'On Error GoTo Error_Handler

Dim bTxt As Boolean
Dim bCbo As Boolean
Dim bLst As Boolean
Dim bChk As Boolean
Dim bOpt As Boolean
Dim ctl As Control


'Check which control types to check
bTxt = (TxtCboLstChkOpt And iTxt)
bCbo = (TxtCboLstChkOpt And iCbo)
bLst = (TxtCboLstChkOpt And iLst)
bChk = (TxtCboLstChkOpt And iChk)
bOpt = (TxtCboLstChkOpt And iOpt)


'Mark controls that are required and empty
 
For Each ctl In Forms(frm).Controls
    If ((bTxt And ctl.ControlType = acTextBox) Or (bCbo And ctl.ControlType = acComboBox) Or _
       (bLst And ctl.ControlType = acListBox) Or (bChk And ctl.ControlType = acCheckBox) Or _
       (bOpt And ctl.ControlType = acOptionGroup)) And ctl.Properties("Visible") = True Then
        
        With Forms(frm).Controls(ctl.Name)
        
            If .tag = tag Then
                
                '.Controls.Item(0) refers to the attached label
                If .Controls.Count > 0 Then
                    If (ctl.ControlType <> acCheckBox And .Value & "" <> "") Or _
                       (ctl.ControlType = acCheckBox And .Value = True) Then
                        .Controls.Item(0).ForeColor = okColor
                    Else
                        .Controls.Item(0).ForeColor = errColor
                        msg.AddText .Controls.Item(0).Caption
                        blnCompleted = False
                    End If
                Else
                    If (ctl.ControlType <> acCheckBox And .Value & "" <> "") Or _
                       (ctl.ControlType = acCheckBox And .Value = True) Then
                        .BorderColor = okBorderColor
                    Else
                        .BorderColor = errColor
                        msg.AddText .Name
                        blnCompleted = False
                    End If
                End If
            
            End If
           
        End With
        
    End If
 
Next ctl

strMesssage = msg.Message

Exit_Procedure:
    Exit Sub
    
Error_Handler:
    If Err.Number = 2467 Then Resume Next

End Sub

Public Property Get Completed() As Boolean
    Completed = blnCompleted
End Property

Public Function Message() As String
Dim strMsg As String

strMsg = "Please fill in the following field"
If msg.ItemCount > 1 Then
    strMsg = strMsg & "s"
End If
strMsg = strMsg & ": " & msg.Message

Message = strMsg

End Function

Public Property Get txt() As Integer
    txt = iTxt
End Property

Public Property Get cbo() As Integer
    cbo = iCbo
End Property

Public Property Get lst() As Integer
    lst = iLst
End Property

Public Property Get chk() As Integer
    chk = iChk
End Property

Public Property Get opt() As Integer
    opt = iOpt
End Property
