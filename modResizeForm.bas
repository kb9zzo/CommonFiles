Attribute VB_Name = "modResizeForm"
Option Compare Database
Option Explicit

Public Sub ResizeForm(frm As Form, Optional RecordCount As Integer = 10, Optional AddRecord As Boolean = False)
   'Variables
      Dim lngCount As Long
      Dim lngWindowHeight As Long
      Dim lngOldWindowHeight As Long
      Dim lngDeltaTop As Long
      Dim rs As DAO.Recordset
 
   'Find the amount of records in form
      Set rs = frm.RecordsetClone
      If Not rs.EOF Then rs.MoveLast
      lngCount = rs.RecordCount
      
      
   'If AddRecord is true then add 1 record to the number of records to
   ' show line for new record
      If AddRecord Then
         lngCount = lngCount + 1
      End If
 
   'Check whether there are more then Max amount of records
      If lngCount > RecordCount Then
         lngCount = RecordCount
         'Enable vertical scrollbar
            frm.ScrollBars = 2 'Vertical
      Else
         'Disable scrollbars
         frm.ScrollBars = 0 'None
      End If
 
   'Calculate new windowheight.
   'If you do not have a header/footer, or they are not visible adjust accordingly
      lngWindowHeight = frm.FormHeader.Height + _
                  frm.Detail.Height * lngCount + _
                  frm.FormFooter.Height + _
                  905
                  'The 567 is to account for title bar Height.
                  'If you use thin border, adjust accordingly
 
   'The form will be "shortened" and we need to adjust the top property as well to keep it properly centered
      lngOldWindowHeight = frm.WindowHeight
      lngDeltaTop = (lngOldWindowHeight - lngWindowHeight) / 2
 
   'Use the move function to move the form
 
      'frm.Move frm.WindowLeft, frm.WindowTop + lngDeltaTop, , lngWindowHeight
      frm.Move 50, 50, , lngWindowHeight

   'Cleanup
      Set rs = Nothing


End Sub

