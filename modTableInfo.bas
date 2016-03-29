Attribute VB_Name = "modTableInfo"
Option Compare Database
Option Explicit

Public Enum gcnTypes
    gcnSimple = 1
    gcnComplex = 2
    gcnCopyTo = 3
    gcncopyfrom = 4
    gcnMine = 5
End Enum
 
'---------------------------------------------------------------------------------------
' Procedure : GetColumnNames
' Author    : Jim
' Date      : 4/28/2014
' Purpose   : Returns Table Field names in optional formats to the immediate window
'             Optional paramerters allow selective return of all tables and fields or wildcard matches
'             Return formats are
'               Field Name only (gcnSimple)
'               TableName.FieldName.Type (gcnComplex)
'               rs1!FieldName=x  (gcnCopyTo)
'               x=rs1!FieldName  (gcnCopyFrom)
'---------------------------------------------------------------------------------------
'
Public Sub GetColumnNames(ReplyType As gcnTypes, Optional TableName_str As String, Optional FieldPrefix_str As String)
' reply types are 1=simple (field names only)
'                 2=complex (table name, field name, field type)
'                 3=Move to
'                 4=Move from
'                 5=Mine (field names and data type
'
' returns data in the immediate window
    Dim tbl As DAO.TableDef
    Dim fld As DAO.Field
    Dim fldTypes(23) As String
    Dim fldTyp As Integer
    Dim fldDesc As String
 
   On Error GoTo GetColumnNames_Error
 
    fldTypes(1) = "Boolean"
    fldTypes(2) = "Byte"
    fldTypes(3) = "Integer"
    fldTypes(4) = "Long"
    fldTypes(5) = "Currency"
    fldTypes(6) = "Single"
    fldTypes(7) = "Double"
    fldTypes(8) = "Date"
    fldTypes(9) = "Binary"
    fldTypes(10) = "Text"
    fldTypes(11) = "Long Binary"
    fldTypes(12) = "Memo"
    fldTypes(13) = "Attachment" '101
    fldTypes(14) = "Complex Byte" '102
    fldTypes(15) = "Complex Integer"
    fldTypes(16) = "Complex Long"
    fldTypes(17) = "Complex Single"
    fldTypes(18) = "Complex Double"
    fldTypes(19) = "Complex GUID"
    fldTypes(20) = "Complex Decimal"
    fldTypes(21) = "Complex Text"  ' 109
    fldTypes(22) = "Other"
 
    ' Print the header.
   On Error GoTo GetColumnNames_Error
 
    ' Loop through all the table definitions.
    
    Debug.Print AddPadding("--FIELD NAME--", 50) & "--DATA TYPE--"
    For Each tbl In CurrentDb.TableDefs
        If Len(Nz(TableName_str)) = 0 Or (Left(tbl.Name, Len(TableName_str)) = TableName_str) Then
            Debug.Print "---------------------------------------"
            Debug.Print "Table Name: " & tbl.Name
            Debug.Print "---------------------------------------"
            For Each fld In tbl.Fields
                fldTyp = fld.Type
                If fldTyp > 100 And fldTyp <= 109 Then
                    fldTyp = fldTyp - 88 ' 101 becomes 13
                End If
                If fldTyp > 0 And fldTyp <= 22 Then
                    fldDesc = fldTypes(fldTyp)
                Else
                    fldDesc = fld.Type & " Other"
                End If
 
                        'only include fields matching the name requested
                If Len(Nz(FieldPrefix_str)) = 0 Or (Left(fld.Name, Len(FieldPrefix_str)) = FieldPrefix_str) Then
                     Select Case ReplyType
                        Case gcnSimple
                                    Debug.Print fld.Name
                        Case gcnComplex
                                 Debug.Print tbl.Name & "." & fld.Name & "." & fldDesc
                        Case gcnCopyTo
                                 Debug.Print "rs1!" & fld.Name & "= x"
                        Case gcncopyfrom
                                Debug.Print "x=rs1!" & fld.Name
                        Case gcnMine
                                'Debug.Print fld.Name & Chr(9) & "  >  " & Chr(9) & fldDesc
                                Debug.Print AddPadding(fld.Name, 50) & AddPadding(fldDesc, 20)
                     End Select
                End If
                Next fld
            End If
 
    Next tbl
 
   On Error GoTo 0
   Exit Sub
 
GetColumnNames_Error:
 
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetColumnNames of Module PublicCode_vb"
    Resume Next
 
   On Error GoTo 0
   Exit Sub
 
End Sub
 
 
Private Function AddPadding(Text As String, PadTo As Integer)

Text = Text & Space(PadTo - Len(Text))

AddPadding = Text

End Function

'Public Sub Test()
'Dim strMsg As String
'strMsg = "This form will not be editable after you save, close or add a new record. Do you want to continue?"
'If MsgBox(strMsg, vbYesNo) = vbYes Then
'    Me.AllowEdits = False
'Else
'    Me.AllowEdits = True
'End If
'
'End Sub

Public Sub ControlProperties()
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim fld As Field
Dim prp As Property

Set db = CurrentDb
Set rst = db.OpenRecordset("TicketSpecs", dbOpenDynaset)

Set fld = rst.Fields("Width")

For Each prp In fld.Properties
    Debug.Print prp.Name
Next prp

Set db = Nothing
rst.Close
Set rst = Nothing
Set fld = Nothing

End Sub
