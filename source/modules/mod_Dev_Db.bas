Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Db
' Level:        Framework module
' Version:      1.00
' Description:  Db functions & procedures
'
' Source/date:  Bonnie Campbell, 12/5/2017
' Revisions:    BLC, 12/5/2017 - 1.00 - initial version
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------
' ---------------------------------
' SUB:          CopyTable
' Description:  Remove the primary key for an existing table
' Assumptions:  -
' Parameters:   tbl - name of table (string)
'               tblNew - name of new table (string)
'               IncludeData - if data should be included (boolean, optional, default = false)
'                             true = include data, false = structure only
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017  - initial version
' ---------------------------------
Public Sub CopyTable(tbl As String, tblNew As String, Optional IncludeData = False)
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim StructureOrData As Boolean  'True = structure only, False = structure & data
    
    StructureOrData = Not (IncludeData)
    
    Set db = CurrentDb

    If IsObject(db.TableDefs(tbl)) Then
        DoCmd.TransferDatabase acExport, "Microsoft Access", CurrentDb.Name, _
            acTable, tbl, tblNew, StructureOnly:=StructureOrData
    End If
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3265 'item not found in this collection
        MsgBox "Sorry, the table '" & tbl & "' doesn't exist in this database." _
            & vbCrLf & vbCrLf & "Check the table name and try again." _
            & vbCrLf & vbCrLf & "Context: (#" & Err.Number & " - CopyTable[mod_Dev_Db])", vbCritical, _
            "Missing Table"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CopyTable[mod_Dev_Db])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RemovePrimaryKey
' Description:  Remove the primary key for an existing table &
'               change it to a Long vs. Autonumber
' Assumptions:  -
' Parameters:   tbl - name of table (string)
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:  Bonnie Campbell, December 5, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 12/5/2017  - initial version
' ---------------------------------
Public Sub RemovePrimaryKey(tbl As String)
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim tdf As TableDef
    
    Set db = CurrentDb

    If IsObject(db.TableDefs(tbl)) Then _
        db.TableDefs(tbl).Indexes.Delete "PrimaryKey"
    
    'change to Number vs. Autonumber
    DoCmd.RunSQL "ALTER TABLE " & tbl & " " & _
                    "ALTER COLUMN ID Long"
    
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case 3265 'item not found in this collection
        MsgBox "Sorry, the table '" & tbl & "' doesn't exist in this database." _
            & vbCrLf & vbCrLf & "Check the table name and try again." _
            & vbCrLf & vbCrLf & "Context: (#" & Err.Number & " - CopyTable[mod_Dev_Db])", vbCritical, _
            "Missing Table"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemovePrimaryKey[mod_Dev_Db])"
    End Select
    Resume Exit_Handler
End Sub

'-------------------------------
' Test Functions
'-------------------------------

Public Function TestCopy(t1 As String, t2 As String)

    CopyTable t1, t2
    
    RemovePrimaryKey t2
End Function