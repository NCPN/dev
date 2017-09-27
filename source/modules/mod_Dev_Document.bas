Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Documentation
' Level:        Development module
' Version:      1.02
'
' Description:  Debugging related functions & procedures for database documentation
'
' Source/date:  Bonnie Campbell, January 24, 2017
' Revisions:    BLC - 1/24/2017 - 1.00 - initial version
'               BLC - 9/13/2017 - 1.01 - added ReferenceProperties()
' -------------------------------------------------------------------------------
'               BLC - 9/27/2017 - 1.02 - moved to NCPN_dev
' =================================

' ---------------------------------
' FUNCTION:     GetDescriptions
' Description:  Returns table descriptions
' Assumptions:  -
' Parameters:   db - name of database (string)
' Returns:      descriptions - table descriptions (string)
' Throws:       none
' References:   -
' Source/date:
' http://databases.aspfaq.com/schema-tutorials/schema-how-do-i-show-the-description-property-of-a-column.html
'
' http://stackoverflow.com/questions/17555174/how-to-loop-through-all-tables-in-an-ms-access-db
' Allen Browne,
' http://allenbrowne.com/func-06.html
' Adapted:      Bonnie Campbell, February 13, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/13/2015 - initial version
' ---------------------------------
Public Function GetDescriptions(db As String) As String
On Error GoTo Err_Handler
    
    Dim Catalog As AccessObject
    Dim dsc As String
    Dim tbl As AccessObject
    Dim TableDefs As Collection '??
    
    Set Catalog = CreateObject("ADOX.Catalog")
    Catalog.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=\" & db & ""
        '"Data Source=<path>\<file>.mdb"
 
    'iterate through tables, then table columns to retrieve descriptions
    For Each tbl In Catalog.Tables
        Debug.Print tbl.Name
    Next
 
    dsc = Catalog.Tables("table_name").Columns("column_name").Properties("Description").Value
 
    For Each tbl In TableDefs
        Debug.Print tbl.Name
    Next
    
    GetDescriptions = dsc
 
 '   If Err.Number <> 0 Then
  '      Response.Write "&lt;" & Err.Description & "&gt;"
   ' Else
   '     Response.Write "Description = " & dsc
   ' End If
    Set Catalog = Nothing

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescriptions[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     TableInfo
' Description:  Display field names, types, sizes, and descriptons for a table.
' Assumptions:  -
' Parameters:   tbl - name of table (string)
'               OutputToFile - whether to output info to text file (optional, boolean)
' Returns:      table info (string, if OutputToFile = false string is empty)
' Throws:       none
' References:   -
' Source/date:
'   Allen Browne, April 2010
'   http://allenbrowne.com/func-06.html
'   HansUp, March 29, 2010
'   http://stackoverflow.com/questions/2536955/retrieve-list-of-indexes-in-an-access-database
'   Microsoft
'   https://msdn.microsoft.com/en-us/library/bb243210(v=office.12).aspx
'   Ben, July 16, 2012
'   http://stackoverflow.com/questions/11503174/how-to-create-and-write-to-a-txt-file-using-vba
' Adapted:      Bonnie Campbell, February 13, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/13/2015 - initial version
'   BLC - 1/24/2017 - adjusted to allow output to text file, fixed error handling
' ---------------------------------
Public Function TableInfo(tbl As String, Optional OutputToFile = False, _
            Optional strPath As String = "db_documentation_") As String
On Error GoTo Err_Handler
   
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.field
    Dim idx As DAO.index
    
    Set db = CurrentDb()
    Set tdf = db.TableDefs(tbl)
   
    'retrieve # records
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset(tdf.Name, dbOpenDynaset)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
    End If
      
    Dim strText As String
    strText = ""
      
    If OutputToFile Then
                
        'table info (tdf.RecordCount = -1 for all linked tables)
        strText = "======  ==============================   ==========" _
                & vbCrLf & "TABLE    DESCRIPTION                      LINKED?" _
                & vbCrLf & "======  ==============================   ==========" _
                & vbCrLf & tdf.Name & " " & IIf(Len(tdf.Properties("Description")) > 0, _
                  tdf.Properties("Description"), "") _
                & "  " & IIf(tdf.RecordCount = -1, "x", "") _
                & vbCrLf & "# RECORDS:     " & rs.RecordCount _
                & vbCrLf
        
        'index info
        If tdf.Indexes.Count > 0 Then
            strText = strText & vbCrLf & "==========  =============  ===========" _
                & vbCrLf & "INDEX NAME   INDEX FIELDS    PRIMARY    " _
                & vbCrLf & "==========  =============  ==========="
            For Each idx In tdf.Indexes
                strText = strText & vbCrLf & idx.Name & "  " _
                    & idx.Fields & "  " _
                    & IIf(idx.Primary, "x", "")
            Next idx
            strText = strText & vbCrLf
        End If
       
        'field info
        strText = strText & "==========   ====   ==========   ====   ===========" _
                & vbCrLf & "FIELD NAME   REQD    FIELD TYPE  SIZE   DESCRIPTION" _
                & vbCrLf & "==========   ====   ==========   ====   ==========="
    
        For Each fld In tdf.Fields
            strText = strText & vbCrLf & fld.Name & "  " & _
                    IIf(fld.Required, "x", "") & "   " & _
                    FieldTypeName(fld) & "    " & _
                    fld.Size & "    " & _
                    fld.Properties("Description")
        Next
        strText = strText & vbCrLf & "==========    ====    ==========    ====    ===========" _
                & vbCrLf & vbCrLf

    Else
      
        'table info (tdf.RecordCount = -1 for all linked tables)
        Debug.Print "TABLE", "DESCRIPTION", "LINKED?"
        Debug.Print "======", "==============================", "=========="
        Debug.Print tdf.Name, IIf(Len(tdf.Properties("Description")), _
                    tdf.Properties("Description"), ""), _
                    IIf(tdf.RecordCount = -1, "x", "")
        Debug.Print "# RECORDS:     " & rs.RecordCount
        Debug.Print vbCrLf
        
        'index info
        If tdf.Indexes.Count > 0 Then
            Debug.Print "INDEX NAME", "INDEX FIELDS", "PRIMARY    "
            Debug.Print "==========", "=============", "==========="
            For Each idx In tdf.Indexes
                Debug.Print idx.Name, idx.Fields, IIf(idx.Primary, "x", "")
            Next idx
            Debug.Print vbCrLf
        End If
       
        'field info
        Debug.Print "FIELD NAME", "REQD", "FIELD TYPE", "SIZE", "DESCRIPTION"
        Debug.Print "==========", "====", "==========", "====", "==========="
    
        For Each fld In tdf.Fields
            Debug.Print fld.Name,
            Debug.Print IIf(fld.Required, "x", ""),
            Debug.Print FieldTypeName(fld),
            Debug.Print fld.Size,
            Debug.Print fld.Properties("Description")
            'Debug.Print GetDescrip(fld)
        Next
        Debug.Print "==========", "====", "==========", "====", "==========="
        Debug.Print vbCrLf
    
    End If
    
    TableInfo = strText

Exit_Handler:
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case 3265 'table name invalid
        MsgBox tbl & " table doesn't exist", vbCritical, _
            "Error encountered (#" & Err.Number & " - TableInfo[mod_Dev_Document])"
      Case 3270 'property not found -> ignore & move on
        Resume Next
'        MsgBox tbl & " property doesn't exist", vbCritical, _
'            "Error encountered (#" & Err.Number & " - TableInfo[mod_Dev_Document])"
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - TableInfo[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetDescrip
' Description:  Returns table descriptions
' Assumptions:  -
' Parameters:   obj - database object
' Returns:      description - object description (string)
' Throws:       none
' References:   -
' Source/date:
' Allen Browne, April 2010
' http://allenbrowne.com/func-06.html
' Adapted:      Bonnie Campbell, February 13, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/13/2015 - initial version
' ---------------------------------
Public Function GetDescrip(obj As Object) As String
On Error GoTo Err_Handler

    GetDescrip = obj.Properties("Description")

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescrip[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          DocumentDb
' Description:  Prepare database documentation
' Assumptions:  -
' Parameters:   tbl - include tables (boolean)
'               idx - include indexes (boolean)
'               fld - include fields (boolean)
'
' Returns:      text
' Throws:       none
' References:   -
' Source/date:
'   HansUp, July 9, 2013
'   http://stackoverflow.com/questions/17555174/how-to-loop-through-all-tables-in-an-ms-access-db
' Adapted:      Bonnie Campbell, January 24, 2017 - for NCPN tools
' Revisions:
'   BLC - 1/24/2017 - initial version
' ---------------------------------
Public Function DocumentDb() 'Optional tbl = True, Optional idx = True, Optional fld = True)
On Error GoTo Err_Handler

    Dim strText As String, strPath As String
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb
    
    'retrieve database info
    ' CurrentDb.Name = path & name
    ' use Application.CurrentProject.Name for db name
    strText = "***************************************" _
        & vbCrLf & "*  " & Application.CurrentProject.Name _
        & vbCrLf & "***************************************" & vbCrLf
    
    'retrieve table info
    For Each tdf In db.TableDefs
        ' ignore system and temporary tables
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*") Then
           strText = strText & TableInfo(tdf.Name, True)
           Debug.Print tdf.Name & "..."
        End If
    Next
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    
    strPath = CurrentProject.Path & "\db_doc_" & strPath & Format(Now(), "YYYYmmDD_HHMMss") & ".txt"
    
    Set oFile = FSO.CreateTextFile(strPath)
    oFile.WriteLine strText
    oFile.Close
    
    Set FSO = Nothing
    Set oFile = Nothing
    
    
    Debug.Print "DONE"
    
Exit_Handler:
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DocumentDb[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          DbProperties
' Description:  Print database properties
' Assumptions:  -
' Parameters:   -
' Returns:      text
' Throws:       none
' References:   -
' Source/date:
'   NeoPa, January 11, 2012
'   https://bytes.com/topic/access/insights/929840-database-properties
' Adapted:      Bonnie Campbell, January 24, 2017 - for NCPN tools
' Revisions:
'   BLC - 1/24/2017 - initial version
' ---------------------------------
Public Sub DbProperties()
On Error GoTo Err_Handler

    Dim prop As Property
    For Each prop In CurrentDb.Properties
        Debug.Print prop.Name
    Next

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescrip[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ReferenceProperties
' Description:  Print (debug) reference properties
' Assumptions:  -
' Parameters:   -
' Returns:      text
' Throws:       none
' References:   -
' Source/date:
'   John Austin (Microsoft), June 12, 2017
'   https://msdn.microsoft.com/en-us/vba/access-vba/articles/reference-guid-property-access
' Adapted:      Bonnie Campbell, September 13, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/13/2017 - initial version
' ---------------------------------
Public Function ReferenceProperties()
On Error GoTo Err_Handler

     Dim ref As Reference

    ' Enumerate through References collection.
    For Each ref In References
       ' Check IsBroken property.
       If ref.IsBroken = False Then
          Debug.Print "Name: ", ref.Name
          Debug.Print "FullPath: ", ref.FullPath
         ' Debug.Print "Version: ", ref.Major &; "." &; ref.Minor
         Debug.Print "GUID: ", ref.GUID
       Else
          Debug.Print "GUIDs of broken references:"
          Debug.Print ref.GUID
       End If
    Next ref

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescrip[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function