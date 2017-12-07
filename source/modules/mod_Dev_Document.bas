Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Documentation
' Level:        Development module
' Version:      1.04
'
' Description:  Debugging related functions & procedures for database documentation
'
' Source/date:  Bonnie Campbell, January 24, 2017
' Revisions:    BLC - 1/24/2017 - 1.00 - initial version
'               BLC - 9/13/2017 - 1.01 - added ReferenceProperties()
' -------------------------------------------------------------------------------
'               BLC - 9/27/2017 - 1.02 - moved to NCPN_dev
'               BLC - 10/3/2017 - 1.03 - added ObjectType enum, GetObjectList(),
'                                        GetProjectDetails() documentation
'               BLC - 10/4/2017 - 1.04 - added module descriptions to GetProjectDetails()
' =================================

' ---------------------------------
' ENUM:     ObjectType
' Description:  Database objects
' Note: Since there are gaps in this enum so iteration as below is *not* possible
'
'           For N = ObjectType.[_Last] To ObjectType.[_First]
'                Debug.Print N
'           Next
'
' Source/date:
'   Daniel Pineault (CARDA Consultants Inc., www.cardaconsultants.com), June 12, 2010
'   https://www.devhut.net/2010/06/12/ms-access-listing-of-database-objects/
'   Chip Pearson (Pearson Software Consulting), March 12, 2008
'   http://www.cpearson.com/excel/Enums.aspx
' Adapted:      Bonnie Campbell, October 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 10/3/2017 - initial version
' ---------------------------------
Enum ObjectType
    [_First] = 1
    Tables_Local = 1
    Tables_Linked_ODBC = 4
    Tables_Linked = 6
    Queries = 5
    Forms = -32768
    Reports = -32764
    Macros = -32766
    Modules = -32761
    [_Last] = -32761
End Enum

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
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
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
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
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
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
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
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
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
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescrip[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     ReferenceProperties
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
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetDescrip[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetProjectDetails
' Description:  Creates a text file listing project modules, procedures, functions,
'               subroutines, and properties
' Assumptions:  -
' Parameters:   IncludeCounts - include line counts (optional, default false, boolean)
' Returns:      text
' Throws:       none
' References:   -
' Source/date:
'   Daniel Pineault (CARDA Consultants Inc., www.cardaconsultants.com), June 4, 2011
'   https://www.devhut.net/2011/06/04/vba-vbe-enumerate-modules-procedures-and-line-count/
'   Andre, October 20, 2017
'   https://stackoverflow.com/questions/46842684/how-to-print-module-descriptions-from-vba/46844242#46844242
' Adapted:      Bonnie Campbell, September 29, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/29/2017 - initial version
'   BLC - 10/3/2017 - adjusted header to include project description & date
'   BLC - 10/4/2017 - added module descriptions
'   BLC - 10/21/2017 - revised to print module
' ---------------------------------
Public Function GetProjectDetails(Optional IncludeCounts As Boolean = False)
On Error GoTo Err_Handler
    Dim vbProj                As VBIDE.VBProject
    Dim vbComp                As VBIDE.VBComponent
    Dim vbMod                 As VBIDE.CodeModule
    Dim pk                    As VBIDE.vbext_ProcKind
    Dim sProcName             As String
    Dim strFile               As String
    Dim iCounter              As Long
    Dim FileNumber            As Integer
    Dim bFileClosed           As Boolean
    Dim strKind               As String
    Dim description           As String
    Const vbNormalFocus = 1
    
    'documentation file
    strFile = Application.CurrentProject.Path & "\" & CurrentProject.Name & "_ProjectDetails.txt"
    If Len(Dir(strFile)) > 0 Then Kill strFile
    FileNumber = FreeFile                           'Get unused file number
    Open strFile For Append As #FileNumber          'Create file name
    
'    Print #FileNumber, String(80, "=")
'    Print #FileNumber, "VBA Project Name: " & Application.VBE.ActiveVBProject.Name
'    Print #FileNumber, "Description:      " & Application.VBE.ActiveVBProject.Description
'    Print #FileNumber, String(80, "-")
    Print #FileNumber, "Database:         " & Application.CurrentProject.Name
    Print #FileNumber, "Db Path:          " & Application.CurrentProject.Path
    Print #FileNumber, String(80, "=")
    Print #FileNumber, "As of:            " & Now()
    Print #FileNumber, String(80, "=")
    
'    Print #FileNumber, "Database: " & Application.CurrentProject.Name
'    Print #FileNumber, "Database Path: " & Application.CurrentProject.Path
'    Print #FileNumber, String(80, "*")
'    Print #FileNumber, String(80, "*")
    Print #FileNumber, ""
 
    'iterate through projects
    For Each vbProj In Application.VBE.VBProjects
        Print #FileNumber, String(80, "*")
        Print #FileNumber, "VBA Project Name: " & Application.VBE.ActiveVBProject.Name
        Print #FileNumber, "Description:      " & Application.VBE.ActiveVBProject.description
        Print #FileNumber, String(80, "*")
'        Print #FileNumber, "VBA Project Name: " & vbProj.Name
        
        'iterate through modules
        For Each vbComp In vbProj.VBComponents
            Set vbMod = vbComp.CodeModule
            
            'fetch description
            Select Case TypeName(vbMod)
                Case "CodeModule" 'module
'                    description = CurrentDb.Containers("Modules").Documents(vbMod).Properties("Description")
                    description = Nz(CurrentDb.Containers("Modules").Documents(vbComp.Name).Properties("Description"), "")
                Case Else
                    description = ""
            End Select
            
            Print #FileNumber, "   " & String(77, "-")
            Print #FileNumber, "   " & vbComp.Name & _
                IIf(IncludeCounts = True, " :: " & _
                    vbMod.CountOfLines & " total lines", "")
            Print #FileNumber, "   " & description
            Print #FileNumber, "   " & String(77, "-")
            
            iCounter = 1
    
            'iterate through procedures
            Do While iCounter < vbMod.CountOfLines
                sProcName = vbMod.ProcOfLine(iCounter, pk)
                
                Select Case pk
                    Case vbext_pk_Proc  'sub or function
                        strKind = "SUB / FUNCTION: "
                    Case vbext_pk_Get   'property get
                        strKind = "PROPERTY GET: "
                    Case vbext_pk_Let   'property let
                        strKind = "PROPERTY LET: "
                    Case vbext_pk_Set   'property set
                        strKind = "PROPERTY SET: "
                End Select
                
                If sProcName <> "" Then
                    Print #FileNumber, "      " & _
                        strKind & sProcName & _
                        IIf(IncludeCounts = True, " :: " & _
                            vbMod.ProcCountLines(sProcName, pk) & " lines", "")
                    
                    'print description
                    
                    'print parameters
                    
                    
                    iCounter = iCounter + vbMod.ProcCountLines(sProcName, pk)
                Else
                    iCounter = iCounter + 1
                End If
            Loop
            
'            Print #FileNumber, ""
        Next vbComp
    Next vbProj
    
    Close #FileNumber  'close file.
    bFileClosed = True
    
    'Open the generated text file
    Application.FollowHyperlink strFile
 
Exit_Handler:
    If bFileClosed = False Then Close #FileNumber   'close file
    If Not vbMod Is Nothing Then Set vbMod = Nothing
    If Not vbComp Is Nothing Then Set vbComp = Nothing
    If Not vbProj Is Nothing Then Set vbProj = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetProjectDetails[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetObjectList
' Description:  Creates a list of project objects (see also ObjectType enum above)
'
'   Tables_Local = 1            Queries = 5             Macros = -32766
'   Tables_Linked_ODBC = 4      Forms = -32768          Modules = -32761
'   Tables_Linked = 6           Reports = -32764
'
' Assumptions:  -
' Parameters:   ObjType - type of object to include (optional, default , long)
' Returns:      text
' Throws:       none
' References:   -
' Source/date:
'   Daniel Pineault (CARDA Consultants Inc., www.cardaconsultants.com), June 12, 2010
'   https://www.devhut.net/2010/06/12/ms-access-listing-of-database-objects/
'   Chip Pearson (Pearson Software Consulting), March 12, 2008
'   http://www.cpearson.com/excel/Enums.aspx
' Adapted:      Bonnie Campbell, October 3, 2017 - for NCPN tools
' Revisions:
'   BLC - 10/3/2017 - initial version
' ---------------------------------
Public Function GetObjectList(ExportToFile As Boolean, _
                             ObjType As Variant)
On Error GoTo Err_Handler
 
    Dim db          As DAO.Database
    Dim rs          As DAO.Recordset
    Dim strSQL      As String
    Dim strText     As String
    Dim strFile     As String
    Dim FileNumber  As Integer
    Dim strObjTypes As String
 
    'include all ObjectTypes
    strObjTypes = "(1, 4, 5, 6, -32761, -32764, -32766, -32768)"
 
    'defaults
    If ObjType = "ALL" Then
    
        'include all ObjectTypes
        strObjTypes = " IN " & strObjTypes
    
    Else
    
        strObjTypes = "= " & ObjType
    
    End If
    
    strSQL = "SELECT MsysObjects.Name AS [ObjectName]" & vbCrLf & _
        " FROM MsysObjects" & vbCrLf & _
        " WHERE (((MsysObjects.Name Not Like '~*') And (MsysObjects.Name Not Like 'MSys*'))" & vbCrLf & _
        "     AND (MsysObjects.Type" & strObjTypes & "))" & vbCrLf & _
        " ORDER BY MsysObjects.Name;"
    
'    strSQL = "SELECT MsysObjects.Name AS [ObjectName]" & vbCrLf & _
'           " FROM MsysObjects" & vbCrLf & _
'           " WHERE (((MsysObjects.Name Not Like '~*') And (MsysObjects.Name Not Like 'MSys*'))" & vbCrLf & _
'           "     AND (MsysObjects.Type=" & ObjType & "))" & vbCrLf & _
'           " ORDER BY MsysObjects.Name;"
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    With rs
        If .RecordCount <> 0 Then
            Do While Not .EOF
                If ExportToFile = True Then
                    strText = strText & ![objectname] & vbCrLf
                Else
                    Debug.Print ![objectname]
                End If
                .MoveNext
            Loop
        End If
    End With
 
    If ExportToFile = True Then
        
        'documentation file
        strFile = Application.CurrentProject.Path & "\" & CurrentProject.Name & "_ObjectList.txt"
        If Len(Dir(strFile)) > 0 Then Kill strFile
        FileNumber = FreeFile                           'Get unused file number
        Open strFile For Append As #FileNumber          'Create file name
        'Print #FileNumber, ""
        Print #FileNumber, String(80, "=")
        Print #FileNumber, "VBA Project Name: " & Application.VBE.ActiveVBProject.Name
        Print #FileNumber, "Description:      " & Application.VBE.ActiveVBProject.description
        Print #FileNumber, String(80, "-")
        Print #FileNumber, "Database:         " & Application.CurrentProject.Name
        Print #FileNumber, "Db Path:          " & Application.CurrentProject.Path
        Print #FileNumber, String(80, "=")
        Print #FileNumber, "As of:            " & Now()
        Print #FileNumber, String(80, "=")
        Print #FileNumber, "OBJECTS:"
        Print #FileNumber, String(80, "=")
        Print #FileNumber, strText
    
        Close #FileNumber  'close file.
        
    End If
 
Exit_Handler:
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    If Not db Is Nothing Then Set db = Nothing
    Exit Function
    
Err_Handler:
    'Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl)
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetObjectList[mod_Dev_Document])"
    End Select
    Resume Exit_Handler
End Function