Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Git
' Level:        Development module
' Version:      1.05
'
' Description:  Git related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, 2/12/2015
' Revisions:    BLC - 2/12/2015 - 1.00 - initial version
'               BLC - 5/27/2015 - 1.01 - renamed to note as development module
'               BLC - 6/30/2015 - 1.02 - added error handling to FieldTypeName()
'               BLC - 6/24/2016 - 1.03 - replaced Exit_Function > Exit_Handler
'               BLC - 1/24/2017 - 1.04 - shifted documentation functions to mod_Dev_Document
'                                        this includes:
'                                        GetDescriptions(),TableInfo(),
'                                        GetDescrip(), DocumentDb()
' -------------------------------------------------------------------------------
'               BLC - 9/27/2017 - 1.05 - moved to NCPN_dev project
' ===================================================================================
'  NOTE:
'  To regenerate components backed up w/ functions using SaveAsText
'  Use the following:
'       Application.LoadFromText acForm, "YourFormName", "C:\Temp\Form_frmTest.txt"
' ===================================================================================

' ---------------------------------
' FUNCTION:     ExportVBComponent
' Description:  Export VB components (forms, modules, classes) as text files for later use
' Assumptions:  -
' Parameters:   VBComp - VB component
'               FolderName - name of folder component is in (string)
'               fileName - name of VB component file (string)
'               OverwriteExisting - true if existing file should be overwritten, false if not (boolean)
' Returns:      boolean - true if successful export, false if not
' Throws:       none
' References:   Requires Microsoft Visual Basic for Applications Extensibility 5.3 library (add via Tools > References)
' Source/date:
' Chip Pearson
' http://www.cpearson.com/excel/vbe.aspx
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, _
                FolderName As String, _
                Optional FileName As String, _
                Optional OverwriteExisting As Boolean = True) As Boolean
On Error GoTo Err_Handler

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Exports module code of a VBComponent to a text file.
    ' If FileName is missing, the code will be exported to
    ' a file with the same name as the VBComponent followed by the
    ' appropriate extension.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Extension As String
    Dim fName As String
    Extension = GetFileExtension(VBComp:=VBComp)
    If Trim(FileName) = vbNullString Then
        fName = VBComp.Name & Extension
    Else
        fName = FileName
        If InStr(1, fName, ".", vbBinaryCompare) = 0 Then
            fName = fName & Extension
        End If
    End If
    
    If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
        fName = FolderName & fName
    Else
        fName = FolderName & "\" & fName
    End If
    
    If dir(fName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
        If OverwriteExisting = True Then
            Kill fName
        Else
            ExportVBComponent = False
            Exit Function
        End If
    End If
    
    VBComp.Export FileName:=fName
    ExportVBComponent = True

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExportVBComponent[mod_Dev_Git])"
    End Select
    Resume Exit_Handler
End Function
    
' ---------------------------------
' FUNCTION:     GetFileExtension
' Description:  Return appropriate file extension for VB Components(forms, modules, classes) as text files for later use
' Assumptions:  -
' Parameters:   VBComp - VB component
' Returns:      string - file extension of VB component
' Throws:       none
' References:   none
' Source/date:
' Chip Pearson
' http://www.cpearson.com/excel/vbe.aspx
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
On Error GoTo Err_Handler
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This returns the appropriate file extension based on the Type of
    ' the VBComponent.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                GetFileExtension = ".cls"
            Case vbext_ct_Document
                GetFileExtension = ".cls"
            Case vbext_ct_MSForm
                GetFileExtension = ".frm"
            Case vbext_ct_StdModule
                GetFileExtension = ".bas"
            Case Else
                GetFileExtension = ".bas"
        End Select

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFileExtension[mod_Dev_Git])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          DocDatabase
' Description:  Documents the database to a series of text files
' Assumptions:  -
' Parameters:   path - path to database file (string, optional)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Arvin Meyer, June 2, 1999
' http://www.datastrat.com/Code/DocDatabase.txt
' Renaud Bompius, June 20, 2011
' http://stackoverflow.com/questions/6408951/text-search-in-properties-access-objects/6410015#6410015
' Usage:
' Call DocDatabase from the Access IDE Immediate window.
' Creates a set of directories under and 'Exploded View' folder that will contain all the files.
' Comment: Uses the undocumented [Application.SaveAsText] syntax
'          To reload use the syntax [Application.LoadFromText]
'          Modified to set a reference to DAO 8/22/2005
'          Modified by Renaud Bompuis to export Queries as proper SQL
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Sub DocDatabase(Optional Path As String = "")
    
    If IsBlank(Path) Then
        Path = Application.CurrentProject.Path & "\" & Application.CurrentProject.Name & " - exploded view\"
    End If

On Error Resume Next
    MkDir Path
    MkDir Path & "\Forms\"
    MkDir Path & "\Queries\"
    MkDir Path & "\Queries(SQL)\"
    MkDir Path & "\Reports\"
    MkDir Path & "\Modules\"
    MkDir Path & "\Scripts\"

On Error GoTo Err_Handler
    Dim dbs As DAO.Database
    Dim cnt As DAO.Container
    Dim doc As DAO.Document
    Dim i As Integer

    Set dbs = CurrentDb() ' use CurrentDb() to refresh Collections

    Set cnt = dbs.Containers("Forms")
    For Each doc In cnt.Documents
        Application.SaveAsText acForm, doc.Name, Path & "\Forms\" & doc.Name & ".txt"
    Next doc

    Set cnt = dbs.Containers("Reports")
    For Each doc In cnt.Documents
        Application.SaveAsText acReport, doc.Name, Path & "\Reports\" & doc.Name & ".txt"
    Next doc

    Set cnt = dbs.Containers("Scripts")
    For Each doc In cnt.Documents
        Application.SaveAsText acMacro, doc.Name, Path & "\Scripts\" & doc.Name & ".txt"
    Next doc

    Set cnt = dbs.Containers("Modules")
    For Each doc In cnt.Documents
        Application.SaveAsText acModule, doc.Name, Path & "\Modules\" & doc.Name & ".txt"
    Next doc

    Dim intfile As Long
    Dim FileName As String
    For i = 0 To dbs.QueryDefs.Count - 1
         Application.SaveAsText acQuery, dbs.QueryDefs(i).Name, Path & "\Queries\" & dbs.QueryDefs(i).Name & ".txt"
         FileName = Path & "\Queries(SQL)\" & dbs.QueryDefs(i).Name & ".txt"
         intfile = FreeFile()
         Open FileName For Output As #intfile
         Print #intfile, dbs.QueryDefs(i).SQL
         Close #intfile
    Next i

    Set doc = Nothing
    Set cnt = Nothing
    Set dbs = Nothing

Exit_Handler:
    Debug.Print "Done."
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DocDatabase[mod_Dev_Git])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RecreateDatabase
' Description:  Recreates the database from series of text files created through SaveAsText
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' Curtis Inderwiesche, May 13, 2009
' http://stackoverflow.com/questions/859530/alternative-to-application-loadfromtext-for-ms-access-queries
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Sub RecreateDatabase()
On Error GoTo Err_Handler
    Dim myFile As Object '??
    Dim Folder As Object '??
    Dim FSO As Object '??
    Dim objecttype As String, objectname As String
    Dim WScript As Object '??
    Dim oApplication As Object '??
    
    For Each myFile In Folder.Files
        objecttype = FSO.GetExtensionName(myFile.Name)
        objectname = FSO.GetBaseName(myFile.Name)
        WScript.Echo "  " & objectname & " (" & objecttype & ")"
    
        If (objecttype = "form") Then
            oApplication.LoadFromText acForm, objectname, myFile.Path
        ElseIf (objecttype = "bas") Then
            oApplication.LoadFromText acModule, objectname, myFile.Path
        ElseIf (objecttype = "mac") Then
            oApplication.LoadFromText acMacro, objectname, myFile.Path
        ElseIf (objecttype = "report") Then
            oApplication.LoadFromText acReport, objectname, myFile.Path
        ElseIf (objecttype = "sql") Then
            oApplication.LoadFromText acQuery, objectname, myFile.Path
        End If
        
    Next
    
Exit_Handler:
    Debug.Print "Done."
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RecreateDatabase[mod_Dev_Git])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     SetPropertyDAO
' Description:  Set the property of a database object
' Assumptions:  -
' Parameters:   obj - object to set property for (object)
'               strPropertyName - property name (string)
'               intType - type of property (integer)
'               varValue - value of property (variant)
'               strErrMsg - error message to display (string)
' Returns:      boolean - true if property value is set, false if not
' Throws:       none
' References:   Requires Microsoft Visual Basic for Applications Extensibility 5.3 library (add via Tools > References)
' Source/date:
' HansUp, August 19, 2010
' http://stackoverflow.com/questions/3521188/ms-access-setting-table-column-caption-or-description-in-ddl
' Allen Browne
' http://allenbrowne.com/AppPrintMgtCode.html#SetPropertyDAO
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, _
    varValue As Variant, Optional strErrMsg As String) As Boolean
On Error GoTo errHandler
    'Purpose:   Set a property for an object, creating if necessary.
    'Arguments: obj = the object whose property should be set.
    '           strPropertyName = the name of the property to set.
    '           intType = the type of property (needed for creating)
    '           varValue = the value to set this property to.
    '           strErrMsg = string to append any error message to.

    If HasProperty(obj, strPropertyName) Then
        obj.Properties(strPropertyName) = varValue
    Else
        obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
    End If
    SetPropertyDAO = True

ExitHandler:
    Exit Function

errHandler:
    strErrMsg = strErrMsg & obj.Name & "." & strPropertyName & " not set to " & _
        varValue & ". Error encountered (#" & Err.Number & " - SetPropertyDAO[mod_Dev_Git])" & _
        Err.Number & " - " & Err.Description & vbCrLf
    
    Resume ExitHandler
End Function

' ---------------------------------
' FUNCTION:     HasProperty
' Description:  Returns true if object has the property
' Assumptions:  -
' Parameters:   obj - object to inspect (object)
'               strPropName - property name to find (string)
' Returns:      boolean - true if object has property, false if not
' Throws:       none
' References:   -
' Source/date:
' HansUp, August 19, 2010
' http://stackoverflow.com/questions/3521188/ms-access-setting-table-column-caption-or-description-in-ddl
' Allen Browne
' http://allenbrowne.com/AppPrintMgtCode.html#SetPropertyDAO
' Adapted:      Bonnie Campbell, February 12, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/12/2015 - initial version
' ---------------------------------
Public Function HasProperty(obj As Object, strPropName As String) As Boolean
    'Purpose: Return true if the object has the property.
    Dim varDummy As Variant

    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function

' ---------------------------------
' FUNCTION:     FieldTypeName
' Description:  Converts the numeric results of DAO Field.Type to text
' Assumptions:  -
' Parameters:   fld - DAO field
' Returns:      type name - name of field's type (string) - Yes/No, Byte, Integer, etc.
' Throws:       none
' References:   -
' Source/date:
' Allen Browne, April 2010
' http://allenbrowne.com/func-06.html
' Adapted:      Bonnie Campbell, February 13, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/13/2015 - initial version
'   BLC - 6/30/2015 - added error handling
' ---------------------------------
Public Function FieldTypeName(fld As DAO.field) As String
On Err GoTo Err_Handler

    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FieldTypeName[mod_Dev_Git])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     GetFieldTypeName
' Description:  Converts the numeric results of DAO Field.Type to text
' Assumptions:  -
' Parameters:   fld - DAO field
' Returns:      type name - name of field's type (string) - Yes/No, Byte, Integer, etc.
' Throws:       none
' References:   -
' Source/date:
' Allen Browne, April 2010
' http://allenbrowne.com/func-06.html
' Adapted:      Bonnie Campbell, May 26, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/26/2015 - initial version
' ---------------------------------
Function GetFieldTypeName(fld As Integer) As String
On Err GoTo Err_Handler

    Dim strReturn As String    'Name to return

    Select Case CLng(fld) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean, 1: strReturn = "Yes/No"           ' 1
        Case dbByte, 2: strReturn = "Byte"                ' 2
        Case dbInteger, 3: strReturn = "Integer"          ' 3
        Case dbLong, 4                                    ' 4
            'If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            'Else
            '    strReturn = "AutoNumber"
            'End If
        Case dbCurrency, 5: strReturn = "Currency"        ' 5
        Case dbSingle, 6: strReturn = "Single"            ' 6
        Case dbDouble, 7: strReturn = "Double"            ' 7
        Case dbDate, 8: strReturn = "Date/Time"           ' 8
        Case dbBinary, 9: strReturn = "Binary"            ' 9 (no interface)
        Case dbText, 10                                    '10
            'If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            'Else
            '    strReturn = "Text (fixed width)"        '(no interface)
            'End If
        Case dbLongBinary, 11: strReturn = "OLE Object"    '11
        Case dbMemo, 12                                    '12
            'If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            'Else
            '    strReturn = "Hyperlink"
            'End If
        Case dbGUID, 15: strReturn = "GUID"                '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt, 16: strReturn = "Big Integer"       '16
        Case dbVarBinary, 17: strReturn = "VarBinary"      '17
        Case dbChar, 18: strReturn = "Char"                '18
        Case dbNumeric, 19: strReturn = "Numeric"          '19
        Case dbDecimal, 20: strReturn = "Decimal"          '20
        Case dbFloat, 21: strReturn = "Float"              '21
        Case dbTime, 22: strReturn = "Time"                '22
        Case dbTimeStamp, 23: strReturn = "Time Stamp"     '23

        'Constants for complex types don't work prior to Access 2007 and later.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld & " unknown"
    End Select

    GetFieldTypeName = strReturn


Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetFieldTypeName[mod_Dev_Git])"
    End Select
    Resume Exit_Handler
End Function