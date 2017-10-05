Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Project
' Level:        Framework module
' Version:      1.02
' Description:  Project functions & procedures
'
' Source/date:  Bonnie Campbell, 9/19/2017
' Revisions:    BLC, 9/19/2017 - 1.00 - initial version
'               BLC, 9/21/2017 - 1.01 - add ListObjects()
' -------------------------------------------------------------------------------
'               BLC, 9/27/2017 - 1.02 - moved to NCPN_dev
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
' SUB:          ExportObjects
' Description:  Export objects from the application/project
' Assumptions:  -
'
' Parameters:   all parameters are optional, if none are supplied the objects will
'               be exported to a "components" folder within the current database's directory
'               and files will not be overwritten and objects will not be removed from the project
'
'               Overwrite - whether files should be overwritten (optional boolean, default = false)
'               Remove - whether files should be removed (optional boolean, default = false)
'               DestDir - main directory (optional string)
'               DestFolder - new folder to place files into (optional string)
' Returns:      -
' Throws:       none
' References:   Extensibility library
' Requires:     -
' Source/date:
'   Andy Pop, May 25, 2007
'   http://www.ozgrid.com/forum/showthread.php?t=69333
' Adapted:      Bonnie Campbell, September 19, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/19/2017  - initial version
' ---------------------------------
Public Sub ExportObjects(Optional Overwrite As Boolean = False, _
                            Optional Remove As Boolean = False, _
                            Optional DestDir As String = "", _
                            Optional DestFolder As String = "")
On Error GoTo Err_Handler
    
    Dim oProj As VBProject
    Dim oVBComp As VBComponent
    Dim ext As String, fName As String
    
    'set default directory
    If Len(DestDir) = 0 Then
          
          If Len(DestFolder) = 0 Then DestFolder = "components\"
          
          DestDir = Application.CurrentProject.Path & "\" & DestFolder
    
    End If
    
    'make directory if needed
    If Dir(DestDir, vbDirectory) = vbNullString Then MkDir DestDir
    
    Set oProj = Application.VBE.ActiveVBProject
    
    'iterate
    For Each oVBComp In oProj.VBComponents
    
        Select Case oVBComp.Type
            Case vbext_ct_StdModule 'module
                ext = "bas"
            Case vbext_ct_ClassModule 'class
                ext = "cls"
            Case vbext_ct_MSForm 'form
                ext = "frm"
            Case vbext_ct_Document 'document
                ext = "cls"
            Case Else
                ext = vbNullString
        End Select
        
        If ext <> vbNullString Then
        
            'export all components
            
            fName = DestDir & oVBComp.Name & "." & ext
                        
            'overwrite existing file?
            If Overwrite = True And _
                Dir(fName, vbNormal) <> vbNullString Then
                Kill (fName)
            End If
            
            oVBComp.Export (fName)
            
        End If
    
        'remove?
        If Remove = True Then oProj.VBComponents.Remove oVBComp
    
    Next

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ExportObject[mod_Dev_Project])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          RemoveObject
' Description:  Remove object(s) from the application/project
' Assumptions:
'    oType options:  vbext_ct_StdModule 'module     vbext_ct_ClassModule 'class
'                    vbext_ct_MSForm    'form       vbext_ct_Document    'document (user forms & reports)
'
' Parameters:   removes one object (oName & oType provided), a set of like objects (oType provided),
'               or all objects (AllObjects = true)
'               AllObjects - remove all objects? (optional boolean, default = false)
'               oName - name of object to remove (optional string, default = "")
'               oType - type of object to remove (optional variant)
' Returns:      -
' Throws:       none
' References:   Extensibility library
'   Alain Bryden, January 20, 2011
'   https://www.experts-exchange.com/articles/1457/Automate-Exporting-all-Components-in-an-Excel-Project.html
' Requires:     -
' Source/date:  Bonnie Campbell, September 20, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/20/2017  - initial version
' ---------------------------------
Public Sub RemoveObject(Optional AllObjects As Boolean = False, _
                        Optional oName As String = "", _
                        Optional oType As Variant)
On Error GoTo Err_Handler
    
    Dim oProj As VBProject
    Dim oVBComp As VBComponent
    Dim blnRemove As Boolean
        
    Set oProj = Application.VBE.ActiveVBProject
    
    'iterate
    For Each oVBComp In oProj.VBComponents
    
        'default
        blnRemove = False
    
        'remove all objects?
        If AllObjects = True Then blnRemove = True
        
        'remove all objects of single type?
        If Len(oName) = 0 And oVBComp.Type = oType Then blnRemove = True
        
        'remove single object?
        If oVBComp.Name = oName And oVBComp.Type = oType Then blnRemove = True
        
        'remove?
        If blnRemove = True Then oProj.VBComponents.Remove oVBComp
            
    Next

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveObject[mod_Dev_Project])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ImportObjects
' Description:  Import objects into the application/project
' Assumptions:  -
' Parameters:   adds one object (oName provided), a set of like objects (oType provided),
'               or all objects (AllObjects = true)
'               AllObjects - add all objects? (optional boolean, default = false)
'               SrcDir - source directory (optional string, default = "" -> current project's directory)
'               oName - name of object to remove (optional string, default = "")
'               oType - type of object to remove (optional variant)
' Returns:      -
' Throws:       none
' References:   Extensibility library
' Requires:     -
' Source/date:  Bonnie Campbell, September 20, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/20/2017  - initial version
' ---------------------------------
Public Sub ImportObject(Optional AllObjects As Boolean = False, _
                        Optional SrcDir As String = "", _
                        Optional oName As String = "", _
                        Optional oType As Variant = vbext_ct_StdModule)
On Error GoTo Err_Handler
    
    Dim oProj As VBProject
    Dim oVBComp As VBComponent
    Dim SrcFile As File
    Dim ext As String
    Dim blnAdd As Boolean
        
    Set oProj = Application.VBE.ActiveVBProject
    
    Select Case oType
        Case vbext_ct_StdModule 'module
            ext = ".bas"
        Case vbext_ct_ClassModule 'class
            ext = ".cls"
        Case vbext_ct_MSForm 'form
            ext = ".frm"
        Case vbext_ct_Document 'document
            ext = ".cls"
        Case Else
            ext = vbNullString
    End Select
    
    'source directory?
    If Len(SrcDir) = 0 Then SrcDir = Application.CurrentProject.Path
    
    Dim fs As Scripting.FileSystemObject
    Dim f As Folder
    Dim fc 'As Collection
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(SrcDir)
    Set fc = f.Files
    
    'iterate
    For Each SrcFile In fc
    
        'default
        blnAdd = False
    
        'add all objects?
        If AllObjects = True Then blnAdd = True
        
        'add all objects of single type?
        If Len(oName) = 0 And InStr(SrcFile.Name, ext) Then blnAdd = True
        
        'add single object?
        If SrcFile.Name = oName And InStr(SrcFile, ext) And ext <> vbNullString Then blnAdd = True
        
        'add?
        If blnAdd = True Then oProj.VBComponents.Import SrcFile
            
    Next

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ImportObject[mod_Dev_Project])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          ListObjects
' Description:  Lists objects in the application/project
' Assumptions:  -
' Parameters:   a set of like objects (oType provided),
'               or all objects (AllObjects = true)
'               AllObjects - list all objects? (optional boolean, default = false)
'               oType - type of objects to list (optional variant)
' Returns:      -
' Throws:       none
' References:   Extensibility library
' Requires:     -
' Source/date:  Bonnie Campbell, September 21, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 9/21/2017  - initial version
' ---------------------------------
Public Sub ListObjects(Optional AllObjects As Boolean = False, _
                        Optional oType As Variant = vbext_ct_StdModule)
On Error GoTo Err_Handler
    
    Dim oProj As VBProject
    Dim oVBComp As VBComponent
    Dim blnList As Boolean
        
    Set oProj = Application.VBE.ActiveVBProject
        
    'iterate
    For Each oVBComp In oProj.VBComponents
    
        'default
        blnList = False
    
        'list all objects?
        If AllObjects = True Then blnList = True
        
        'list all objects of single type?
        If oVBComp.Type = oType Then blnList = True
                
        'list?
        If blnList = True Then Debug.Print oVBComp.Name
            
    Next

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ListObjects[mod_Dev_Project])"
    End Select
    Resume Exit_Handler
End Sub

'-------------------------------
' Test Functions
'-------------------------------
Public Function expobj()
    ExportObjects

End Function

Public Function remobj()
    Dim d As String
    d = Application.CurrentProject.Path & "\imp\"
    'RemoveObject False, "VegPlot", vbext_ct_ClassModule
    RemoveObject False, "VegTransect1", vbext_ct_ClassModule
End Function

Public Function impfil()
    Dim d As String
    d = Application.CurrentProject.Path & "\imp\"
    ImportObject True, d
End Function

Public Function listobj()
    'ListObjects False, vbext_ct_ClassModule
    'ListObjects
    ListObjects True, vbext_ct_ClassModule
    'ListObjects True
    'ListObjects False, vbext_ct_MSForm
    'ListObjects False, vbext_ct_Document

End Function