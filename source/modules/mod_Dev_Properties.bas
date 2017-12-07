Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Properties
' Level:        Development module
' Version:      1.00
'
' Description:  Property related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, November 22, 2017
' Revisions:    BLC - 11/22/2017 - 1.00 - initial version
' =================================

' ===================================================================================
'  NOTE:
'       Functions and subroutines within this module are for debugging and test
'       purposes.
'
'       When the application is ready for release, this module can be
'       removed without negative impact to the application.
'
'       All mod_Debug_XX (debugging) and VCS_XX (version control system) modules can also be removed.
' ===================================================================================

' ---------------------------------
' FUNCTION:     AddDbProperty
' Description:  add custom properties to a database application
' Assumptions:  -
' Examples (from Immediate window):
'    ?AddDbProperty("Copyright Notice", "© 2017 B.Campbell for NCPN")
'    ?CurrentDb.Properties![Copyright Notice]
'       © 2017 B.Campbell for NCPN
'    ?AddDbProperty("Designed & Developed By", "B.Campbell")
'    ?CurrentDb.Properties![Designed & Developed By]
'       B.Campbell
' Parameters:   DbProperty - property to add (string)
'               DbPropertyValue - value for (string)
'               DbPropertyType - type values should be (optional, default DB_TEXT)
'               DbFilename - database to add property to (e.g. "C:\mydb.accdb", optional, default = Current aka CurrentDb)
' Returns:      -
' Throws:       none
' References:
'   Paul Murray, 6/14/1995
'   http://allenbrowne.com/ser-09.html
' Source/date:  Bonnie Campbell, November 22, 2017
' Adapted:      -
' Revisions:
'   BLC - 11/22/2017 - initial version
' ---------------------------------
Public Function AddDbProperty(DbProperty As String, _
                    DbPropertyValue As String, _
                    Optional DbPropertyType As Long = DB_TEXT, _
                    Optional DbFilename As String = "Current")
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim prop As Property
    
    If DbFilename = "Current" Then
        Set db = DBEngine(0)(0)
    Else
        Set db = OpenDatabase(DbFilename)
    End If

    'add the property
    Set prop = db.CreateProperty(DbProperty, DbPropertyType, DbPropertyValue)
    db.Properties.Append prop
    
Exit_Handler:
    db.Close
    Set db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddDbProperty[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     RemoveDbProperty
' Description:  remove custom properties from a database applciation
' Assumptions:  -
' Examples (from Immediate window):
'    ?RemoveDbProperty("Copyright Notice")
' Parameters:   DbProperty - property to add (string)
'               DbFilename - database to add property to (e.g. "C:\mydb.accdb", optional, default = Current aka CurrentDb)
' Returns:      -
' Throws:       none
' References:
'   Paul Murray, 6/14/1995
'   http://allenbrowne.com/ser-09.html
' Source/date:  Bonnie Campbell, November 22, 2017
' Adapted:      -
' Revisions:
'   BLC - 11/22/2017 - initial version
' ---------------------------------
Public Function RemoveDbProperty(DbProperty As String, _
                    Optional DbFilename As String = "Current")
On Error GoTo Err_Handler

    Dim db As DAO.Database
    
    If DbFilename = "Current" Then
        Set db = DBEngine(0)(0)
    Else
        Set db = OpenDatabase(DbFilename)
    End If

    'remove the property
    db.Properties.Delete DbProperty
    
Exit_Handler:
    db.Close
    Set db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RemoveDbProperty[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     UpdateDbProperty
' Description:  add custom properties to a database application
' Assumptions:  -
' Examples (from Immediate window):
'    ?UpdateDbProperty("Developer", "B.Campbell for NCPN")
'    ?CurrentDb.Properties![Developer]
'       B.Campbell for NCPN
' Parameters:   DbProperty - property to add (string)
'               DbPropertyValue - value for (string)
'               DbFilename - database to add property to (e.g. "C:\mydb.accdb", optional, default = Current aka CurrentDb)
' Returns:      -
' Throws:       none
' References:
'   Paul Murray, 6/14/1995
'   http://allenbrowne.com/ser-09.html
' Source/date:  Bonnie Campbell, November 22, 2017
' Adapted:      -
' Revisions:
'   BLC - 11/22/2017 - initial version
' ---------------------------------
Public Function UpdateDbProperty(DbProperty As String, _
                    DbPropertyValue As String, _
                    Optional DbFilename As String = "Current")
On Error GoTo Err_Handler

    Dim db As DAO.Database
    
    If DbFilename = "Current" Then
        Set db = DBEngine(0)(0)
    Else
        Set db = OpenDatabase(DbFilename)
    End If

    'add the property
    db.Properties(DbProperty) = DbPropertyValue
    
Exit_Handler:
    db.Close
    Set db = Nothing
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - UpdateDbProperty[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Function