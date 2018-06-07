Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_IDE
' Level:        Framework module
' Version:      1.00
' Description:  Project functions & procedures
'
' Source/date:  Bonnie Campbell, 3/26/2018
' Revisions:    BLC, 3/26/2018 - 1.00 - initial version
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
' SUB:          CloseAllModules
' Description:  Closes all open modules in the IDE interface
' Assumptions:  -
'
' Parameters:   all parameters are optional, if none are supplied modules will be
'               closed without saving any unsaved alterations within them
'
'               SaveModules - whether modules should be saved (optional long, default = acSaveNo)
'                             other values: acSaveYes, acSavePrompt
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Ken Sheridan, January 31, 2016
'   https://social.msdn.microsoft.com/Forums/sqlserver/en-US/5a001d45-187a-46b0-a6fc-857cb25fdbb9/how-to-close-all-openvba-code-windows?forum=accessdev
' Adapted:      Bonnie Campbell, March 26, 2018
' Revisions:
'   BLC - 3/26/2018  - initial version
' ---------------------------------
Public Sub CloseAllModules(Optional SaveModules As Long = acSaveNo)
On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim mods As Modules
    
    Set mods = Application.Modules
    
    'loop backward through all open modules
    'close each (save if SaveModules = True)
    For i = mods.Count - 1 To 0 Step -1
        
        DoCmd.Close acModule, mods(i).Name, SaveModules
    
    Next
    

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CloseAllModules[mod_Dev_Project])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     RunCloseAllModules
' Description:  Closes all open modules (can be run from immediate window??)
' Assumptions:  -
'
' Parameters:   all parameters are optional, if none are supplied modules will be
'               closed without saving any unsaved alterations within them
'
'               SaveModules - whether modules should be saved (optional long, default = acSaveNo)
'                             other values: acSaveYes, acSavePrompt
' Returns:      -
' Throws:       none
' References:   -
' Requires:     -
' Source/date:
'   Ken Sheridan, January 31, 2016
'   https://social.msdn.microsoft.com/Forums/sqlserver/en-US/5a001d45-187a-46b0-a6fc-857cb25fdbb9/how-to-close-all-openvba-code-windows?forum=accessdev
' Adapted:      Bonnie Campbell, March 26, 2018
' Revisions:
'   BLC - 3/26/2018  - initial version
' ---------------------------------
Public Function RunCloseAllModules(Optional SaveModules As Long = acSaveNo)
On Error GoTo Err_Handler
    
    CloseAllModules

Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RunCloseAllModules[mod_Dev_Project])"
    End Select
    Resume Exit_Handler
End Function