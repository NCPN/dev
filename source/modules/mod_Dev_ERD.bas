Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_ERD
' Level:        Development module
' Version:      1.02
'
' Description:  ERD related functions & procedures
'
' Source/date:  Bonnie Campbell, April 27, 2016
' Revisions:    BLC - 2/27/2016 - 1.00 - initial version
'               BLC - 6/15/2016 - 1.01 - added Stephen Leban's setup to retain ERD
' -------------------------------------------------------------------------------
'               BLC - 9/27/2017 - 1.02 - moved to NCPN_dev project
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32" Alias _
"FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, _
ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function apiGetWindow Lib "user32" _
Alias "GetWindow" (ByVal hwnd As Long, _
ByVal wCmd As Long) As Long

Private Declare Function apiGetClassName Lib "user32" _
Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long

Private Declare Function GetWindowRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function PostMessage Lib "user32" Alias _
"PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long

' ---------------------------------
'  Constants
' ---------------------------------
Private Const SWP_NOSIZE = &H1
Private Const WM_CLOSE = &H10
' GetWindow() Constants
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4
Private Const GW_CHILD = 5
Private Const GW_MAX = 5

' ---------------------------------
' Sub:          FixERD
' Description:  ERD fixing actions for positioning tables so they
'               are visible in the diagram (fixes negative positions)
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Stephen Lebans, August 28, 2006
'   https://bytes.com/topic/access/answers/528324-releationship-diagram-goes-haywire
' Source/date:  Bonnie Campbell, April 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub FixERD()
On Error GoTo Err_Handler


Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - FixERD[mod_Dev_ERD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cmdCloseWindow_Click
' Description:  Close the Debug Window via code
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Stephen Lebans, August 28, 2006
'   https://bytes.com/topic/access/answers/528324-releationship-diagram-goes-haywire
' Source/date:  Bonnie Campbell, April 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cmdCloseWindow_Click()
On Error GoTo Err_Handler
    Dim lngRet As Long
    Dim HwndMDI As Long
    Dim hWndDebug As Long
    
    ' If this instance of Access has set the
    ' Debug Window to "Always on Top" via the menu:
    ' Tools->Options->Module
    ' then the Debug WIndow is a top level window.
    hWndDebug = FindWindow(vbNullString, "Debug Window")
    If hWndDebug < 0 Then
        lngRet = PostMessage(hWndDebug, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If

    ' The Debug Window is a child of the MDI window
    ' find MDIClient first
    HwndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)

    ' Find the Debug Window
    hWndDebug = FindWindowEx(HwndMDI, 0&, "OImmediate", "Debug Window")
    If hWndDebug < 0 Then
        lngRet = PostMessage(hWndDebug, WM_CLOSE, 0&, 0&)
    Else
        MsgBox "The Debug Window is not open."
    End If

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdCloseWindow_Click[mod_Dev_ERD])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          cmdFix_Click
' Description:  Fix any windows that are off the Left edge of the Relationships window
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   Stephen Lebans, August 28, 2006
'   https://bytes.com/topic/access/answers/528324-releationship-diagram-goes-haywire
' Source/date:  Bonnie Campbell, April 27, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 4/27/2016 - initial version
' ---------------------------------
Private Sub cmdFix_Click()
On Error GoTo Err_Handler

    Dim lngRet As Long
    Dim HwndMDI As Long
    Dim hWndRel As Long
    Dim hWndODsk As Long
    Dim hWndTemp
    Dim rc As RECT
    
    ' Force the Relationships window to open
    DoCmd.RunCommand acCmdRelationships
    
    ' Window must be maximized
    DoCmd.Maximize
    DoEvents
    
    ' Relationships Window is a child of the MDI Client window
    ' find MDIClient first.
    HwndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)
    
    ' Find the Relationships Window
    hWndRel = FindWindowEx(HwndMDI, 0&, "OSysRel", "Relationships")
    
    If hWndRel = 0 Then
        MsgBox "The Relationships Window is not open.", vbCritical, "Critical Error"
        Exit Sub
    End If
    
    ' The first child window is of class ODsk
    hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)
    
    ' Loop through all of this level's Windows.
    ' We are looking for any windows with a negative
    ' Left value in it's Window Rectangle
    ' Let's get first Child Window of the ODsk window
    hWndTemp = apiGetWindow(hWndODsk, GW_CHILD)
    If hWndTemp = 0 Then
        MsgBox "Their are no Relationships!", vbCritical, "Critical Error"
        Exit Sub
    Else
        lngRet = GetWindowRect(hWndTemp, rc)
        If rc.Left < 1 Then
            rc.Left = 1
            lngRet = SetWindowPos(hWndTemp, 0&, rc.Left, rc.Top, 0&, 0&, SWP_NOSIZE)
        End If
    End If
    
    ' Let's walk through every sibling window
    Do
    
        ' Let's get the NEXT SIBLING Window
        hWndTemp = apiGetWindow(hWndTemp, GW_HWNDNEXT)
        
        If hWndTemp < 0 Then
            lngRet = GetWindowRect(hWndTemp, rc)
            If rc.Left < 1 Then
                rc.Left = 1
                lngRet = SetWindowPos(hWndTemp, 0&, rc.Left, rc.Top, 0&, _
                                        0&, SWP_NOSIZE)
            End If
        End If
        
        ' Let's Start the process from the Top again.
        ' End this loop if no more Windows.
    Loop While hWndTemp < 0
    ' All done

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - cmdFix_Click[mod_Dev_ERD])"
    End Select
    Resume Exit_Handler
End Sub

'*******
'Option Compare Database
'Option Explicit

''DEVELOPED AND TESTED UNDER MICROSOFT ACCESS 2000 VBA
''
''Copyright: Stephen Lebans - Lebans Holdings 1999 Inc.
''           Please feel free to use this code within your own projects,
''           both private and commercial, with no obligation.
''           You may not resell this code by itself or as part of a collection.
''
''
''Name:      Select Table in Relationship Window view
''
''Version:   1.1
''
''Purpose:   To allow the use to select and scroll into view a specific
''           table in the Relationship View window.
'' 
''Requires:  This code module
''
''Author:    Stephen Lebans
''
''Email:     Stephen@lebans.com
''
''Web Site:  www.lebans.com
''
''Date:      March 15, 2005, 11:11:11 PM
''
''Credits:   It's yours for the taking.
''
''BUGS:      Please report any bugs to:
''           Stephen@lebans.com
''
''What's Missing:
''           All Error handling
''           Add it yourself!
''
''How it Works:
''           Walk through the source code!<grin>
''
'' Enjoy
'' Stephen Lebans
''
'
'Private Type Rectl
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type
'
'
''Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
''
''Private Declare Function apiGetWindow Lib "user32" _
''Alias "GetWindow" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
''
''Private Declare Function GetWindowRect Lib "user32" _
''(ByVal hWnd As Long, lpRect As Rectl) As Long
''
''Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
''(ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
''
''Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
''(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
''
''' GetWindow() Constants
''Private Const GW_HWNDNEXT = 2
''Private Const GW_CHILD = 5
''
''' Collection for Window Text and hWnd
'''Private colWindows As New Collection
''
''Public Function fGetRelationshipViewWindows(aString As Variant) As Boolean
''
''On Error GoTo Err_Restore
''
''Dim lngRet As Long
''Dim HwndMDI As Long
''Dim hWndRel As Long
''Dim hWndODsk As Long
''Dim hWndTemp
''Dim s As String
''Dim X, Y, X1, Y1 As Integer
''Dim sWinName As String
''
''' I'm leaving the collection logic in for anyone that wants
''' to fill a Menu/ToolBar Combo control. The current implementation
''' uses a seperate form. would be much cleaner as a Menu/Toolbar ComboBox
'''' Collection for Window Text and hWnd
'''Dim colWindows As New Collection
''
''' Ensure we are zero
''hWndTemp = 0
''
''' The Relationships Window is a child of the MDI Client window
''' find MDIClient first.
''HwndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)
''' Find the Relationships Window
''hWndRel = FindWindowEx(HwndMDI, 0&, "OSysRel", "Relationships")
''' Do we have a valid Window handle
''If hWndRel = 0 Then
''    MsgBox "The Relationships Window is not open.", vbCritical, "Critical Error"
''   fGetRelationshipViewWindows = False
''    Exit Function
''End If
''
''' The first child window is of class ODsk
''hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)
''
''' Let's get first Child Window of the ODsk window
''hWndTemp = apiGetWindow(hWndODsk, GW_CHILD)
''
''If hWndTemp = 0 Then
''    MsgBox "Their are no Relationships!", vbCritical, "Critical Error"
''    fGetRelationshipViewWindows = False
''    Exit Function
''Else
''    ' Add this window to our collection
''    ' Grab the Windows Text
''    s = Space(256)
''    lngRet = GetWindowText(hWndTemp, s, 256)
''    'Debug.Print "S:" & s & Time
''    s = Left$(s, lngRet)
''    'colWindows.Add hWndTemp, s
''    ReDim aString(0)
''    aString(UBound(aString)) = s
''End If
''
''' Loop through the rest of the sibling windows adding them to our collection
''Do
''    hWndTemp = apiGetWindow(hWndTemp, GW_HWNDNEXT)
''    If hWndTemp = 0 Then Exit Do
''    ' Grab the Windows Text
''    s = Space(256)
''    lngRet = GetWindowText(hWndTemp, s, 256)
''    'Debug.Print "S:" & s & Time
''    s = Left$(s, lngRet)
''    'colWindows.Add hWndTemp, s
''    ReDim Preserve aString(UBound(aString) + 1)
''    aString(UBound(aString)) = s
''
''Loop
''
''' Return Success
''fGetRelationshipViewWindows = True
''
'' ' All done
''Exit_Restore:
''    ' Delete our Collection object
''    ' Remove the first object each time  through the loop until there are
''    ' no objects left in the collection.
'''    For X = 1 To colWindows.Count
'''        colWindows.Remove 1
'''    Next X
''Exit Function
''
''Err_Restore:
''    MsgBox Err.Description
''    Resume Exit_Restore
''
''End Function
''
''
''Public Function fScrollToRelationshipViewWindow(sName As String) As Boolean
''
''On Error GoTo Err_Restore
''
''Dim lngRet As Long
''Dim HwndMDI As Long
''Dim hWndRel As Long
''Dim hWndODsk As Long
''Dim hWndTemp
''Dim s As String
''Dim sWinName As String
''Dim lPos As Long
''
''Dim rc As Rectl
''Dim rcODsk As Rectl
''
''
''' Ensure we are zero
''hWndTemp = 0
''
''' The Relationships Window is a child of the MDI Client window
''' find MDIClient first.
''HwndMDI = FindWindowEx(Application.hWndAccessApp, 0&, "MDIClient", vbNullString)
''' Find the Relationships Window
''hWndRel = FindWindowEx(HwndMDI, 0&, "OSysRel", "Relationships")
''' Do we have a valid Window handle
''If hWndRel = 0 Then
''    MsgBox "The Relationships Window is not open.", vbCritical, "Critical Error"
''   fScrollToRelationshipViewWindow = False
''    Exit Function
''End If
''
''' The first child window is of class ODsk
''hWndODsk = FindWindowEx(hWndRel, 0&, "ODsk", vbNullString)
''
''' Let's get first Child Window of the ODsk window
''hWndTemp = apiGetWindow(hWndODsk, GW_CHILD)
''
''If hWndTemp = 0 Then
''    MsgBox "Their are no Relationships!", vbCritical, "Critical Error"
''    fScrollToRelationshipViewWindow = False
''    Exit Function
''Else
''    ' Add this window to our collection
''    ' Grab the Windows Text
''    s = Space(256)
''    lngRet = GetWindowText(hWndTemp, s, 256)
''    ' trim
''    s = Left$(s, lngRet)
''    If s = sName Then
''        ' We have a match!
''        ' Scroll Back to Home first
''        ' as Access physically moves the windows around
''        lngRet = fSetScrollBarPosH(hWndODsk, 0)
''        lngRet = fSetScrollBarPosV(hWndODsk, 0)
''
''        ' Grab Windows location
''        lngRet = GetWindowRect(hWndODsk, rcODsk)
''        lngRet = GetWindowRect(hWndTemp, rc)
''
''        ' Calculate for Horizontal ScrollBar
''        With rc
''           If .Left < (rcODsk.Right - (.Right - .Left)) Then
''               lPos = 0
''           Else
''               lPos = (.Left - rcODsk.Left) - (.Right - .Left)
''           End If
''        End With
''
''        ' Set the Horizontal SB position
''        lngRet = fSetScrollBarPosH(hWndODsk, lPos)
''
''        ' Calculate for Vertical ScrollBar
''        With rc
''           If .Top < (rcODsk.Bottom - (.Bottom - .Top)) Then
''               lPos = 0
''           Else
''               lPos = Abs((.Top - rcODsk.Top) - (.Bottom - .Top))
''           End If
''        End With
''
''        ' Set the Vertical SB position
''        lngRet = fSetScrollBarPosV(hWndODsk, lPos)
''    End If
''
''
''End If
''
''' Loop through the rest of the sibling windows adding them to our collection
''Do
''hWndTemp = apiGetWindow(hWndTemp, GW_HWNDNEXT)
''    If hWndTemp = 0 Then Exit Do
''    ' Grab the Windows Text
''    s = Space(256)
''    lngRet = GetWindowText(hWndTemp, s, 256)
''    'trim
''    s = Left$(s, lngRet)
''    If s = sName Then
''        ' We have a match!
''        ' Scroll Back to Home first
''        ' as Access physically moves the windows around
''        lngRet = fSetScrollBarPosH(hWndODsk, 0)
''         lngRet = fSetScrollBarPosV(hWndODsk, 0)
''
''        ' Grab Windows location
''        lngRet = GetWindowRect(hWndODsk, rcODsk)
''        lngRet = GetWindowRect(hWndTemp, rc)
''        ' Calculate for Horizontal ScrollBar
''        With rc
''           If .Left < (rcODsk.Right - (.Right - .Left)) Then
''               lPos = 0
''           Else
''               lPos = (.Left - rcODsk.Left) - (.Right - .Left)
''           End If
''        End With
''
''        ' Set the Horizontal SB position
''        lngRet = fSetScrollBarPosH(hWndODsk, lPos)
''
''        ' Calculate for Vertical ScrollBar
''        With rc
''           If .Top < (rcODsk.Bottom - (.Bottom - .Top)) Then
''               lPos = 0
''           Else
''               lPos = Abs((.Top - rcODsk.Top) - (.Bottom - .Top))
''           End If
''        End With
''
''        ' Set the Vertical SB position
''        lngRet = fSetScrollBarPosV(hWndODsk, lPos)
''
''    End If
''
''Loop
''
''' Return Success
''fScrollToRelationshipViewWindow = True
''
''
'' ' All done
''Exit_Restore:
''
''Exit Function
''
''Err_Restore:
''    MsgBox Err.Description
''    Resume Exit_Restore
''
''End Function
''
''
''
''' Fill Control(Combo/List) with Table/Query Names from Relationship View window
''
''Function fListFill(ctl As Control, varID As Variant, varRow As Variant, _
''                varCol As Variant, varCode As Variant) As Variant
''
''
''Dim varRet As Variant
''' Array of our Table/Query Names
''Static varWindowNames() As Variant
''
''    On Error GoTo ErrHandler
''    Select Case varCode
''
''        Case acLBInitialize
''        ' Fill our array of Table/Query names
''        Erase varWindowNames
''        fGetRelationshipViewWindows varWindowNames
''
''        QuickSort varWindowNames
''            varRet = True
''
''        Case acLBOpen
''            varRet = Timer
''
''        Case acLBGetRowCount
''            varRet = UBound(varWindowNames) + 1 ' ZERO based Array
''
''        Case acLBGetColumnWidth
''            'Set the widths of the column
''            varRet = -1
''
''        Case acLBGetColumnCount
''            varRet = 1
''
''        Case acLBGetValue
''            'Return the selected Table/Query name
''            varRet = varWindowNames(varRow)
''    End Select
''
''    fListFill = varRet
''ExitHere:
''    Exit Function
''ErrHandler:
''    Resume ExitHere
''End Function
''' ************* Code End **************
''
''
''Public Function PopWindow() As Boolean
''' Open our form to allow the user to select a Table/Query
''DoCmd.OpenForm "frmSelectWindowRelationshipView"
''
''End Function
''
''
''
'''Date: 9/4/1999
'''Versions: VB4 VB5 VB6 Level: Intermediate
'''Author: The VB2TheMax Team
''
''' QuickSort an array of any type
''' QuickSort is especially convenient with large arrays (>1,000
''' items) that contains items in random order. Its performance
''' quickly degrades if the array is already almost sorted. (There are
''' variations of the QuickSort algorithm that work good with
''' nearly-sorted arrays, though, but this routine doesn't use them.)
'''
''' NUMELS is the index of the last item to be sorted, and is
''' useful if the array is only partially filled.
'''
''' Works with any kind of array, except UDTs and fixed-length
''' strings, and including objects if your are sorting on their
''' default property. String are sorted in case-sensitive mode.
'''
''' You can write faster procedures if you modify the first two lines
''' to account for a specific data type, eg.
''' Sub QuickSortS(arr() As Single, Optional numEls As Variant,
'''  '     Optional descending As Boolean)
'''   Dim value As Single, temp As Single
''
''Sub QuickSort(arr As Variant, Optional numEls As Variant, _
''    Optional descending As Boolean)
''
''    Dim value As Variant, temp As Variant
''    Dim sp As Integer
''    Dim leftStk(32) As Long, rightStk(32) As Long
''    Dim leftNdx As Long, rightNdx As Long
''    Dim i As Long, j As Long
''
''
''' add our Header for this array where it will become the
''' first row for the ListBox
'''ReDim Preserve arr(UBound(arr) + 1)
'''arr(UBound(arr)) = " Font"
''
''    ' account for optional arguments
''    If IsMissing(numEls) Then numEls = UBound(arr)
''    ' init pointers
''    leftNdx = LBound(arr)
''    rightNdx = numEls
''    ' init stack
''    sp = 1
''    leftStk(sp) = leftNdx
''    rightStk(sp) = rightNdx
''
''    Do
''        If rightNdx > leftNdx Then
''            value = arr(rightNdx)
''            i = leftNdx - 1
''            j = rightNdx
''            ' find the pivot item
''            If descending Then
''                Do
''                    Do: i = i + 1: Loop Until arr(i) <= value
''                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) >= value
''                    temp = arr(i)
''                    arr(i) = arr(j)
''                    arr(j) = temp
''                Loop Until j <= i
''            Else
''                Do
''                    Do: i = i + 1: Loop Until arr(i) >= value
''                    Do: j = j - 1: Loop Until j = leftNdx Or arr(j) <= value
''                    temp = arr(i)
''                    arr(i) = arr(j)
''                    arr(j) = temp
''                Loop Until j <= i
''            End If
''            ' swap found items
''            temp = arr(j)
''            arr(j) = arr(i)
''            arr(i) = arr(rightNdx)
''            arr(rightNdx) = temp
''            ' push on the stack the pair of pointers that differ most
''            sp = sp + 1
''            If (i - leftNdx) > (rightNdx - i) Then
''                leftStk(sp) = leftNdx
''                rightStk(sp) = i - 1
''                leftNdx = i + 1
''            Else
''                leftStk(sp) = i + 1
''                rightStk(sp) = rightNdx
''                rightNdx = i - 1
''            End If
''        Else
''            ' pop a new pair of pointers off the stacks
''            leftNdx = leftStk(sp)
''            rightNdx = rightStk(sp)
''            sp = sp - 1
''            If sp = 0 Then Exit Do
''        End If
''    Loop
''End Sub
''
''' ****CODE START****
''' Place this code in a standard module.
''' make sure you do not name the module
''' to conflict with any of the functions below.
''
''
'''Author:    Stephen Lebans
'''           Stephen@lebans.com
'''           www.lebans.com
'''           March 15, 2005
'''
'''Copyright: Lebans Holdings 1999 Ltd.
'''
'''Functions: fSGetScrollBarPosH(hWnd as long) As Long
'''           fSGetScrollBarPosV(hWnd as long) As Long
'''
'''Credits:   Yours for the taking
'''
'''Why?:      Somebody asked for it!
'''
'''BUGS:      Let me know!
'''           :-)
''
''
'' Private Declare Function SendMessage Lib _
'' "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
'' ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
''
''' Windows Message Constant
''Private Const WM_HSCROLL = &H114
''Private Const WM_VSCROLL = &H115
''' Scroll Bar Commands
''
''' Scroll Bar Commands
''Private Const SB_THUMBPOSITION = 4
''
''
''
''Public Function fSetScrollBarPosH(ByVal hWnd As Long, ByVal lngIndex As Long) As Long
''' Set the Thumb Position for the
''' Vertical ScrollBar of the Form passed to
''' this Function.
''
''Dim hWndSB As Long
''Dim lngRet As Long
''Dim LngThumb As Long
''
''' Set the value  for the ScrollBar.
''' This corresponds to the top most record
''' that will be displayed in the Form.
''LngThumb = MakeDWord(SB_THUMBPOSITION, CInt(lngIndex))
''lngRet = SendMessage(hWnd, WM_HSCROLL, ByVal LngThumb, ByVal 0)
''
''' Return Success as our new ScrollBar Position
''fSetScrollBarPosH = lngIndex
''
''End Function
''
''
''Public Function fSetScrollBarPosV(ByVal hWnd As Long, ByVal lngIndex As Long) As Long
''' Set the Thumb Position for the
''' Vertical ScrollBar of the Form passed to
''' this Function.
''
''Dim hWndSB As Long
''Dim lngRet As Long
''Dim LngThumb As Long
''
''
''' Set the value  for the ScrollBar.
''' This corresponds to the top most record
''' that will be displayed in the Form.
''LngThumb = MakeDWord(SB_THUMBPOSITION, CInt(lngIndex))
''lngRet = SendMessage(hWnd, WM_VSCROLL, ByVal LngThumb, ByVal 0)
''
''' Return Success as our new ScrollBar Position
''fSetScrollBarPosV = lngIndex
''
''End Function
''
''
''' Here's the MakeDWord function from the MS KB
''Private Function MakeDWord(loword As Integer, hiword As Integer) As Long
''MakeDWord = (hiword * &H10000) Or (loword And &HFFFF&)
''End Function '***END CODE
''
''
'''************* relform *************
''Private Sub Form_Load()
''DoCmd.MoveSize 200, 1200, 4900, 4950
''' Setup Listbox Control Source
''' Force the RelationShips window to open
''DoCmd.RunCommand acCmdRelationships
''DoEvents
''Me.listStored.RowSourceType = "fListFill"
''End Sub
''
''Private Sub listStored_Click()
''fScrollToRelationshipViewWindow Me.listStored.value
''End Sub
''