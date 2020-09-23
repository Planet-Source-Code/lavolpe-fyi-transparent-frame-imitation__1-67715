VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTransParentFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Just a fun project I put together.
' This is as close to a real transparent frame as I think one can get. Of course I would
' expect others to improve upon this and post something a bit better.

' The class accomplishes faking transparency by having the parent container paint
' its background image to a memory DC that we then paint to our picturebox.
' Subclassing is used to catch most of the needed "AutoRedraw" events we would need
' but can't catch them all, so you will need to call Refresh in some specific cases.

' When can this AutoRedraw?
'   - When the picturebox is moved or sized
'   - When the system settings change (i.e., system colors and/or screen resolution)
'   - When the picturebox is partially hidden and then made completely visible

' When must you call the Refresh method?
'   - When any graphics under the picture box change. This includes any labels,
'       image controls or manual painting (i.e., gradients)
'   - When you change the background color of the picturebox's container
'   - Whenever you change one or more of the properties in this class

' How to use this class...
' 1. Simply call Attach and provide the picturebox that will act as a transparent frame
' 2. Set whatever properties you want set
' 3. Call Refresh
' 4. When you no longer need the picturebox to act as a transparent frame, call Detach

' Last but not least. If you are going to put a picturebox (as a transparent frame)
' inside a usercontrol or another picturebox, then keep this in mind too...
'   When the container (picturebox or usercontrol) is custom painted (i.e., gradients),
'   not a solid color, you should paint your object like below, before calling the Refresh
'   method of this class; otherwise, the graphics may be sent shifted 2 pixels vertically.

'    Object.AutoRedraw = True
'    < do your drawing routine >
'    Object.Picture = Object.Image
'    Object.AutoRedraw = False ' < optional


' That's it

'-------------------------------------------------------------------------------------------------
' Much appreciation goes towards Paul Caton for his self-subclassing thunks; makes some things much easier
'-Thunking/Callback declarations---------------------------------------------------------------------------
Private z_CbMem   As Long    'Callback allocated memory address
Private z_Cb()    As Long    'Callback thunk array

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC As Long = -4

' subclass message we are interested in
Private Const WM_PAINT As Long = &HF&
Private Const WM_DESTROY As Long = &H2
Private Const WM_DISPLAYCHANGE As Long = &H7E
Private Const WM_MOVE As Long = &H3
Private Const WM_SIZE As Long = &H5
Private Const WM_WININICHANGE As Long = &H1A
Private Const WM_SYSCOLORCHANGE As Long = &H15
'-------------------------------------------------------------------------------------------------


Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As Any) As Long
Private Declare Function GetUpdateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum etfAlignmentOptions
    tbpAlignTopLeft = 0
    tbpAlignTopCenter = 1
    tbpAlignTopRight = 2
    tbpAlignBottomLeft = 3
    tbpAlignBottomCenter = 4
    tbpAlignBottomRight = 5
End Enum
Public Enum etfBorderStyle
    tbpNone = 0
    tbpFlat = 1
    tbp3D = 2
End Enum

Private Type ClientSpecs
    hPointer As Long
    TextColorLite As Long       ' primary caption color (flat style)
    BorderColorLite As Long     ' primary border color (flat style)
    TextColorShadow As Long     ' shadow (darker) caption color
    BorderColorShadow As Long   ' shadow (darker) border color
    RenderedX As Long           ' how much of the picbox was painted? used to AutoRefresh ourselves
    RenderedY As Long           ' how much of the picbox was painted? used to AutoRefresh ourselves
    Alignment As etfAlignmentOptions ' caption alignment to frame
    BorderStyle As etfBorderStyle ' border 3D, flat or none , default is 3D
    isRaised As Boolean         ' caption 3D or not, default is not
    Caption As String           ' the frame caption
End Type

Private m_WndProc As Long       ' previous window procedure of subclassing picturebox
Private Const TextEdgeOffset = 3    ' gap between horizontal frame line stops & where text starts, adjust to taste
Private Const FrameEdgeOffset = 8  ' minimal horizontal edge before text is displayed, adjust to taste
Private objPicBox As ClientSpecs

Public Sub Attach(objPictureBox As PictureBox)
    
    ClearSubClassing    ' remove any previous subclassing
    With objPicBox      ' reset everything
        .hPointer = ObjPtr(objPictureBox)   ' pointer to picturebox
        .Alignment = tbpAlignTopLeft        ' default caption alignment
        .BorderColorLite = -1               ' use picturebox's forecolor
        .BorderColorShadow = -1             '   and or use system color
        .BorderStyle = tbp3D                ' 3D border is default
        .isRaised = False                   ' Flat caption is default
        .TextColorLite = -1                 ' use picturebox's forecolor
        .TextColorShadow = -1               '   and oruse system color
        .RenderedX = 0                      ' amount of picturebox that is rendered
        .RenderedY = 0                      ' amount of picturebox that is rendered
    End With
    
End Sub

Public Sub Detach()
    ' clear any subclassing.
    ClearSubClassing
End Sub

' Properties that can be set/read
' NOTE: Attach first, then set properties

Public Property Let Caption(Value As String)
    objPicBox.Caption = Value
End Property
Public Property Get Caption() As String
    Caption = objPicBox.Caption
End Property

Public Property Let BorderStyle(Style As etfBorderStyle)
    objPicBox.BorderStyle = Style
End Property
Public Property Get BorderStyle() As etfBorderStyle
    BorderStyle = objPicBox.BorderStyle
End Property

Public Sub SetBorderColors(ByVal Shadow As Long, Optional ByVal HighLight As Long)
    ' flat borders only use the Shadow color
    objPicBox.BorderColorShadow = Shadow
    If Me.BorderStyle = tbp3D Then objPicBox.BorderColorLite = HighLight
End Sub
Public Sub GetBorderColors(ByVal Shadow As Long, Optional ByVal HighLight As Long)
    Shadow = objPicBox.BorderColorShadow
    HighLight = objPicBox.BorderColorLite
End Sub

Public Property Let FlatCaption(isFlat As Boolean)
    objPicBox.isRaised = Not isFlat
End Property
Public Property Get FlatCaption() As Boolean
    FlatCaption = Not objPicBox.isRaised
End Property

Public Sub SetTextColors(ByVal Shadow As Long, Optional ByVal HighLight As Long)
    ' flat captions only use the Shadow color
    objPicBox.TextColorShadow = Shadow
    If Not Me.FlatCaption Then objPicBox.TextColorLite = HighLight
End Sub
Public Sub GetTextColors(ByVal Shadow As Long, Optional ByVal HighLight As Long)
    Shadow = objPicBox.TextColorShadow
    HighLight = objPicBox.TextColorLite
End Sub

Public Property Let Align(Alignment As etfAlignmentOptions)
    objPicBox.Alignment = Alignment
End Property
Public Property Get Align() As etfAlignmentOptions
    Align = objPicBox.Alignment
End Property

Public Sub Refresh()

    Dim wRect As RECT
    
    Dim hBmp As Long, hOldBmp As Long
    Dim tDC As Long, dDC As Long, lTop As Long, lOffset As Long
    
    Dim pBox As PictureBox
    Dim phWnd As Long, pAutoRedraw As Boolean, pScaleMode As ScaleModeConstants
    Dim pCurrentX As Long, pCurrentY As Long, pForeColor As Long
    
    Dim CaptionCx As Long, CaptionCy As Long
    Dim capX As Long, capY As Long
    Dim xOffset As Long
    Dim sCaption As String
    
    Dim Looper As Long
    Dim lX1 As Long, lX2 As Long
    Dim lY1 As Long, lY2 As Long

    ' has a picturebox been attached?
    If objPicBox.hPointer = 0& Then Exit Sub
    
    CopyMemory pBox, objPicBox, 4&  ' create instance of target
    
    ' are we subclassing it yet?
    If m_WndProc = 0& Then
        ' nope, first Refresh method, let's subclass it now
        m_WndProc = zb_AddressOf(1, 4)
        If Not m_WndProc = 0& Then
            phWnd = pBox.hWnd
            wRect.Right = pBox.ScaleX(pBox.ScaleWidth, pBox.ScaleMode, vbPixels)
            wRect.Bottom = pBox.ScaleY(pBox.ScaleHeight, pBox.ScaleMode, vbPixels)
            CopyMemory pBox, 0&, 4&
            m_WndProc = SetWindowLong(phWnd, GWL_WNDPROC, m_WndProc)
            ' here we will not Refresh, but we will have the zWindowProc refresh it
            InvalidateRect phWnd, wRect, 1
            Exit Sub
        End If
    End If
    
    
    On Error GoTo ExitRoutine
    With pBox
        
        pScaleMode = .ScaleMode     ' cache properties that may be changed for our use
        pAutoRedraw = .AutoRedraw   ' so we can replace them when we are done
        pCurrentX = .CurrentX
        pCurrentY = .CurrentY
        pForeColor = .ForeColor

        .AutoRedraw = True          ' turn this on
        .ScaleMode = vbPixels       ' set scalemode to pixels
    
        phWnd = GetParent(pBox.hWnd)    ' the parent container may not have a ScaleMode property
        GetWindowRect .hWnd, wRect      ' use this method to find the picBox's Left/Top coords
        ScreenToClient phWnd, wRect     ' what are the Left/Top coordinates?
        
        wRect.Right = .ScaleWidth - 1   ' used heavily in drawing the frame
        wRect.Bottom = .ScaleHeight - 1 ' so subtract now vs many subtractions later
    
        ' create temporary bitmap & DC
        dDC = GetDC(0&)
        tDC = CreateCompatibleDC(dDC)
        ' need size of bitmap to include the dimensions and position of our picturebox
        hBmp = CreateCompatibleBitmap(dDC, wRect.Left + .ScaleWidth, wRect.Top + .ScaleHeight)
        ReleaseDC 0&, dDC
        hOldBmp = SelectObject(tDC, hBmp)
    
        ' have form paint its bkg to our temp bitmap (will include windowless controls like labels, image controls)
        SendMessage phWnd, WM_PAINT, tDC, ByVal 0&
        ' copy the portion of the bitmap that is under our picturebox
        BitBlt .hDC, 0, 0, .ScaleWidth, .ScaleHeight, tDC, wRect.Left, wRect.Top, vbSrcCopy
    
        ' Done, clean up
        DeleteObject SelectObject(tDC, hOldBmp)
        DeleteDC tDC
    
        ' calculate width/height of the frame's caption
        sCaption = objPicBox.Caption
        If Len(sCaption) = 0 Then
            If objPicBox.Alignment > tbpAlignTopRight Then ' top aligned
                capY = wRect.Bottom - 1
                CaptionCy = wRect.Bottom - 1
            End If
        Else
            CaptionCx = .TextWidth(sCaption)
            CaptionCy = .TextHeight(sCaption)
            
            ' where text would be drawn if alignment was top/left
            xOffset = FrameEdgeOffset + TextEdgeOffset
            
            ' a little simplified but could be made better with with DrawText API usage
            Do Until CaptionCx < .ScaleWidth - xOffset * 2
                ' truncate caption until we get something close that we can use
                If Len(sCaption) < 6 Then
                    sCaption = Left$(sCaption, 1) & "..."
                    CaptionCx = .TextWidth(sCaption)
                    Exit Do
                End If
                sCaption = Left$(sCaption, Len(sCaption) - 4) & "..."
                CaptionCx = .TextWidth(sCaption)
            Loop
        
            If objPicBox.Alignment > tbpAlignTopRight Then ' top aligned
                capY = wRect.Bottom - CaptionCy
                CaptionCy = wRect.Bottom - CaptionCy \ 2
            Else
                capY = 0
                CaptionCy = CaptionCy \ 2
            End If
    
            Select Case objPicBox.Alignment
                Case tbpAlignTopLeft, tbpAlignBottomLeft
                    capX = xOffset
                Case tbpAlignTopCenter, tbpAlignBottomCenter
                    capX = (wRect.Right - CaptionCx) \ 2
                Case Else
                    capX = wRect.Right - CaptionCx - xOffset
            End Select
        End If
        
        On Error Resume Next    ' just in case user passed an invalid color value
        
                
        For Looper = 1 To objPicBox.BorderStyle
        
            If objPicBox.BorderStyle = tbpFlat Then
                If objPicBox.BorderColorShadow <> -1 Then .ForeColor = objPicBox.BorderColorShadow
            ElseIf Looper = 1 Then
                If objPicBox.BorderColorLite = -1 Then .ForeColor = vb3DHighlight Else .ForeColor = objPicBox.BorderColorLite
                lY1 = 1             ' offsets when drawing the highlight edge
                lX1 = 1
            Else
                If objPicBox.BorderColorShadow = -1 Then .ForeColor = vb3DShadow Else .ForeColor = objPicBox.BorderColorShadow
                lY1 = 0: lY2 = 1    ' offsets when drawing the shadow edge
                lX1 = 0: lX2 = 1
            End If
            ' draw the frame
            If objPicBox.Alignment < tbpAlignBottomLeft Then    ' top aligned vs bottom
                If sCaption = vbNullString Then
                    pBox.Line (lX1, lY1)-(wRect.Right - lX1, lY1)
                Else
                    pBox.Line (capX - TextEdgeOffset, CaptionCy + lY1)-(lX1, CaptionCy + lY1)  ' draw from text to left edge
                    pBox.Line (wRect.Right - lX1 - lX2, CaptionCy + lY1)-(CaptionCx + capX + TextEdgeOffset, CaptionCy + lY1) ' from text to right edge
                End If
                pBox.Line (lX1, CaptionCy + lY1)-(lX1, wRect.Bottom - lY2)              ' left edge
                pBox.Line (0, wRect.Bottom - lY2)-(wRect.Right - lX2, wRect.Bottom - lY2) ' bottom edge
                pBox.Line (wRect.Right - lX2, wRect.Bottom - lY2)-(wRect.Right - lX2, CaptionCy - lY1) ' right edge
            Else
                If sCaption = vbNullString Then
                    pBox.Line (0, CaptionCy + lY1)-(wRect.Right - lX2, CaptionCy + lY1)
                Else
                    pBox.Line (0, CaptionCy + lY1)-(capX - TextEdgeOffset, CaptionCy + lY1)   ' bottom edge
                    pBox.Line (wRect.Right - lX1 - lX2, CaptionCy + lY1)-(CaptionCx + capX + TextEdgeOffset, CaptionCy + lY1)    ' bottom edge
                End If
                pBox.Line (lX1, lY1)-(lX1, CaptionCy + lY1)                                 ' left edge
                pBox.Line (wRect.Right - lX2, 0)-(wRect.Right - lX2, CaptionCy + lY1 + 1) ' right edge
                pBox.Line (lX1, lY1)-(wRect.Right - lX2, lY1)   ' from text to right edge
            End If
        Next

        If Not sCaption = vbNullString Then
            ' set DC position to draw the caption
            If objPicBox.isRaised = True Then
                .CurrentX = capX + 1
                .CurrentY = capY + 1
                If objPicBox.TextColorLite = -1 Then .ForeColor = vb3DHighlight Else .ForeColor = objPicBox.TextColorLite
                pBox.Print sCaption
                
                If objPicBox.TextColorShadow = -1 Then .ForeColor = vbButtonText Else .ForeColor = objPicBox.TextColorShadow
                .CurrentX = capX
                .CurrentY = capY
            Else
                .CurrentX = capX
                .CurrentY = capY
                If objPicBox.TextColorShadow = -1 Then .ForeColor = pForeColor Else .ForeColor = objPicBox.TextColorShadow
            End If
        End If
        pBox.Print sCaption
            
    End With
        


ExitRoutine:
    With pBox
        ' return/reset the original picBox properties
        .AutoRedraw = pAutoRedraw
        .ScaleMode = pScaleMode
        .CurrentX = pCurrentX
        .CurrentY = pCurrentY
        .ForeColor = pForeColor
    End With

    
    ' remove our instance
    CopyMemory pBox, 0&, 4&
    
End Sub

Private Sub ClearSubClassing()

    ' Stops subclassing.
    ' Called by Detach, Attach or zWindowProc
    If Not m_WndProc = 0 Then
        If Not objPicBox.hPointer = 0 Then
            Dim tPicBox As PictureBox
            CopyMemory tPicBox, objPicBox.hPointer, 4&
            SetWindowLong tPicBox.hWnd, GWL_WNDPROC, m_WndProc
            CopyMemory tPicBox, 0&, 4&
        End If
        m_WndProc = 0&
    End If
    zb_Terminate
    objPicBox.hPointer = 0&
    objPicBox.Caption = vbNullString

End Sub


Private Sub Class_Terminate()
    ClearSubClassing
End Sub

'====================================================================================
' THUNKING CALLBACK ROUTINES -- DO NOT MODIFY
'====================================================================================

'-Callback code-----------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 2                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  
  If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
    MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
    Set oCallback = Me                                                      'Then it is me
  End If
  
  nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
  If nAddr = 0 Then
    MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
    Exit Function
  End If
  
  If z_CbMem = 0 Then                                                       'If memory hasn't been allocated
    ReDim z_Cb(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
  End If
  
  If z_Cb(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
    z_Cb(3, nThunkNo) = _
              GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
    z_Cb(4, nThunkNo) = &HBB60E089
    z_Cb(5, nThunkNo) = VarPtr(z_Cb(0, nThunkNo))                           'Set the data address
    z_Cb(6, nThunkNo) = &H73FFC589: z_Cb(7, nThunkNo) = &HC53FF04: z_Cb(8, nThunkNo) = &H7B831F75: z_Cb(9, nThunkNo) = &H20750008: z_Cb(10, nThunkNo) = &HE883E889: z_Cb(11, nThunkNo) = &HB9905004: z_Cb(13, nThunkNo) = &H74FF06E3: z_Cb(14, nThunkNo) = &HFAE2008D: z_Cb(15, nThunkNo) = &H53FF33FF: z_Cb(16, nThunkNo) = &HC2906104: z_Cb(18, nThunkNo) = &H830853FF: z_Cb(19, nThunkNo) = &HD87401F8: z_Cb(20, nThunkNo) = &H4589C031: z_Cb(21, nThunkNo) = &HEAEBFC
  End If
  
  z_Cb(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
  z_Cb(1, nThunkNo) = nAddr                                                 'Set the callback address
  
  If bIdeSafety Then                                                        'If the user wants IDE protection
    z_Cb(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
  End If
    
  z_Cb(12, nThunkNo) = nParamCount                                          'Set the parameter count
  z_Cb(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
  
  nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
  RtlMoveMemory nAddr, VarPtr(z_Cb(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
  zb_AddressOf = nAddr + 16                                                 'Thunk code start address
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Terminate the callback thunks
Private Sub zb_Terminate()
Const MEM_RELEASE As Long = &H8000&                                         'Release allocated memory flag

  If z_CbMem <> 0 Then                                                      'If memory allocated
    If VirtualFree(z_CbMem, 0, MEM_RELEASE) <> 0 Then                       'Release
      z_CbMem = 0                                                           'Indicate memory released
    End If
  End If
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Sub

' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim i     As Long                                                         'Loop index
  Dim j     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, i, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, i, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, i, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  i = i + 4                                                                 'Bump to the next entry
  j = i + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While i < j
    RtlMoveMemory VarPtr(nAddr), i, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), i - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    i = i + 4                                                             'Next vTable entry
  Loop
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function

' BUTTON TEMPLATE CODE\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*************************************************************************************************
'* Callbacks - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************
'Callback ordinal 3

'Callback ordinal 2

'Callback ordinal 1
Private Function zWindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo EH

    Select Case uMsg
    Case WM_PAINT
        If wParam = 0& Then
            Dim uRect As RECT
            
            ' what part of the picturebox was updated?
            GetUpdateRect hWnd, uRect, 0&
            ' forward the paint event
            zWindowProc = CallWindowProc(m_WndProc, hWnd, uMsg, wParam, lParam)
            
            uMsg = 0&   ' reset & use as flag to refresh if needed
            
            If objPicBox.RenderedX < uRect.Right Then
                ' we haven't painted the entire picbox yet
                objPicBox.RenderedX = uRect.Right   ' update how much is painted
                uMsg = WM_PAINT
            ElseIf objPicBox.RenderedX > uRect.Right Then
                ' if picbox was resized, our cached flags may be wrong
                objPicBox.RenderedX = uRect.Right
            End If
            
            If objPicBox.RenderedY < uRect.Bottom Then
                ' we haven't painted the entire picbox yet
                objPicBox.RenderedY = uRect.Bottom  ' update how much is painted
                uMsg = WM_PAINT
            ElseIf objPicBox.RenderedY > uRect.Bottom Then
                ' if picbox was resized, our cached flags may be wrong
                objPicBox.RenderedY = uRect.Bottom
            End If
            
            If uMsg = WM_PAINT Then Me.Refresh ' refresh our picbox
            Exit Function
            
        End If
        
    Case WM_SIZE
        Dim sRect As RECT
        zWindowProc = CallWindowProc(m_WndProc, hWnd, uMsg, wParam, lParam)
        GetClientRect hWnd, sRect
        If sRect.Right < objPicBox.RenderedX Then
            Me.Refresh
        ElseIf sRect.Bottom < objPicBox.RenderedY Then
            Me.Refresh
        End If
        Exit Function
        
    Case WM_DESTROY
        zWindowProc = CallWindowProc(m_WndProc, hWnd, uMsg, wParam, lParam)
        SetWindowLong hWnd, GWL_WNDPROC, m_WndProc
        m_WndProc = 0&
        Exit Function
        
    Case WM_MOVE, WM_DISPLAYCHANGE, WM_SYSCOLORCHANGE, WM_WININICHANGE
        zWindowProc = CallWindowProc(m_WndProc, hWnd, uMsg, wParam, lParam)
        Refresh
        Exit Function
    End Select
EH:
zWindowProc = CallWindowProc(m_WndProc, hWnd, uMsg, wParam, lParam)
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the callbacks will not be fired & maybe GPF.
' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
End Function



