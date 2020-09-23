VERSION 5.00
Begin VB.UserControl iCatcherLabel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3045
   FillStyle       =   0  'Solid
   ScaleHeight     =   87
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   203
   ToolboxBitmap   =   "iCatcherLabel.ctx":0000
   Begin VB.Image imgCustomPic 
      Height          =   480
      Left            =   1440
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgNext 
      Height          =   480
      Left            =   960
      Picture         =   "iCatcherLabel.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgSuccess 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   0
      Picture         =   "iCatcherLabel.ctx":0BDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgFailed 
      Height          =   480
      Left            =   480
      Picture         =   "iCatcherLabel.ctx":14A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "iCatcherLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'+  File Description:
'       iCatcherLabel - Enhanced Status and Label Control
'
'   Product Name:
'       iCatcherLabel.ctl
'
'   Compatability:
'       Windows: 98, ME, NT, 2000, XP
'
'   Software Developed by:
'       Paul R. Territo, Ph.D
'
'   Based on the following On-Line Articles
'       (isButton - Fred.cpp)
'           URL: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=56053&lngWId=1
'       (SelfSubclasser - Paul Caton)
'           URL: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=54117&lngWId=1
'
'   Legal Copyright & Trademarks:
'       Copyright © 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'       Trademark ™ 2005, by Paul R. Territo, Ph.D, All Rights Reserved Worldwide
'
'   Comments:
'       No claims or warranties are expressed or implied as to accuracy or fitness
'       for use of this software. Paul R. Territo, Ph.D shall not be liable
'       for any incidental or consequential damages suffered by any use of
'       this  software. This software is owned by Paul R. Territo, Ph.D and is
'       sold for use as a license in accordance with the terms of the License
'       Agreement in the accompanying the documentation.
'
'       As a technical note, there are a couple of residual routines in this control
'       which I left for the develoepr to play with. These routines will make the
'       development of custom drawing easier, and could be removed if size is a premium.
'
'       Lastly, a huge thanks to Fred.cpp (Drawing Routines) and Paul Caton (SelfSubclasser)
'       for their very nice examples. This project would not have the look and feel
'       if it were not for these two programmers. Also, I want to thank Paul Turcksin
'       for his review of this control prior to release.
'
'   Contact Information:
'       For Technical Assistance:
'       Email: pwterrito@insightbb.com
'
'-  Modification(s) History:
'       27Aug05 - Initial test harness and usercontrol finished
'       10Sep05 - Fixed Rectangular shape bug which caused the Icon to
'                 be mis-alligned. Optimized the Shading models and the
'                 code for drawing Rectangular Gradients.
'               - Eliminated all extraneous code not used by the control.
'       16Sep05 - Added additional comments to the test harness and control to
'                 ensure clarity of the properties and how to use them...
'       23Sep05 - Cleaned up and added additional documnetation for each
'                 properties, variables, and sub/function.
'
'   Force Declarations
Option Explicit

'   32Bit Windows API Declarations
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINT, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINT) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal hPalette As Long, pccolorref As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, lpPoint As POINT) As Long
'*************************************************************
'
'   Private Constants
'
'**************************************
'Auxiliary Constants
Private Const COLOR_BTNFACE             As Long = 15
Private Const COLOR_BTNSHADOW           As Long = 16
Private Const COLOR_BTNTEXT             As Long = &H800000 '18
Private Const COLOR_HIGHLIGHT           As Long = 13
Private Const COLOR_WINDOW              As Long = 5
Private Const COLOR_INFOTEXT            As Long = 23
Private Const COLOR_INFOBK              As Long = 24
Private Const BDR_RAISEDOUTER           As Long = &H1
Private Const BDR_SUNKENOUTER           As Long = &H2
Private Const BDR_RAISEDINNER           As Long = &H4
Private Const BDR_SUNKENINNER           As Long = &H8
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const BF_LEFT                   As Long = &H1
Private Const BF_TOP                    As Long = &H2
Private Const BF_RIGHT                  As Long = &H4
Private Const BF_BOTTOM                 As Long = &H8
Private Const BF_RECT                   As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const AC_SRC_ALPHA              As Long = &H1
Private Const DIB_RGB_COLORS            As Long = 0

'Windows Messages
Private Const GWL_STYLE                 As Long = -16           'Global Window Style
Private Const WS_CAPTION                As Long = &HC00000      'Window Style Caption
Private Const WS_THICKFRAME             As Long = &H40000       'Window Style Frame
Private Const WS_SYSMENU                As Long = &H80000       'Window Style System Menu
Private Const WS_MINIMIZEBOX            As Long = &H20000       'Window Style Min Box
Private Const SWP_REFRESH               As Long = (&H1 Or &H2 Or &H4 Or &H20) 'Window Style AutoRedraw
Private Const WS_EX_TOOLWINDOW          As Long = &H80          'Window Style Extender ToolTipText
Private Const GWL_EXSTYLE               As Long = -20           'Global Window Style of Translucency (XP & NT Only)
Private Const SW_SHOWDEFAULT            As Long = 10
Private Const SW_SHOWMAXIMIZED          As Long = 3
Private Const SW_SHOWMINIMIZED          As Long = 2
Private Const SW_SHOWMINNOACTIVE        As Long = 7
Private Const SW_SHOWNA                 As Long = 8
Private Const SW_SHOWNOACTIVATE         As Long = 4
Private Const SW_SHOWNORMAL             As Long = 1
Private Const SWP_FRAMECHANGED          As Long = &H20
Private Const SWP_DRAWFRAME             As Long = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW            As Long = &H80
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_NOCOPYBITS            As Long = &H100
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOOWNERZORDER         As Long = &H200
Private Const SWP_NOREDRAW              As Long = &H8
Private Const SWP_NOREPOSITION          As Long = SWP_NOOWNERZORDER
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOZORDER              As Long = &H4
Private Const SWP_SHOWWINDOW            As Long = &H40
Private Const HWND_TOPMOST              As Long = -&H1
Private Const CW_USEDEFAULT             As Long = &H80000000
Private Const RGN_AND                   As Long = &H1           'Combine two regions
Private Const RGN_OR                    As Long = &H2
Private Const RGN_XOR                   As Long = &H3
Private Const RGN_DIFF                  As Long = &H4
Private Const RGN_COPY                  As Long = &H5
Private Const DST_BITMAP                As Long = &H4
Private Const DST_COMPLEX               As Long = &H0
Private Const DST_ICON                  As Long = &H3
Private Const DSS_MONO                  As Long = &H80
Private Const DSS_NORMAL                As Long = &H0
Private Const NULLREGION                As Long = &H1           'Empty region
Private Const SIMPLEREGION              As Long = &H2           'Rectangle Region
Private Const COMPLEXREGION             As Long = &H3           'The region is complex

'Constants for nPolyFillMode in CreatePolygonRgn y CreatePolyPolygonRgn:
Private Const ALTERNATE                 As Long = 1
Private Const WINDING                   As Long = 2

'   Private Types
Private Type POINT
    X As Long                   'X Position for API Calls
    Y As Long                   'Y Position for API Calls
End Type

Private Type RECT
    Left As Long                'Left Coordinates for API Calls
    Top As Long                 'Top Coordinates for API Calls
    Right As Long               'Right Coordinates for API Calls
    Bottom As Long              'Botton Coordinates for API Calls
End Type

'   Public Enumerations
Public Enum ImageType           '[Button Icon Constants]
    [clbNext] = &H0             'Arrow Button Icon  -> Blue Arrow
    [clbSuccess] = &H1          'Check Button Icon  -> Green Check Mark
    [clbFailed] = &H2           'Failed Button Icon -> Red X
    [clbCustom] = &H3           'Custom Button Icon -> Anything the User Sets
End Enum

Public Enum clAlign             '[Text Alignment Constants]
    [clCenter] = &H0            'Text Alignment is Centered
    [clLeft] = &H1              'Text Alignment is Left
    [clRight] = &H2             'Text Alignment is Right
    [clTop] = &H3               'Text Alignment is Top
    [clBottom] = &H4            'Text Alignment is Bottom
End Enum

Public Enum clbAlign            '[Button Alignment Constants]
    [clbLeft] = &H0             'Button Alignment is Left
    [clbRight] = &H1            'Button Alignment is Right
End Enum

Public Enum clbShape            '[Background Shape Constants]
    [clbEllipse] = &H0          'Elliptical Button Background Shape
    [clbRndRect] = &H1          'Rounded Rectangle Button Background Shape
    [clbRectangle] = &H2        'Rectangle Button Background Shape
End Enum

'   Private Enumerations
Private Enum clState            '[Usercontrol State Constants]
    [clNormal] = &H1            'iCatcherLabel State = Normal
    [clHot] = &H2               'iCatcherLabel State = Hot (Hovering)
    [clPressed] = &H3           'iCatcherLabel State = Pressed
    [clDisabled] = &H4          'iCatcherLabel State = Disabled
    [clDefault] = &H5           'iCatcherLabel State = Default = Normal
End Enum

Private Enum DrawTextFlags      '[DrawTextFlags Enumerations]
    DT_TOP = &H0                'Align Top
    DT_LEFT = &H0               'Align Left
    DT_CENTER = &H1             'Align Center
    DT_RIGHT = &H2              'Align Right
    DT_VCENTER = &H4            'Align Vertically Centered
    DT_BOTTOM = &H8             'Align Bottom
    DT_WORDBREAK = &H10         'Permit WordBreaks
    DT_SINGLELINE = &H20        'Use Only One Line
    DT_EXPANDTABS = &H40        '?
    DT_TABSTOP = &H80           'Use Tab Stops
    DT_NOCLIP = &H100           'No Clipping
    DT_EXTERNALLEADING = &H200  'Use External Leading
    DT_CALCRECT = &H400         'Calculate the RECT for the Text
    DT_NOPREFIX = &H800         'No Prefixes
    DT_INTERNAL = &H1000        'Internal Text
    DT_EDITCONTROL = &H2000     '?
    DT_PATH_ELLIPSIS = &H4000   'Used for Curved Text?
    DT_END_ELLIPSIS = &H8000    'Used for Curved Text?
    DT_MODIFYSTRING = &H10000   '?
    DT_RTLREADING = &H20000     '?
    DT_WORD_ELLIPSIS = &H40000   'Used for Curved Text?
    DT_NOFULLWIDTHCHARBREAK = &H80000   '?
    DT_HIDEPREFIX = &H100000    '?
    DT_PREFIXONLY = &H200000    '?
End Enum

'   Private variables
Private m_Enabled                   As Boolean      'Control Dis/Enabled Property
Private m_UseCustomColors           As Boolean      'UseCustomColor Property Flag
Private m_ButtonAlign               As clbAlign     'Button Alignment Property
Private m_ButtonShape               As clbShape     'Button Background Shape Property
Private m_CaptionAlign              As clAlign      'Caption Alignment Property
Private m_CornerSize                As Long         'Corner Size Property (Used only with clbRndRect Shape)
Private m_ctlRect                   As RECT         'RECT for the Control
Private m_btnRect                   As RECT         'RECT for the Button Background Shape
Private m_Font                      As StdFont      'Font Property
Private m_Icon                      As StdPicture   'Icon Property
Private m_ButtonBackSize            As Long         'Background Shape Size (Min = 33)
Private m_ButtonIcon                As ImageType    'Button Icon Image Property
Private m_ButtonToolTipText         As String       'Button ToolTipText String
Private m_iState                    As clState      'Usercontrol State
Private m_BackColor                 As Long         'Usercontrol Backcolor
Private m_lButtonRgn                As Long         'Button Background Shape Region
Private m_FontColor                 As Long         'ForeColor(Text Color) when m_iState = clNormal or clDefault
Private m_FontHighlightColor        As Long         'ForeColor(Text Color) when m_iState = clHot
Private m_CustomColor               As Long         'CustomColor Property (only active when UseCustomColor = True)
Private m_lRegion                   As Long         'Label Shaped Region
Private m_lwFontAlign               As Long         'Text Alignment Variable used with API calls
Private m_PrevImage                 As ImageType    'Previous Icon Holder - Used when we change icons on the fly..
Private m_Caption                   As String       'Caption property
Private m_txtRect                   As RECT         'RECT for the Text section of the control
Private m_UseCustomColor            As Boolean      'UseCustomColor Property -> Sets the CustomColor as the active color
Private m_UseCustomIcon             As Boolean      'UseCustomIcon Property -> Sets the CustomIcon as the active button icon
Private m_UseGradient               As Boolean      'UseGradient Property -> (False = Flat, True = Convexed)
Private m_Pnt                       As POINT        'Current Point used by API Functions for HitTest

'   Public Events (not Captured by SelfSubclassing Template)
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Hover(X As Single, Y As Single)
Public Event ButtonClick()
Public Event ButtonDblClick()
Public Event ButtonMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonMouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ButtonHover(X As Single, Y As Single)
'==================================================================================================
' ucSubclass - A sample UserControl demonstrating self-subclassing
'
' Paul_Caton@hotmail.com
' Copyright free, use and abuse as you see fit.
'
' v1.0.0000 20040525 First cut.....................................................................
' v1.1.0000 20040602 Multi-subclassing version.....................................................
' v1.1.0001 20040604 Optimized the subclass code...................................................
' v1.1.0002 20040607 Substituted byte arrays for strings for the code buffers......................
' v1.1.0003 20040618 Re-patch when adding extra hWnds..............................................
' v1.1.0004 20040619 Optimized to death version....................................................
' v1.1.0005 20040620 Use allocated memory for code buffers, no need to re-patch....................
' v1.1.0006 20040628 Better protection in zIdx, improved comments..................................
' v1.1.0007 20040629 Fixed InIDE patching oops.....................................................
' v1.1.0008 20040910 Fixed bug in UserControl_Terminate, zSubclass_Proc procedure hidden...........

'==================================================================================================

'   Public SelfSubclasser Events
Public Event MouseEnter()
Public Event MouseLeave()
Public Event Status(ByVal sStatus As String)

'   Private Windows Message Constants
Private Const WM_EXITSIZEMOVE           As Long = &H232
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_MOUSELEAVE             As Long = &H2A3
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_MOVING                 As Long = &H216
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_SIZING                 As Long = &H214
Private Const WM_SYSCOLORCHANGE         As Long = &H15
Private Const WM_THEMECHANGED           As Long = &H31A
Private Const WM_USER                   As Long = &H400

'   Private Mouse Tracking Enums
Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Type TRACKMOUSEEVENT_STRUCT
  cbSize                             As Long
  dwFlags                            As TRACKMOUSEEVENT_FLAGS
  hwndTrack                          As Long
  dwHoverTime                        As Long
End Type

'   Private SelfSubclasser Variables
Private bTrack                       As Boolean
Private bTrackUser32                 As Boolean
Private bInCtrl                      As Boolean

'   Private SelfSubclassing Win32 API Declares
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long
Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long
Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

'==================================================================================================
'   SelfSubclassing Declarations
Private Enum eMsgWhen
  MSG_AFTER = 1                                                                         'Message calls back after the original (previous) WndProc
  MSG_BEFORE = 2                                                                        'Message calls back before the original (previous) WndProc
  MSG_BEFORE_AND_AFTER = MSG_AFTER Or MSG_BEFORE                                        'Message calls back before and after the original (previous) WndProc
End Enum

Private Const ALL_MESSAGES           As Long = -1                                       'All messages added or deleted
Private Const GMEM_FIXED             As Long = 0                                        'Fixed memory GlobalAlloc flag
Private Const GWL_WNDPROC            As Long = -4                                       'Get/SetWindow offset to the WndProc procedure address
Private Const PATCH_04               As Long = 88                                       'Table B (before) address patch offset
Private Const PATCH_05               As Long = 93                                       'Table B (before) entry count patch offset
Private Const PATCH_08               As Long = 132                                      'Table A (after) address patch offset
Private Const PATCH_09               As Long = 137                                      'Table A (after) entry count patch offset

Private Type tSubData                                                                   'Subclass data type
  hwnd                               As Long                                            'Handle of the window being subclassed
  nAddrSub                           As Long                                            'The address of our new WndProc (allocated memory).
  nAddrOrig                          As Long                                            'The address of the pre-existing WndProc
  nMsgCntA                           As Long                                            'Msg after table entry count
  nMsgCntB                           As Long                                            'Msg before table entry count
  aMsgTblA()                         As Long                                            'Msg after table array
  aMsgTblB()                         As Long                                            'Msg Before table array
End Type

Private sc_aSubData()                As tSubData                                        'Subclass data array

'   Private SelfSubclassing Win32 API Declares
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'==================================================================================================

'Determine if the passed function is supported
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
  Dim hmod        As Long
  Dim bLibLoaded  As Boolean

  hmod = GetModuleHandleA(sModule)

  If hmod = 0 Then
    hmod = LoadLibraryA(sModule)
    If hmod Then
      bLibLoaded = True
    End If
  End If

  If hmod Then
    If GetProcAddress(hmod, sFunction) Then
      IsFunctionExported = True
    End If
  End If

  If bLibLoaded Then
    Call FreeLibrary(hmod)
  End If
End Function

'======================================================================================================
'Subclass code - The programmer may call any of the following Subclass_??? routines

'Add a message to the table of those that will invoke a callback. You should Subclass_Start first and then add the messages
Private Sub Subclass_AddMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    'lng_hWnd  - The handle of the window for which the uMsg is to be added to the callback table
    'uMsg      - The message number that will invoke a callback. NB Can also be ALL_MESSAGES, ie all messages will callback
    'When      - Whether the msg is to callback before, after or both with respect to the the default (previous) handler
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zAddMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zAddMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Delete a message from the table of those that will invoke a callback.
Private Sub Subclass_DelMsg(ByVal lng_hWnd As Long, ByVal uMsg As Long, Optional ByVal When As eMsgWhen = MSG_AFTER)
    'Parameters:
    '   lng_hWnd  - The handle of the window for which the uMsg is to be removed from the callback table
    '   uMsg      - The message number that will be removed from the callback table. NB Can also be ALL_MESSAGES, ie all messages will callback
    '   When      - Whether the msg is to be removed from the before, after or both callback tables
    With sc_aSubData(zIdx(lng_hWnd))
        If When And eMsgWhen.MSG_BEFORE Then
            Call zDelMsg(uMsg, .aMsgTblB, .nMsgCntB, eMsgWhen.MSG_BEFORE, .nAddrSub)
        End If
        If When And eMsgWhen.MSG_AFTER Then
            Call zDelMsg(uMsg, .aMsgTblA, .nMsgCntA, eMsgWhen.MSG_AFTER, .nAddrSub)
        End If
    End With
End Sub

'Return whether we're running in the IDE.
Private Function Subclass_InIDE() As Boolean
  Debug.Assert zSetTrue(Subclass_InIDE)
End Function

'Start subclassing the passed window handle
Private Function Subclass_Start(ByVal lng_hWnd As Long) As Long
    'Parameters:
    '   lng_hWnd  - The handle of the window to be subclassed
    'Returns;
    '   The sc_aSubData() index
    Const CODE_LEN              As Long = 200                                             'Length of the machine code in bytes
    Const FUNC_CWP              As String = "CallWindowProcA"                             'We use CallWindowProc to call the original WndProc
    Const FUNC_EBM              As String = "EbMode"                                      'VBA's EbMode function allows the machine code thunk to know if the IDE has stopped or is on a breakpoint
    Const FUNC_SWL              As String = "SetWindowLongA"                              'SetWindowLongA allows the cSubclasser machine code thunk to unsubclass the subclasser itself if it detects via the EbMode function that the IDE has stopped
    Const MOD_USER              As String = "user32"                                      'Location of the SetWindowLongA & CallWindowProc functions
    Const MOD_VBA5              As String = "vba5"                                        'Location of the EbMode function if running VB5
    Const MOD_VBA6              As String = "vba6"                                        'Location of the EbMode function if running VB6
    Const PATCH_01              As Long = 18                                              'Code buffer offset to the location of the relative address to EbMode
    Const PATCH_02              As Long = 68                                              'Address of the previous WndProc
    Const PATCH_03              As Long = 78                                              'Relative address of SetWindowsLong
    Const PATCH_06              As Long = 116                                             'Address of the previous WndProc
    Const PATCH_07              As Long = 121                                             'Relative address of CallWindowProc
    Const PATCH_0A              As Long = 186                                             'Address of the owner object
    Static Abuf(1 To CODE_LEN)  As Byte                                                   'Static code buffer byte array
    Static pCWP                 As Long                                                   'Address of the CallWindowsProc
    Static pEbMode              As Long                                                   'Address of the EbMode IDE break/stop/running function
    Static pSWL                 As Long                                                   'Address of the SetWindowsLong function
    Dim i                       As Long                                                   'Loop index
    Dim j                       As Long                                                   'Loop index
    Dim nSubIdx                 As Long                                                   'Subclass data index
    Dim sHex                    As String                                                 'Hex code string
  
    '   If it's the first time through here..
    If Abuf(1) = 0 Then
  
        '   The hex pair machine code representation.
        sHex = "5589E583C4F85731C08945FC8945F8EB0EE80000000083F802742185C07424E830000000837DF800750AE838000000E84D00" & _
               "00005F8B45FCC9C21000E826000000EBF168000000006AFCFF7508E800000000EBE031D24ABF00000000B900000000E82D00" & _
               "0000C3FF7514FF7510FF750CFF75086800000000E8000000008945FCC331D2BF00000000B900000000E801000000C3E33209" & _
               "C978078B450CF2AF75278D4514508D4510508D450C508D4508508D45FC508D45F85052B800000000508B00FF90A4070000C3"
    
        '   Convert the string from hex pairs to bytes and store in the static machine code buffer
        i = 1
        Do While j < CODE_LEN
            j = j + 1
            Abuf(j) = Val("&H" & Mid$(sHex, i, 2))                                            'Convert a pair of hex characters to an eight-bit value and store in the static code buffer array
            i = i + 2
        Loop                                                                                'Next pair of hex characters
        
        '   Get API function addresses
        If Subclass_InIDE Then                                                              'If we're running in the VB IDE
            Abuf(16) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            Abuf(17) = &H90                                                                   'Patch the code buffer to enable the IDE state code
            pEbMode = zAddrFunc(MOD_VBA6, FUNC_EBM)                                           'Get the address of EbMode in vba6.dll
            If pEbMode = 0 Then                                                               'Found?
                pEbMode = zAddrFunc(MOD_VBA5, FUNC_EBM)                                         'VB5 perhaps
            End If
        End If
        
        pCWP = zAddrFunc(MOD_USER, FUNC_CWP)                                                'Get the address of the CallWindowsProc function
        pSWL = zAddrFunc(MOD_USER, FUNC_SWL)                                                'Get the address of the SetWindowLongA function
        ReDim sc_aSubData(0 To 0) As tSubData                                               'Create the first sc_aSubData element
    Else
        nSubIdx = zIdx(lng_hWnd, True)
        If nSubIdx = -1 Then                                                                'If an sc_aSubData element isn't being re-cycled
          nSubIdx = UBound(sc_aSubData()) + 1                                               'Calculate the next element
          ReDim Preserve sc_aSubData(0 To nSubIdx) As tSubData                              'Create a new sc_aSubData element
        End If
        
        Subclass_Start = nSubIdx
    End If

    With sc_aSubData(nSubIdx)
        .hwnd = lng_hWnd                                                                    'Store the hWnd
        .nAddrSub = GlobalAlloc(GMEM_FIXED, CODE_LEN)                                       'Allocate memory for the machine code WndProc
        .nAddrOrig = SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrSub)                          'Set our WndProc in place
        Call RtlMoveMemory(ByVal .nAddrSub, Abuf(1), CODE_LEN)                              'Copy the machine code from the static byte array to the code array in sc_aSubData
        Call zPatchRel(.nAddrSub, PATCH_01, pEbMode)                                        'Patch the relative address to the VBA EbMode api function, whether we need to not.. hardly worth testing
        Call zPatchVal(.nAddrSub, PATCH_02, .nAddrOrig)                                     'Original WndProc address for CallWindowProc, call the original WndProc
        Call zPatchRel(.nAddrSub, PATCH_03, pSWL)                                           'Patch the relative address of the SetWindowLongA api function
        Call zPatchVal(.nAddrSub, PATCH_06, .nAddrOrig)                                     'Original WndProc address for SetWindowLongA, unsubclass on IDE stop
        Call zPatchRel(.nAddrSub, PATCH_07, pCWP)                                           'Patch the relative address of the CallWindowProc api function
        Call zPatchVal(.nAddrSub, PATCH_0A, ObjPtr(Me))                                     'Patch the address of this object instance into the static machine code buffer
    End With
End Function

'   Stop subclassing the passed window handle
Private Sub Subclass_Stop(ByVal lng_hWnd As Long)
    'Parameters:
    '   lng_hWnd  - The handle of the window to stop being subclassed
    With sc_aSubData(zIdx(lng_hWnd))
        Call SetWindowLongA(.hwnd, GWL_WNDPROC, .nAddrOrig)                                 'Restore the original WndProc
        Call zPatchVal(.nAddrSub, PATCH_05, 0)                                              'Patch the Table B entry count to ensure no further 'before' callbacks
        Call zPatchVal(.nAddrSub, PATCH_09, 0)                                              'Patch the Table A entry count to ensure no further 'after' callbacks
        Call GlobalFree(.nAddrSub)                                                          'Release the machine code memory
        .hwnd = 0                                                                           'Mark the sc_aSubData element as available for re-use
        .nMsgCntB = 0                                                                       'Clear the before table
        .nMsgCntA = 0                                                                       'Clear the after table
        Erase .aMsgTblB                                                                     'Erase the before table
        Erase .aMsgTblA                                                                     'Erase the after table
    End With
End Sub

'   Stop all subclassing
Private Sub Subclass_StopAll()
  Dim i As Long
  
    i = UBound(sc_aSubData())                                                             'Get the upper bound of the subclass data array
    Do While i >= 0                                                                       'Iterate through each element
        With sc_aSubData(i)
            If .hwnd <> 0 Then                                                                'If not previously Subclass_Stop'd
                Call Subclass_Stop(.hwnd)                                                       'Subclass_Stop
            End If
        End With
        i = i - 1                                                                           'Next element
    Loop
End Sub

'   Track the mouse leaving the indicated window
Private Sub TrackMouseLeave(ByVal lng_hWnd As Long)
  Dim tme As TRACKMOUSEEVENT_STRUCT
  
    If bTrack Then
        With tme
            .cbSize = Len(tme)
            .dwFlags = TME_LEAVE
            .hwndTrack = lng_hWnd
        End With
    
        If bTrackUser32 Then
            Call TrackMouseEvent(tme)
        Else
            Call TrackMouseEventComCtl(tme)
        End If
    End If
End Sub

'======================================================================================================
'Subclass handler - MUST be the first Public routine in this file. That includes public properties also
Public Sub zSubclass_Proc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByRef lng_hWnd As Long, ByRef uMsg As Long, ByRef wParam As Long, ByRef lParam As Long)
    'Parameters:
    '   bBefore  - Indicates whether the the message is being processed before or after the default handler - only really needed if a message is set to callback both before & after.
    '   bHandled - Set this variable to True in a 'before' callback to prevent the message being subsequently processed by the default handler... and if set, an 'after' callback
    '   lReturn  - Set this variable as per your intentions and requirements, see the MSDN documentation for each individual message value.
    '   hWnd     - The window handle
    '   uMsg     - The message number
    '   wParam   - Message related data
    '   lParam   - Message related data
    'Notes:
    '   If you really know what you're doing, it's possible to change the values of the
    '   hWnd, uMsg, wParam and lParam parameters in a 'before' callback so that different
    '   values get passed to the default handler.. and optionaly, the 'after' callback
    Static bMoving As Boolean
  
      Select Case uMsg
        Case WM_MOUSEMOVE
            If Not bInCtrl Then
                bInCtrl = True
                Call TrackMouseLeave(lng_hWnd)
                m_iState = clHot
                Refresh
                RaiseEvent MouseEnter
            Else
                '   Get the Cursor Position
                Call GetCursorPos(m_Pnt)
                '   Convert coordinates
                Call ScreenToClient(UserControl.hwnd, m_Pnt)
                '   See if we are over the Button or Label Regions
                If PtInRegion(m_lButtonRgn, m_Pnt.X, m_Pnt.Y) Then
                    RaiseEvent ButtonHover(CSng(m_Pnt.X), CSng(m_Pnt.Y))
                Else
                    RaiseEvent Hover(CSng(m_Pnt.X), CSng(m_Pnt.Y))
                End If
            End If
            
        Case WM_MOUSELEAVE
            bInCtrl = False
            m_iState = clNormal
            Refresh
            RaiseEvent MouseLeave
        
        Case WM_MOVING
            bMoving = True
            RaiseEvent Status("Control is moving...")
        
        Case WM_SIZING
            bMoving = False
            RaiseEvent Status("Control is sizing...")
        
        Case WM_EXITSIZEMOVE
            RaiseEvent Status("Finished " & IIf(bMoving, "moving.", "sizing."))
        
      End Select
End Sub

'======================================================================================================
'   These z??? routines are exclusively called by the Subclass_??? routines.

'   Worker sub for Subclass_AddMsg
Private Sub zAddMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry  As Long                                                                   'Message table entry index
    Dim nOff1   As Long                                                                   'Machine code buffer offset 1
    Dim nOff2   As Long                                                                   'Machine code buffer offset 2
  
    If uMsg = ALL_MESSAGES Then                                                           'If all messages
        nMsgCnt = ALL_MESSAGES                                                              'Indicates that all messages will callback
    Else                                                                                  'Else a specific message number
        Do While nEntry < nMsgCnt                                                           'For each existing entry. NB will skip if nMsgCnt = 0
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = 0 Then                                                       'This msg table slot is a deleted entry
                aMsgTbl(nEntry) = uMsg                                                          'Re-use this entry
                Exit Sub                                                                        'Bail
            ElseIf aMsgTbl(nEntry) = uMsg Then                                                'The msg is already in the table!
                Exit Sub                                                                        'Bail
            End If
        Loop                                                                                'Next entry
    
        nMsgCnt = nMsgCnt + 1                                                               'New slot required, bump the table entry count
        ReDim Preserve aMsgTbl(1 To nMsgCnt) As Long                                        'Bump the size of the table.
        aMsgTbl(nMsgCnt) = uMsg                                                             'Store the message number in the table
    End If

    If When = eMsgWhen.MSG_BEFORE Then                                                    'If before
        nOff1 = PATCH_04                                                                    'Offset to the Before table
        nOff2 = PATCH_05                                                                    'Offset to the Before table entry count
    Else                                                                                  'Else after
        nOff1 = PATCH_08                                                                    'Offset to the After table
        nOff2 = PATCH_09                                                                    'Offset to the After table entry count
    End If

    If uMsg <> ALL_MESSAGES Then
        Call zPatchVal(nAddr, nOff1, VarPtr(aMsgTbl(1)))                                    'Address of the msg table, has to be re-patched because Redim Preserve will move it in memory.
    End If
    Call zPatchVal(nAddr, nOff2, nMsgCnt)                                                 'Patch the appropriate table entry count
End Sub

'   Return the memory address of the passed function in the passed dll
Private Function zAddrFunc(ByVal sDLL As String, ByVal sProc As String) As Long
    zAddrFunc = GetProcAddress(GetModuleHandleA(sDLL), sProc)
    Debug.Assert zAddrFunc                                                                'You may wish to comment out this line if you're using vb5 else the EbMode GetProcAddress will stop here everytime because we look for vba6.dll first
End Function

'   Worker sub for Subclass_DelMsg
Private Sub zDelMsg(ByVal uMsg As Long, ByRef aMsgTbl() As Long, ByRef nMsgCnt As Long, ByVal When As eMsgWhen, ByVal nAddr As Long)
    Dim nEntry As Long
    
    If uMsg = ALL_MESSAGES Then                                                           'If deleting all messages
        nMsgCnt = 0                                                                         'Message count is now zero
        If When = eMsgWhen.MSG_BEFORE Then                                                  'If before
          nEntry = PATCH_05                                                                 'Patch the before table message count location
        Else                                                                                'Else after
          nEntry = PATCH_09                                                                 'Patch the after table message count location
        End If
        Call zPatchVal(nAddr, nEntry, 0)                                                    'Patch the table message count to zero
    Else                                                                                  'Else deleteting a specific message
        Do While nEntry < nMsgCnt                                                           'For each table entry
            nEntry = nEntry + 1
            If aMsgTbl(nEntry) = uMsg Then                                                    'If this entry is the message we wish to delete
                aMsgTbl(nEntry) = 0                                                             'Mark the table slot as available
                Exit Do                                                                         'Bail
            End If
        Loop                                                                                'Next entry
    End If
End Sub

'   Get the sc_aSubData() array index of the passed hWnd
Private Function zIdx(ByVal lng_hWnd As Long, Optional ByVal bAdd As Boolean = False) As Long
    '   Get the upper bound of sc_aSubData() - If you get an error here, you're probably Subclass_AddMsg-ing before Subclass_Start
    zIdx = UBound(sc_aSubData)
    Do While zIdx >= 0                                                                    'Iterate through the existing sc_aSubData() elements
        With sc_aSubData(zIdx)
            If .hwnd = lng_hWnd Then                                                          'If the hWnd of this element is the one we're looking for
                If Not bAdd Then                                                                'If we're searching not adding
                    Exit Function                                                                 'Found
                End If
            ElseIf .hwnd = 0 Then                                                             'If this an element marked for reuse.
                If bAdd Then                                                                    'If we're adding
                    Exit Function                                                                 'Re-use it
                End If
            End If
        End With
        zIdx = zIdx - 1                                                                     'Decrement the index
    Loop
  
  If Not bAdd Then
        Debug.Assert False                                                                  'hWnd not found, programmer error
  End If

'   If we exit here, we're returning -1, no freed elements were found
End Function

'   Patch the machine code buffer at the indicated offset with the relative address to the target address.
Private Sub zPatchRel(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nTargetAddr As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nTargetAddr - nAddr - nOffset - 4, 4)
End Sub

'   Patch the machine code buffer at the indicated offset with the passed value
Private Sub zPatchVal(ByVal nAddr As Long, ByVal nOffset As Long, ByVal nValue As Long)
    Call RtlMoveMemory(ByVal nAddr + nOffset, nValue, 4)
End Sub

'   Worker function for Subclass_InIDE
Private Function zSetTrue(ByRef bValue As Boolean) As Boolean
    zSetTrue = True
    bValue = True
End Function

'*******************************************************************************
'   End Subclasser Section - Start Usercontrol Sections
'*******************************************************************************

Private Sub APIFillRect(ByVal hdc As Long, RC As RECT, ByVal Color As Long)
    On Error GoTo APIFillRect_Error
    
    '   Fill a Rectangular region using its hDC and the color
    '   passed by the caller....sometime called "indirect" in
    '   API nomincature.
    
    Dim NewBrush As Long
    '   Create a new brush
    NewBrush& = CreateSolidBrush(Color&)
    '   Fill the Rectangle
    Call FillRect(hdc&, RC, NewBrush&)
    '   Clean up...
    Call DeleteObject(NewBrush&)
    Exit Sub
    
APIFillRect_Error:
End Sub

Private Sub APIFillRectByCoords(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Color As Long)
    On Error GoTo APIFillRectByCoords_Error
    
    '   Fill a Rectangular region using its hDC and the color
    '   passed by the caller.
    
    Dim NewBrush As Long
    Dim tmpRect As RECT
    '   Create a new brush
    NewBrush& = CreateSolidBrush(Color&)
    '   Setr the active rectangle
    SetRect tmpRect, X, Y, X + W, Y + H
    '   Fill the Rectangle
    Call FillRect(hdc&, tmpRect, NewBrush&)
    '   Clean up
    Call DeleteObject(NewBrush&)
    Exit Sub
    
APIFillRectByCoords_Error:
End Sub

Private Sub APILine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, lColor As Long)
    On Error GoTo APILine_Error
    
    '   Use the API LineTo for Fast Drawing based on the
    '   Usercontrols hDC
    
    Dim pt As POINT
    Dim hPen As Long, hPenOld As Long
    '   Create a new pen
    hPen = CreatePen(0, 1, lColor)
    '   Set the new pen and save the old
    hPenOld = SelectObject(UserControl.hdc, hPen)
    '   Define the segment start
    MoveToEx UserControl.hdc, X1, Y1, pt
    '   Draw the segment
    LineTo UserControl.hdc, X2, Y2
    '   Set the pen back...
    SelectObject UserControl.hdc, hPenOld
    '   Clean up
    DeleteObject hPen
    Exit Sub
    
APILine_Error:
End Sub

Private Sub APILineEx(ByVal lhDCEx As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal lColor As Long)
    On Error GoTo APILineEx_Error
    
    '   Use the API LineTo for Fast Drawing based on the
    '   hDC specified by the calling routine
    
    Dim pt As POINT
    Dim hPen As Long, hPenOld As Long
    '   Create a new pen
    hPen = CreatePen(0, 1, lColor)
    '   Set the new pen and save the old
    hPenOld = SelectObject(lhDCEx, hPen)
    '   Define the segment start
    MoveToEx lhDCEx, X1, Y1, pt
    '   Draw the segment
    LineTo lhDCEx, X2, Y2
    '   Set the pen back...
    SelectObject lhDCEx, hPenOld
    '   Clean up
    DeleteObject hPen
    Exit Sub
    
APILineEx_Error:
End Sub

Private Function APIRectangle(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal W As Long, ByVal H As Long, Optional lColor As OLE_COLOR = -1) As Long
    On Error GoTo APIRectangle_Error
    
    '   Draw a Rectangle using API Line drawing to the
    '   hDC as specified by the caller. One could substitute
    '   the API rouine Rectangle to achive the same result.
    
    Dim hPen As Long, hPenOld As Long
    Dim pt As POINT
    '   Create a new pen
    hPen = CreatePen(0, 1, lColor)
    '   Set the new pen and save the old
    hPenOld = SelectObject(hdc, hPen)
    '   Define the segment start
    MoveToEx hdc, X, Y, pt
    '   Draw the segment(s)
    LineTo hdc, X + W, Y
    LineTo hdc, X + W, Y + H
    LineTo hdc, X, Y + H
    LineTo hdc, X, Y
    '   Set the pen back..
    SelectObject hdc, hPenOld
    '   Clean up
    DeleteObject hPen
    Exit Function
    
APIRectangle_Error:
End Function

Public Property Get BackColor() As OLE_COLOR
    On Error GoTo BackColor_Error
    
    '   Get the backcolor of the control
    
    BackColor = m_BackColor
    Exit Property
    
BackColor_Error:
End Property

'Description: Use this color for drawing
Public Property Let BackColor(ByVal lBackColor As OLE_COLOR)
    On Error GoTo BackColor_Error
    
    '   Set the backcolor of the control
    
    m_BackColor = lBackColor
    PropertyChanged "BackColor"
    Refresh
    Exit Property
    
BackColor_Error:
End Property

'Blend two colors
Private Function BlendColors(ByVal lColor1 As Long, ByVal lColor2 As Long)
    On Error GoTo BlendColors_Error
    
    '   This routine blends two colors to achieve a new color.
    
    BlendColors = RGB(((lColor1 And &HFF) + (lColor2 And &HFF)) / 2, (((lColor1 \ &H100) And &HFF) + ((lColor2 \ &H100) And &HFF)) / 2, (((lColor1 \ &H10000) And &HFF) + ((lColor2 \ &H10000) And &HFF)) / 2)
    Exit Function
    
BlendColors_Error:
End Function

Private Sub BuildRegion()
    Dim pPoligon(8)     As POINT
    Dim pEllipse1(1)    As POINT
    Dim pEllipse2(1)    As POINT
    Dim lTmpRgn         As Long
    Dim lw              As Long
    Dim lh              As Long
    Dim hBrush          As Long
    Dim lRtn            As Long
    Dim lYSz           As Long

    On Error GoTo BuildRegion_Error
    
    '   This routine computes the regions for the double ellipse
    '   and the basic Label shape. The trick here is to use Regions to
    '   allow for hit testing in the various regions when drawing and
    '   to permit clipping of the drawing surface when displaying the
    '   double ellipse gradients.
    
    '   Clean up any previous regions to begin
    If m_lRegion Then DeleteObject m_lRegion
    If m_lButtonRgn Then DeleteObject m_lButtonRgn
    If lTmpRgn Then DeleteObject lTmpRgn
    
    '   Clear the usercontrol to start
    UserControl.Cls
    
    '   Set the size of the control
    lYSz = m_ButtonBackSize
    '   Get the scale width and height
    lw = UserControl.ScaleWidth
    lh = UserControl.ScaleHeight
    
    '   Put the Button and Ellipses on Left or Right side of the label?
    If m_ButtonAlign = clbLeft Then
        '   Define the coordinates of the small region
        pEllipse1(0).X = 12
        pEllipse1(0).Y = ((lh / 2) - ((lYSz / 1.5) / 2))
        pEllipse1(1).X = ((lYSz / 1.5) + pEllipse1(0).X)
        pEllipse1(1).Y = ((lYSz / 1.5) + pEllipse1(0).Y)
                
        '   Now set the type of region to be used....
        Select Case m_ButtonShape
            Case clbEllipse     'Round
                '   Create a small elliptical region
                lTmpRgn = CreateEllipticRgn(pEllipse1(0).X, pEllipse1(0).Y, pEllipse1(1).X, pEllipse1(1).Y)
            Case clbRndRect     'Rounded Rectangle
                '   Create a larger RoundRect region
                lTmpRgn = CreateRoundRectRgn(pEllipse1(0).X, pEllipse1(0).Y, pEllipse1(1).X, pEllipse1(1).Y, m_CornerSize, m_CornerSize)
            Case clbRectangle   'Rectangle
                '   Create a larger Rectangular region
                lTmpRgn = CreateRoundRectRgn(pEllipse1(0).X, pEllipse1(0).Y, pEllipse1(1).X, pEllipse1(1).Y, 0, 0)
        End Select
        
        '   Define the coordinates of the large region
        pEllipse2(0).X = 19
        pEllipse2(0).Y = ((lh / 2) - (lYSz) / 2)
        pEllipse2(1).X = ((lYSz) + pEllipse2(0).X)
        pEllipse2(1).Y = ((lYSz) + pEllipse2(0).Y)
            
        '   Now set the type of region to be used....
        Select Case m_ButtonShape
            Case clbEllipse     'Round
                '   Create a larger elliptical region
                m_lButtonRgn = CreateEllipticRgn(pEllipse2(0).X, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y)
            Case clbRndRect     'Rounded Rectangle
                '   Create a larger RoundRect region
                m_lButtonRgn = CreateRoundRectRgn(pEllipse2(0).X, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y, m_CornerSize, m_CornerSize)
            Case clbRectangle   'Rectangle
                '   Create a larger Rectangular region
                m_lButtonRgn = CreateRoundRectRgn(pEllipse2(0).X, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y, 0, 0)
        End Select
    Else
        '   Define the coordinates of the small region
        pEllipse1(0).X = (lw - lYSz - 19)
        pEllipse1(0).Y = ((lh / 2) - ((lYSz / 1.5) / 2))
        pEllipse1(1).X = ((lYSz / 1.5) + pEllipse1(0).X)
        pEllipse1(1).Y = ((lYSz / 1.5) + pEllipse1(0).Y)
        
        '   Now set the type of region to be used....
        Select Case m_ButtonShape
            Case clbEllipse     'Round
                '   Create a small elliptical region
                lTmpRgn = CreateEllipticRgn(pEllipse1(0).X, pEllipse1(0).Y, pEllipse1(1).X, pEllipse1(1).Y)
            Case clbRndRect     'Rectangle
                '   Create a larger RoundRect region
                lTmpRgn = CreateRoundRectRgn(pEllipse1(0).X, pEllipse1(0).Y, pEllipse1(1).X, pEllipse1(1).Y, m_CornerSize, m_CornerSize)
            Case clbRectangle   'Rounded Rectangle
                '   Create a larger Rectangular region
                lTmpRgn = CreateRoundRectRgn(pEllipse1(0).X, pEllipse1(0).Y, pEllipse1(1).X, pEllipse1(1).Y, 0, 0)
        End Select

        '   Define the coordinates of the large region
        pEllipse2(0).X = (lw - lYSz - 12)
        pEllipse2(0).Y = ((lh / 2) - (lYSz) / 2)
        pEllipse2(1).X = ((lYSz) + pEllipse2(0).X)
        pEllipse2(1).Y = ((lYSz) + pEllipse2(0).Y)
        
        '   Now set the type of region to be used....
        Select Case m_ButtonShape
            Case clbEllipse     'Round
                '   Create a larger elliptical region
                m_lButtonRgn = CreateEllipticRgn(pEllipse2(0).X, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y)
            Case clbRndRect     'Rectangle
                '   Create a larger RoundRect region
                m_lButtonRgn = CreateRoundRectRgn(pEllipse2(0).X, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y, m_CornerSize, m_CornerSize)
            Case clbRectangle   'Rounded Rectangle
                '   Create a larger Rectangular region
                m_lButtonRgn = CreateRoundRectRgn(pEllipse2(0).X, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y, 0, 0)
        End Select
    End If

    If m_ButtonShape <> clbRectangle Then
        '   Combine the regions into one...
        CombineRgn m_lButtonRgn, lTmpRgn, m_lButtonRgn, RGN_OR
        '   Get the Bounding Box for the the Elliptical regions
        SetRect m_btnRect, pEllipse1(0).X, pEllipse2(0).Y, pEllipse2(1).X, (pEllipse2(1).Y - pEllipse2(0).Y)
    Else
        '   Get the Bounding Box for the the Elliptical regions
        SetRect m_btnRect, pEllipse1(0).X + (pEllipse2(0).X - pEllipse1(0).X) - 6, pEllipse2(0).Y, pEllipse2(1).X, pEllipse2(1).Y
    End If
        
    '   Draw a line around region
    hBrush = CreateSolidBrush(&HC0C0C0)
    lRtn = FrameRgn(UserControl.hdc, m_lButtonRgn, hBrush, 1, 1)
    lRtn = DeleteObject(hBrush)
    
    '   Define the coordinates for the body of the
    '   label as a series of points along a polygon
    pPoligon(0).X = 0: pPoligon(0).Y = 2
    pPoligon(1).X = 2: pPoligon(1).Y = 0
    pPoligon(2).X = lw - 3: pPoligon(2).Y = 0
    pPoligon(3).X = lw: pPoligon(3).Y = 3
    pPoligon(4).X = lw: pPoligon(4).Y = lh - 3
    pPoligon(5).X = lw - 3: pPoligon(5).Y = lh
    pPoligon(6).X = 4: pPoligon(6).Y = lh
    pPoligon(7).X = 0: pPoligon(7).Y = lh - 4
    m_lRegion = CreatePolygonRgn(pPoligon(0), 8, ALTERNATE)
    
    '   One could use the following as well, but with
    '   less control over the shape of the corners...
    'm_lRegion = CreateRoundRectRgn(0, 0, lw, lh, 15, 15)
    
    '   Draw a line around region
    hBrush = CreateSolidBrush(&HA5A3A2)
    lRtn = FrameRgn(UserControl.hdc, m_lRegion, hBrush, 1, 1)
    lRtn = DeleteObject(hBrush)

    '   Combine the regions into one...
    CombineRgn m_lRegion, m_lButtonRgn, m_lRegion, RGN_OR
        
    '   Set the active window region
    SetWindowRgn UserControl.hwnd, m_lRegion, True
    
    '   Delete the temporary regions
    DeleteObject lTmpRgn

BuildRegion_Error:
End Sub

Public Property Get ButtonAlign() As clbAlign
    On Error GoTo ButtonAlign_Error
    
    '   Where did the button go...left, right?
    ButtonAlign = m_ButtonAlign
    Exit Property
    
ButtonAlign_Error:
End Property

Public Property Let ButtonAlign(ByVal NewButtonAlign As clbAlign)
' Description: this is the "ButtonAlign" property.
    On Error GoTo ButtonAlign_Error
    
    '   Where does the button go...left, right?
    m_ButtonAlign = NewButtonAlign
    PropertyChanged "ButtonAlign"
    UserControl_Resize
    UserControl_Paint
    Exit Property
    
ButtonAlign_Error:
End Property

Public Property Get ButtonShape() As clbShape
    On Error GoTo ButtonShape_Error
    
    '   What shape are we using for the backdrop?
    ButtonShape = m_ButtonShape
    Exit Property
    
ButtonShape_Error:
End Property

Public Property Let ButtonShape(ByVal NewButtonShape As clbShape)
'   Description: this is the "ButtonAlign" property.
    On Error GoTo ButtonShape_Error
    
    '   Set what shape are we using for the backdrop...
    m_ButtonShape = NewButtonShape
    PropertyChanged "ButtonShape"
    UserControl_Resize
    UserControl_Paint
    Exit Property
    
ButtonShape_Error:
End Property

Public Property Get ButtonToolTipText() As String
    On Error GoTo ButtonToolTipText_Error
    
    '   What was the ButtonToolTipText?
    ButtonToolTipText = m_ButtonToolTipText
    Exit Property
    
ButtonToolTipText_Error:
End Property

Public Property Let ButtonToolTipText(ByVal NewToolTipText As String)
'   Description: this is the "ButtonAlign" property.
    On Error GoTo ButtonToolTipText_Error
    
    '   Add the tooltiptext to the buttons?
    m_ButtonToolTipText = NewToolTipText
    '   Add this to all Image Controls.....
    UserControl.imgCustomPic.ToolTipText = NewToolTipText
    UserControl.imgFailed.ToolTipText = NewToolTipText
    UserControl.imgNext.ToolTipText = NewToolTipText
    UserControl.imgSuccess.ToolTipText = NewToolTipText
    PropertyChanged "ButtonToolTipText"
    UserControl_Resize
    UserControl_Paint
    Exit Property
    
ButtonToolTipText_Error:
End Property

Public Property Get CaptionAlign() As clAlign
'   Description: this is the "CaptionAlign" property.
    On Error GoTo CaptionAlign_Error
    
    '   Get the caption alignment...left, top, right, bottom, center
    CaptionAlign = m_CaptionAlign
    Exit Property
    
CaptionAlign_Error:
End Property

Public Property Let CaptionAlign(ByVal NewCaptionAlign As clAlign)
'   Description: this is the "CaptionAlign" property.
    On Error GoTo CaptionAlign_Error
    
    '   Set the caption alignment...left, top, right, bottom, center
    m_CaptionAlign = NewCaptionAlign
    PropertyChanged "CaptionAlign"
    UserControl_Resize
    UserControl_Paint
    Exit Property
    
CaptionAlign_Error:
End Property

Public Property Get Caption() As String
    On Error GoTo Caption_Error
    
    '   What does the caption say?
    Caption = m_Caption
    Exit Property
    
Caption_Error:
End Property


Public Property Let Caption(ByVal NewCaption As String)
'   Description: this is the "Caption" property.
    On Error GoTo Caption_Error
    
    '   Set the caption text....
    m_Caption = NewCaption
    PropertyChanged "Caption"
    UserControl_Resize
    Refresh
    Exit Property
    
Caption_Error:
End Property

Public Property Get CornerSize() As Long
    On Error GoTo CornerSize_Error
    
    '   What size is the corner when the button is
    '   set to clbRndRect (Rounded Rectangle), not used otherwise
    CornerSize = m_CornerSize
    Exit Property
    
CornerSize_Error:
End Property

Public Property Let CornerSize(ByVal NewCornerSize As Long)
'   Description: this is the "CornerSize" property.
    On Error GoTo CornerSize_Error
    
    '   Set the size of the radius for the corner when the button is
    '   set to clbRndRect (Rounded Rectangle)
    m_CornerSize = NewCornerSize
    PropertyChanged "CornerSize"
    UserControl_Resize
    UserControl_Paint
    Exit Property
    
CornerSize_Error:
End Property

Public Property Get CustomColor() As OLE_COLOR
    On Error GoTo CustomColor_Error
    
    '   Get the Highlighted (clHot, "Hot") Color for
    '   Custom Colors (i.e. when UseCustomColor = True)
    CustomColor = m_CustomColor
    Exit Property
    
CustomColor_Error:
End Property

Public Property Let CustomColor(ByVal lCustomColor As OLE_COLOR)
    On Error GoTo CustomColor_Error
    
    '   Set the Gradient Color (clHot, "Hot") for when
    '   Custom Colors are active (i.e. when UseCustomColor = True)
    m_CustomColor = lCustomColor
    PropertyChanged "CustomColor"
    Refresh
    Exit Property
    
CustomColor_Error:
End Property

Private Sub DrawAlphaRegion(ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    
    '   This routine simulates the alpha blending used to minimize the
    '   the aliasing seen in API drawing routines. When an arc or circle is
    '   drawn with vastly different colors the net effect is a jagged edge caused
    '   by an aliasing of the original pixels on to a rectalinear grid.
    '   To minimize this, one can blur the image along the region interfaces
    '   by computing a 9 element square and averaging the color along the
    '   boundry. This results in an apparent smooth transition with out the
    '   need for the AlphaBlend API. Since AlphaBlend is only available in
    '   WinXP and 2K machines, this would limit the utlity of this control...
    '   Thus this routine simulates this effect by computing the local average
    '   and replacing the pixels accordingly...
    
    On Error GoTo DrawAlphaRegion_Error
    
    Dim eX1 As Long, eX2 As Long
    Dim ni As Long, i As Long, j As Long
    Dim lRet As Long
    Dim lLng As Long, k As Long
    
    For ni = 0 To (Y2 - 2)
        '   Make sure the line is within the region boundries on the left (X----)
        If Not PtInRegion(m_lButtonRgn, X, Y + ni) Then
            For i = X To X2
                If PtInRegion(m_lButtonRgn, i, Y + ni) Then
                    eX1 = i - 1
                    Exit For
                End If
            Next i
        Else
            eX1 = X
        End If
        k = 1
        '   Now set the Averaged Pixel
        For i = -1 To 1
            For j = -1 To 1
                '   Sum up the pixels to average the color
                If GetPixel(UserControl.hdc, eX1 + j, Y + ni + i) Then
                    lLng = lLng + GetPixel(UserControl.hdc, eX1 + j, Y + ni + i)
                    k = k + 1
                End If
            Next j
        Next i
        If lLng > 0 Then
            '   Now set the average color
            lRet = SetPixelV(UserControl.hdc, eX1 + 0, Y + ni, (lLng / k))
        End If
        '   Make sure the line is within the region boundries on the right (----X)
        If Not PtInRegion(m_lButtonRgn, X2, Y + ni) Then
            For i = X2 To X Step -1
                If PtInRegion(m_lButtonRgn, i, Y + ni) Then
                    eX2 = i + 1
                    Exit For
                End If
            Next i
        Else
            eX2 = X2
        End If
        k = 1
        '   Now set the Averaged Pixel
        For i = -1 To 1
            For j = -1 To 1
                '   Sum up the pixels to average the color
                If GetPixel(UserControl.hdc, eX2 + j, Y + ni + i) Then
                    lLng = lLng + GetPixel(UserControl.hdc, eX2 + j, Y + ni + i)
                    k = k + 1
                End If
            Next j
        Next i
        If lLng > 0 Then
            '   Now set the average color
            lRet = SetPixelV(UserControl.hdc, eX2 + 0, Y + ni, (lLng / k))
        End If
    Next 'ni
    
    Exit Sub

DrawAlphaRegion_Error:
End Sub

Private Sub DrawCaption()
    On Error GoTo DrawCaption_Error
    
    '   Draw the caption text on the surface of the control.
    '   The trick to alignment is the flags set in the UserControl_Resize
    '   event handler....along with settting the correct text rectangle...
    
    Dim lColor As Long, lTmpColor As Long
    
    If m_UseCustomColors Then
        If m_iState <> clDisabled Then
            lColor = GetSysColor(COLOR_BTNTEXT)
        Else
            lColor = TranslateColor(vbGrayText)
        End If

    Else
        Select Case m_iState
            Case clNormal
                lColor = m_FontColor
            Case clDisabled
                lColor = TranslateColor(vbGrayText)
            Case Else
                lColor = m_FontHighlightColor
        End Select

    End If
    SelectClipRgn UserControl.hdc, m_lRegion
    lTmpColor = UserControl.ForeColor
    UserControl.ForeColor = lColor
    DrawText UserControl.hdc, m_Caption, -1, m_txtRect, m_lwFontAlign
    UserControl.ForeColor = lTmpColor
    Exit Sub
    
DrawCaption_Error:
End Sub

Private Sub DrawHGradient(lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    On Error GoTo DrawHGradient_Error
    
    '   Draw a Horizontal Gradient in the current hDC
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long
    lh = Y2 - Y
    lw = X2 - X
    '   Get the Starting R,G,B color components
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    '   Get the Ending R,G,B color components
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    '   Compute the Delta and divide by the number of steps (lw)
    dR = (sR - eR) / lw
    dG = (sG - eG) / lw
    dB = (sB - eB) / lw
    
    For ni = 0 To lw
        APILine X + ni, Y, X + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawHGradient_Error:
End Sub

Private Sub DrawRectGradient(lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, Optional X3 As Long, Optional Y3 As Long)
    On Error GoTo DrawRectGradient_Error
    
    '   Draw a Rectangular Gradient in the current hDC
    '
    '   This routine will generate rectangular gradients which give
    '   the illusion of depth by starting from a color and RECT and
    '   decreasing the RECT and color in a stepwise fashion. If the
    '   Colors go from light to dark then illusion is a rectangle which
    '   appears to get deeper at the center. This drawing method can be
    '   expanded to include Ellipses and Arcs which give nonlinear gradients
    '   if this effect is desired....for more details contact the author
    '   Paul R. Territo, Ph.D at the above e-mail address.
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim lh As Long, lw As Long
    Dim ni As Long, lColor As Long, lRet As Long
    Dim hPen As Long, hPenOld As Long, hBrsh As Long

    lh = Y2 - Y
    lw = X2 - X
    '   Get the Starting R,G,B color components
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    '   Get the Ending R,G,B color components
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    '   Compute the Delta and divide by the number of steps (lw)
    dR = (sR - eR) / lw
    dG = (sG - eG) / lw
    dB = (sB - eB) / lw

    '   Paint the background of the region
    hBrsh = CreateSolidBrush(m_BackColor)
    lRet = FillRect(UserControl.hdc, m_btnRect, hBrsh)
    
    '   Now fill the gradient rectangles
    For ni = 0 To lw / 3
        lColor = RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        hPen = CreatePen(0, 1, lColor)
        hPenOld = SelectObject(UserControl.hdc, hPen)
        '   We could use either Rectangle or RoudnRect, but I choose RoudnRect
        '   with a X3, Y3 of 0, 0, because we already have this declared in the
        '   and there was no real difference in performace....
        lRet = RoundRect(UserControl.hdc, (X + ni), (Y + ni), (X2 - ni), (Y2 - ni), X3, Y3)
        SelectObject UserControl.hdc, hPenOld
        DeleteObject hPen
    Next 'ni
    
    
DrawRectGradient_Error:
End Sub

Private Sub DrawVGradient(lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    On Error GoTo DrawVGradient_Error
    
    '   Draw a Vertical Gradient in the current hDC
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    '   Get the Starting R,G,B color components
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    '   Get the Ending R,G,B color components
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    '   Compute the Delta and divide by the number of steps (Y2)
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    For ni = 0 To Y2
        APILine X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawVGradient_Error:
End Sub

Private Sub DrawVGradientEx(lhDCEx As Long, lEndColor As Long, lStartColor As Long, ByVal X As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    On Error GoTo DrawVGradientEx_Error
    
    '   Draw a Vertical Gradient in the hDC passed
    '   by the Caller...
    
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    '   Get the Starting R,G,B color components
    sR = (lStartColor And &HFF)
    sG = (lStartColor \ &H100) And &HFF
    sB = (lStartColor And &HFF0000) / &H10000
    '   Get the Ending R,G,B color components
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    '   Compute the Delta and divide by the number of steps (Y2)
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    For ni = 0 To Y2
        APILineEx lhDCEx, X, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next 'ni
    
    Exit Sub
    
DrawVGradientEx_Error:
End Sub

Public Property Get FontColor() As OLE_COLOR
    On Error GoTo FontColor_Error
    
    '   Which font color is it?
    
    FontColor = m_FontColor
    Exit Property
    
FontColor_Error:
End Property

'Description: Use this color for drawing normal font
Public Property Let FontColor(ByVal lFontColor As OLE_COLOR)
    On Error GoTo FontColor_Error
    
    '   Set the font color...
    m_FontColor = lFontColor
    PropertyChanged "FontColor"
    Refresh
    Exit Property
    
FontColor_Error:
End Property

Public Property Get Font() As StdFont
    On Error GoTo Font_Error
    
    '   Which font is it?
    Set Font = UserControl.Font
    Exit Property
    
Font_Error:
End Property

Public Property Get FontHighlightColor() As OLE_COLOR
    On Error GoTo FontHighlightColor_Error
    
    '   Get the Highlighted (clHot, "Hot") Color of the Font
    FontHighlightColor = m_FontHighlightColor
    Exit Property
    
FontHighlightColor_Error:
End Property

Public Property Get hdc() As Long
    '   Get the hDC for the control
    hdc = UserControl.hdc
End Property

Public Property Get hwnd() As Long
    '   Get the hWnd for the control
    hwnd = UserControl.hwnd
End Property

Public Property Let FontHighlightColor(ByVal lFontHighlightColor As OLE_COLOR)
    On Error GoTo FontHighlightColor_Error
    
    '   Set the Highlighted (clHot, "Hot") Color of the Font
    m_FontHighlightColor = lFontHighlightColor
    PropertyChanged "FontHighlightColor"
    Refresh
    Exit Property
    
FontHighlightColor_Error:
End Property

Public Property Set Font(NewFont As StdFont)
    On Error GoTo Font_Error
    
    '   Set a new font....
    Set m_Font = NewFont
    Set UserControl.Font = NewFont
    Refresh
    PropertyChanged "Font"
    Exit Property
    
Font_Error:
End Property

Public Property Get Icon() As StdPicture
    On Error GoTo Icon_Error
    
    '   Get the current picture from the container
    If UserControl.imgCustomPic.Picture <> 0 Then
        Set Icon = UserControl.imgCustomPic.Picture
    End If
    Exit Property
    
Icon_Error:
End Property

Public Property Set Icon(ByVal NewPicture As StdPicture)
    '   Description: this is the "Icon" property.
    On Error GoTo Icon_Error
    
    '   Set the new cusom picture into the control....
    Set UserControl.imgCustomPic.Picture = NewPicture
    PropertyChanged "Icon"
    UserControl_Resize
    Refresh
    Exit Property
    
Icon_Error:
End Property

Public Property Get ButtonIcon() As ImageType
    On Error GoTo ButtonIcon_Error
    
    '   Get the ButtonIcon Used
    ButtonIcon = m_ButtonIcon
    
ButtonIcon_Error:
End Property

Public Property Let ButtonIcon(ByVal lButtonIcon As ImageType)
    On Error GoTo ButtonIcon_Error
    
    '   Set the ButtonIcon to use
    m_ButtonIcon = lButtonIcon
'    BuildRegion
    If m_ButtonIcon = clbCustom Then
        UseCustomIcon = True
    Else
        UseCustomIcon = False
        m_ButtonIcon = lButtonIcon
    End If
    Refresh
    PropertyChanged "ButtonIcon"
    
ButtonIcon_Error:
End Property

Public Property Get ButtonBackSize() As Long
    On Error GoTo ButtonBackSize_Error
    
    '   This property is helpful if you wan to set the Button
    '   Background size to a specific size...i.e. Icon = 64x64, then
    '   the ButtonBackSize = 66 or 68 to make it slightly larger than
    '   the icon....
    ButtonBackSize = m_ButtonBackSize
    
ButtonBackSize_Error:
End Property

Public Property Let ButtonBackSize(ByVal lButtonBackSize As Long)
    On Error GoTo ButtonBackSize_Error
    
    '   Prevent the Image from being smaller than icons
    m_ButtonBackSize = IIf((lButtonBackSize < 33) And (m_ButtonIcon <> clbCustom), 33, lButtonBackSize)
    BuildRegion
    Refresh
    PropertyChanged "ButtonBackSize"
    
ButtonBackSize_Error:
End Property

Public Property Get Enabled() As Boolean
    On Error GoTo Enabled_Error
    
    Enabled = m_Enabled
    
Enabled_Error:
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)
    On Error GoTo Enabled_Error
    
    '   Enable/Disable the control
    m_Enabled = bEnabled
    BuildRegion
    Refresh
    PropertyChanged "Enabled"
    
Enabled_Error:
End Property

Private Sub imgCustomPic_Click()
    If m_Enabled Then
        RaiseEvent ButtonClick
    End If
End Sub

Private Sub imgCustomPic_DblClick()
    If m_Enabled Then
        RaiseEvent ButtonDblClick
    End If
End Sub

Private Sub imgCustomPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgCustomPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgCustomPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgFailed_Click()
    If m_Enabled Then
        RaiseEvent ButtonClick
    End If
End Sub

Private Sub imgFailed_DblClick()
    If m_Enabled Then
        RaiseEvent ButtonDblClick
    End If
End Sub

Private Sub imgFailed_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgFailed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgFailed_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgNext_Click()
    If m_Enabled Then
        RaiseEvent ButtonClick
    End If
End Sub

Private Sub imgNext_DblClick()
    If m_Enabled Then
        RaiseEvent ButtonDblClick
    End If
End Sub

Private Sub imgNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgSuccess_Click()
    If m_Enabled Then
        RaiseEvent ButtonClick
    End If
End Sub

Private Sub imgSuccess_DblClick()
    If m_Enabled Then
        RaiseEvent ButtonDblClick
    End If
End Sub

Private Sub imgSuccess_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgSuccess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub imgSuccess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent ButtonMouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Function OffsetColor(ByVal lColor As OLE_COLOR, ByVal lOffset As Long) As OLE_COLOR
    On Error GoTo OffsetColor_Error
    
    '   This routine computes an offset color from the color based on the
    '   offset value passed by the caller.
    Dim lRed As OLE_COLOR
    Dim lGreen As OLE_COLOR
    Dim lBlue As OLE_COLOR
    Dim lr As OLE_COLOR, lg As OLE_COLOR, lb As OLE_COLOR
    
    '   Make sure to translate the colors to allow
    '   System Color constants to be used....see
    '   http://www.vb-helper.com/tut10.htm for more details
    lColor = TranslateColor(lColor)
    lr = (lColor And &HFF)
    lg = ((lColor And 65280) \ 256)
    lb = ((lColor) And 16711680) \ 65536
    lRed = (lOffset + lr)
    lGreen = (lOffset + lg)
    lBlue = (lOffset + lb)
    
    If lRed > 255 Then lRed = 255
    If lRed < 0 Then lRed = 0
    If lGreen > 255 Then lGreen = 255
    If lGreen < 0 Then lGreen = 0
    If lBlue > 255 Then lBlue = 255
    If lBlue < 0 Then lBlue = 0
    
    OffsetColor = RGB(lRed, lGreen, lBlue)
    
    Exit Function
    
OffsetColor_Error:
End Function

Public Sub Refresh()
    Dim lStartColor As Long
    Dim lEndColor As Long
    Dim lTmpColor As Long
    Dim lRectOffset As Long
    
    '   This routine is the main routine for drawing and
    '   displaying the visual effects seen in the label. The
    '   color gradients, and color offsets are all computed here...
    '   Care should be taken not to change the order of the routines
    '   as this will result in undesirable drawings which may
    '   paint over the previously drawn sections.... ;-)
        
    '   Clear the control
    UserControl.Cls
    
    'If Not m_bVisible Then Exit Sub
    If Not UserControl.Ambient.UserMode Then
        m_iState = clHot
    End If
    If Not m_Enabled Then
        m_iState = clDisabled
        UserControl.BackColor = GetSysColor(COLOR_BTNFACE)
    Else
        UserControl.BackColor = IIf(m_iState = clNormal, &HDDDDDD, OffsetColor(&HDDDDDD, &HD))
    End If
    
    '   Select the Region to draw in and only this region...
    SelectClipRgn UserControl.hdc, m_lRegion
    
    '   Fill the background with a VGradient
    If UseCustomColor = False Then
        '   Used for custom colors
        lStartColor = IIf(m_iState = clNormal, &HFFFFFF, OffsetColor(&HFFFFFF, &HF))
        lEndColor = IIf(m_iState = clNormal, &HD0D0D0, OffsetColor(&HD0D0D0, &HF))
    Else
        '   Used with standard grey color
        lStartColor = IIf(m_iState = clNormal, &HFFFFFF, OffsetColor(&HFFFFFF, &HF))
        lEndColor = IIf(m_iState = clNormal, OffsetColor(m_CustomColor, -&HF), OffsetColor(m_CustomColor, &H2))
    End If
    
    If m_UseGradient = False Then
        '   If no gradient then set the start = end
        lStartColor = lEndColor
    End If
    DrawVGradient lStartColor, lEndColor, 2, 2, UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 5

    '   Draw the caption
    DrawCaption
    
    '   Select the Region to draw in and only this region...
    SelectClipRgn UserControl.hdc, m_lButtonRgn
    
    '   Draw the Full Gradiant in the Button Area
    If UseCustomColor = False Then
        '   Used with standard colors
        lStartColor = IIf(m_iState = clNormal, &HF0F0F0, OffsetColor(&HF0F0F0, &HF))
        lEndColor = IIf(m_iState = clNormal, &HFFFFFF, OffsetColor(&HFFFFFF, &HF))
    Else
        '   Used for custom colors
        lTmpColor = OffsetColor(m_CustomColor, -&H20)
        lStartColor = IIf(m_iState = clNormal, lTmpColor, OffsetColor(lTmpColor, &HF))
        lEndColor = IIf(m_iState = clNormal, &HFFFFFF, OffsetColor(&HFFFFFF, &H2)) '&HFA))
    End If
    '   See if we want it Monochrome
    If m_UseGradient = False Then
        '   If no gradient then set the start = end
        lEndColor = lStartColor
    End If
    If (m_ButtonShape <> clbRectangle) Then
        '   Ellipse, and RndRectangle shapes
        DrawVGradient lStartColor, lEndColor, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom
    Else
        '   Adjust the color brighter to give the rectangles
        '   the appearance of depth as it goes from outer edge to center
        If UseCustomColor = False Then
            lTmpColor = OffsetColor(&HC0C0C0, -&HF0)
            lStartColor = IIf(m_iState = clNormal, &HFFFFFF, OffsetColor(&HFFFFFF, &HF))
            lEndColor = IIf(m_iState = clNormal, lTmpColor, OffsetColor(lTmpColor, &HF))
        Else
            lTmpColor = BlendColors(&HFFFFFF, m_CustomColor)
            lStartColor = IIf(m_iState = clNormal, lTmpColor, OffsetColor(lTmpColor, &H2)) '&HF))
            lTmpColor = OffsetColor(m_CustomColor, -&HC8)
            lEndColor = IIf(m_iState = clNormal, lTmpColor, OffsetColor(lTmpColor, &H2)) '&HF))
        End If
        DrawRectGradient lStartColor, lEndColor, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom
    End If
    '   Now the top half...in the Button Area
    If (UseCustomColor = False) And (m_ButtonShape <> clbRectangle) Then
        lStartColor = IIf(m_iState = clNormal, &HA0A0A0, OffsetColor(&HA0A0A0, &HF))
        lEndColor = IIf(m_iState = clNormal, &HF0F0F0, OffsetColor(&HF0F0F0, &HF))
        '   See if we want it Monochrome
        If m_UseGradient = False Then
            '   Make it monochrome
            lEndColor = lStartColor
            '   Draw the whole region
            DrawVGradient lStartColor, lEndColor, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom
        Else
            '   Draw only the top 1/2
            DrawVGradient lStartColor, lEndColor, m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom / 2
        End If
    End If

    '   Blend the edges to remove alaising....
    DrawAlphaRegion m_btnRect.Left, m_btnRect.Top, m_btnRect.Right, m_btnRect.Bottom
    
    '   Set the position of the icon images
    With UserControl
        '   If Shape is Rectangle we need to adjust the location
'        lRectOffset = IIf((m_ButtonShape = clbRectangle) And (UseCustomIcon = True), 1, 0)
        lRectOffset = IIf(m_ButtonShape = clbRectangle, 1, 0)
        If m_ButtonAlign = clbLeft Then
            With .imgNext
                .Visible = False
                .Left = ((m_btnRect.Right - m_btnRect.Left) / 1.16) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
            With .imgSuccess
                .Visible = False
                .Left = ((m_btnRect.Right - m_btnRect.Left) / 1.16) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
            With .imgFailed
                .Visible = False
                .Left = ((m_btnRect.Right - m_btnRect.Left) / 1.16) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
            With .imgCustomPic
                .Visible = False
                .Left = ((m_btnRect.Right - m_btnRect.Left) / 1.16) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
        Else
            With .imgNext
                .Visible = False
                .Left = (m_btnRect.Left - (.Width / 2)) + ((m_btnRect.Right - m_btnRect.Left) / 1.05) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
            With .imgSuccess
                .Visible = False
                .Left = (m_btnRect.Left - (.Width / 2)) + ((m_btnRect.Right - m_btnRect.Left) / 1.05) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
            With .imgFailed
                .Visible = False
                .Left = (m_btnRect.Left - (.Width / 2)) + ((m_btnRect.Right - m_btnRect.Left) / 1.05) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
            With .imgCustomPic
                .Visible = False
                .Left = (m_btnRect.Left - (.Width / 2)) + ((m_btnRect.Right - m_btnRect.Left) / 1.05) - (.Width / 2) + lRectOffset
                .Top = (UserControl.ScaleHeight / 2) - ((.Height - 2) / 2)
            End With
        End If
        '   Set the correct icon to be visible
        Select Case m_ButtonIcon
            Case clbNext
                .imgNext.Visible = True
            Case clbSuccess
                .imgSuccess.Visible = True
            Case clbFailed
                .imgFailed.Visible = True
            Case clbCustom
                .imgCustomPic.Visible = True
        End Select
    End With

End Sub

Private Function TranslateColor(ByVal lColor As Long) As Long
    On Error GoTo TranslateColor_Error
    
    '   System Color code to long RGB
    If OleTranslateColor(lColor, 0, TranslateColor) Then
        TranslateColor = -1
    End If
    
    Exit Function
    
TranslateColor_Error:
End Function

Public Property Get UseCustomColor() As Boolean
    On Error GoTo UseCustomColor_Error
    
    '   See if we are using custom colors?
    UseCustomColor = m_UseCustomColor

UseCustomColor_Error:
End Property

Public Property Let UseCustomColor(ByVal bValue As Boolean)
    On Error GoTo UseCustomColor_Error
    
    '   Use custom colors?
    m_UseCustomColor = bValue
    PropertyChanged "UseCustomColor"
    UserControl_Resize
    Refresh
    
UseCustomColor_Error:
End Property

Public Property Get UseCustomIcon() As Boolean
    On Error GoTo UseCustomIcon_Error
    
    '   See if we are using a Custom Icon?
    UseCustomIcon = m_UseCustomIcon

UseCustomIcon_Error:
End Property

Public Property Let UseCustomIcon(ByVal bValue As Boolean)
    On Error GoTo UseCustomIcon_Error
    
    If bValue = True Then
        '   Store our previous selection for rollback....if needed
        m_PrevImage = m_ButtonIcon
        '   Now set the ButtonIcon to the Custom setting
        m_ButtonIcon = clbCustom
    Else
        If (m_ButtonIcon = m_PrevImage) And (m_PrevImage = clbCustom) Then
            m_ButtonIcon = clbNext
        Else
            m_ButtonIcon = m_PrevImage
        End If
    End If
    m_UseCustomIcon = bValue
    PropertyChanged "UseCustomIcon"
    UserControl_Resize
    Refresh
    
UseCustomIcon_Error:
End Property

Public Property Get UseGradient() As Boolean
    On Error GoTo UseGradient_Error
    
    '   See if we are using a gradient
    UseGradient = m_UseGradient

UseGradient_Error:
End Property

Public Property Let UseGradient(ByVal bValue As Boolean)
    On Error GoTo UseGradient_Error
    
    '   Set UseGradient flag
    m_UseGradient = bValue
    PropertyChanged "UseGradient"
    UserControl_Resize
    Refresh
    
UseGradient_Error:
End Property

Private Sub UserControl_Click()
    If m_Enabled Then
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_DblClick()
    If m_Enabled Then
        RaiseEvent DblClick
    End If
End Sub

Private Sub UserControl_InitProperties()
    '   Initial UserControl Settings
    m_BackColor = &HFFFFFF
    m_ButtonAlign = clbLeft
    m_ButtonBackSize = 34
    m_ButtonIcon = clbNext
    m_ButtonShape = clbEllipse
    m_Caption = UserControl.Extender.Name
    m_CaptionAlign = clCenter
    m_CornerSize = 12
    m_Enabled = True
    m_FontHighlightColor = &H800000
    m_CustomColor = &HFF8080
    m_UseCustomColor = False
    m_UseCustomIcon = False
    m_UseGradient = True
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Enabled Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_Paint()
    On Error GoTo UserControl_Paint_Error
    
    Call Refresh
    Exit Sub
    
UserControl_Paint_Error:
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    '   Read the properties from the property bag
    '   Also, a good place to start the subclassing (if we're running)
    m_iState = clNormal
    
    With PropBag
        m_BackColor = .ReadProperty("BackColor", &HFFFFFF)
        m_ButtonAlign = .ReadProperty("ButtonAlign", clbLeft)
        m_ButtonBackSize = .ReadProperty("ButtonBackSize", 34)
        m_ButtonIcon = .ReadProperty("ButtonIcon", clbNext)
        m_ButtonShape = .ReadProperty("ButtonShape", clbEllipse)
        m_ButtonToolTipText = .ReadProperty("ButtonToolTipText", vbNullString)
        UserControl.imgCustomPic.ToolTipText = .ReadProperty("ButtonToolTipText", vbNullString)
        UserControl.imgFailed.ToolTipText = .ReadProperty("ButtonToolTipText", vbNullString)
        UserControl.imgNext.ToolTipText = .ReadProperty("ButtonToolTipText", vbNullString)
        UserControl.imgSuccess.ToolTipText = .ReadProperty("ButtonToolTipText", vbNullString)
        m_Caption = .ReadProperty("Caption", UserControl.Extender.Name)
        m_CaptionAlign = .ReadProperty("CaptionAlign", m_CaptionAlign)
        m_CornerSize = .ReadProperty("CornerSize", 12)
        m_Enabled = .ReadProperty("Enabled", True)
        m_FontColor = .ReadProperty("FontColor", GetSysColor(COLOR_BTNTEXT))
        m_FontHighlightColor = .ReadProperty("FontHighlightColor", GetSysColor(COLOR_BTNTEXT))
        m_CustomColor = .ReadProperty("CustomColor", &HFF8080)
        m_UseCustomColor = .ReadProperty("UseCustomColor", False)
        m_UseCustomIcon = .ReadProperty("UseCustomIcon", False)
        m_UseGradient = .ReadProperty("UseGradient", True)
        Set m_Icon = .ReadProperty("Icon", Nothing)
        Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
        Set UserControl.imgCustomPic.Picture = .ReadProperty("Icon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", 0)
    End With
    
    If Ambient.UserMode Then                                                              'If we're not in design mode
        bTrack = True
        bTrackUser32 = IsFunctionExported("TrackMouseEvent", "User32")
        
        If Not bTrackUser32 Then
            If Not IsFunctionExported("_TrackMouseEvent", "Comctl32") Then
                bTrack = False
            End If
        End If
    
        If bTrack Then
            'OS supports mouse leave so subclass for it
            With UserControl
                'Start subclassing the UserControl
                Call Subclass_Start(.hwnd)
                Call Subclass_AddMsg(.hwnd, WM_MOUSEMOVE, MSG_AFTER)
                Call Subclass_AddMsg(.hwnd, WM_MOUSELEAVE, MSG_AFTER)
                
                'Start subclassing the Parent form
                With .Parent
                    Call Subclass_Start(.hwnd)
                    Call Subclass_AddMsg(.hwnd, WM_MOVING, MSG_AFTER)
                    Call Subclass_AddMsg(.hwnd, WM_SIZING, MSG_AFTER)
                    Call Subclass_AddMsg(.hwnd, WM_EXITSIZEMOVE, MSG_AFTER)
                End With
            End With
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    Dim tmpRect As RECT
    Dim lh As Long, lw As Long
    Dim lOffset As Long
    
    '   This is where all of the work starts....The main sections of
    '   the control are set here...(label and control areas). Also the
    '   text alignment sections are defined here to permit aligments
    '   that are specified in the property pages.
    
    On Error Resume Next
    
    '   Min size for the control.....this way we don't have only
    '   a button and no text or visa versa...
    If UserControl.Width < 1575 Then UserControl.Width = 1575
    If UserControl.Height < 615 Then UserControl.Height = 615
    
    lh = UserControl.ScaleHeight
    lw = UserControl.ScaleWidth
        
    If m_ButtonAlign = clbLeft Then
        '   The button is on the left, so we need to adjust the text areas
        '   so that they will fit....
        lOffset = 60
        SetRect m_ctlRect, 0, 0, lw, lh
        SetRect m_txtRect, lOffset, (lh / 2) - 8, lw - 4, lh - 4
        CopyRect tmpRect, m_txtRect
        DrawText UserControl.hdc, m_Caption, Len(m_Caption), tmpRect, DT_CALCRECT Or DT_WORDBREAK
        '   Setup the Caption alignment within the text area to match the property...
        Select Case m_CaptionAlign
            Case clCenter
                m_txtRect.Top = (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2
                m_lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_WORDBREAK
            Case clLeft
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2, lw - 8, (lh + tmpRect.Bottom - tmpRect.Top) / 2
                m_lwFontAlign = DT_VCENTER Or DT_LEFT Or DT_WORDBREAK
            Case clRight
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2, lw - 8, (lh + tmpRect.Bottom - tmpRect.Top) / 2
                m_lwFontAlign = DT_VCENTER Or DT_RIGHT Or DT_WORDBREAK
            Case clTop
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, 4, lw - 8, lh
                m_lwFontAlign = DT_CENTER Or DT_TOP Or DT_WORDBREAK
            Case clBottom
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, (lh - 4) - (tmpRect.Bottom - tmpRect.Top), lw - 8, lh - 4 'lOffset, (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2, lw - 8, (lh / 2) + (tmpRect.Bottom - tmpRect.Top)
                m_lwFontAlign = DT_CENTER Or DT_BOTTOM Or DT_WORDBREAK
        End Select
    Else
        '   The button is on the right, so we need to adjust the text areas
        '   so that they will fit....
        lOffset = 8
        SetRect m_ctlRect, 0, 0, lw, lh
        SetRect m_txtRect, lOffset, (lh / 2) - 8, lw - 64, lh - 4
        CopyRect tmpRect, m_txtRect
        DrawText UserControl.hdc, m_Caption, Len(m_Caption), tmpRect, DT_CALCRECT Or DT_WORDBREAK
        '   Setup the Caption alignment within the text area to match the property...
        Select Case m_CaptionAlign
            Case clCenter
                m_txtRect.Top = (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2
                m_lwFontAlign = DT_CENTER Or DT_VCENTER Or DT_WORDBREAK
            Case clLeft
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2, lw - 60, (lh + tmpRect.Bottom - tmpRect.Top) / 2
                m_lwFontAlign = DT_VCENTER Or DT_LEFT Or DT_WORDBREAK
            Case clRight
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2, lw - 60, (lh + tmpRect.Bottom - tmpRect.Top) / 2
                m_lwFontAlign = DT_VCENTER Or DT_RIGHT Or DT_WORDBREAK
            Case clTop
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, 4, lw - 60, lh
                m_lwFontAlign = DT_CENTER Or DT_TOP Or DT_WORDBREAK
            Case clBottom
                CopyRect m_txtRect, tmpRect
                SetRect m_txtRect, lOffset, (lh - 4) - (tmpRect.Bottom - tmpRect.Top), lw - 60, lh - 4 'lOffset, (lh / 2) - (tmpRect.Bottom - tmpRect.Top) / 2, lw - 8, (lh / 2) + (tmpRect.Bottom - tmpRect.Top)
                m_lwFontAlign = DT_CENTER Or DT_BOTTOM Or DT_WORDBREAK
        End Select
    End If
    '   Build the regions and frame them
    BuildRegion
    '   The paint the control...
    Refresh
End Sub

'The control is terminating - a good place to stop the subclasser
Private Sub UserControl_Terminate()
    On Error GoTo Catch
    'Stop all subclassing
    Call Subclass_StopAll
Catch:
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo UserControl_WriteProperties_Error
    
    '   Keep a copy of the properties for later...
    With PropBag
        Call .WriteProperty("BackColor", m_BackColor, &HFFFFFF)
        Call .WriteProperty("ButtonAlign", m_ButtonAlign, clbLeft)
        Call .WriteProperty("ButtonBackSize", m_ButtonBackSize, 34)
        Call .WriteProperty("ButtonIcon", m_ButtonIcon, clbNext)
        Call .WriteProperty("ButtonShape", m_ButtonShape, clbEllipse)
        Call .WriteProperty("ButtonToolTipText", m_ButtonToolTipText, vbNullString)
        Call .WriteProperty("ButtonToolTipText", UserControl.imgCustomPic.ToolTipText, vbNullString)
        Call .WriteProperty("ButtonToolTipText", UserControl.imgFailed.ToolTipText, vbNullString)
        Call .WriteProperty("ButtonToolTipText", UserControl.imgNext.ToolTipText, vbNullString)
        Call .WriteProperty("ButtonToolTipText", UserControl.imgSuccess.ToolTipText, vbNullString)
        Call .WriteProperty("Caption", m_Caption, UserControl.Extender.Name)
        Call .WriteProperty("CaptionAlign", m_CaptionAlign, clCenter)
        Call .WriteProperty("CornerSize", m_CornerSize, 12)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("Font", UserControl.Font)
        Call .WriteProperty("FontColor", m_FontColor, GetSysColor(COLOR_BTNTEXT))
        Call .WriteProperty("FontHighlightColor", m_FontHighlightColor, GetSysColor(COLOR_BTNTEXT))
        Call .WriteProperty("CustomColor", m_CustomColor, &HFF8080)
        Call .WriteProperty("Icon", UserControl.imgCustomPic, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("UseCustomColor", m_UseCustomColor, False)
        Call .WriteProperty("UseCustomIcon", m_UseCustomIcon, False)
        Call .WriteProperty("UseGradient", m_UseGradient, True)
    End With
    
    Exit Sub
    
UserControl_WriteProperties_Error:
End Sub
