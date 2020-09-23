VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iCatcherLabel Example"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8970
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optActive 
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   4200
      TabIndex        =   30
      Top             =   3000
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.OptionButton optActive 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   4200
      TabIndex        =   29
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optActive 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   28
      Top             =   1320
      Width           =   255
   End
   Begin VB.OptionButton optActive 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   4200
      TabIndex        =   27
      Top             =   480
      Width           =   255
   End
   Begin VB.CheckBox chkUseGradients 
      Caption         =   "UseGradients"
      Height          =   255
      Left            =   6120
      TabIndex        =   26
      Top             =   2960
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox chkUseCustomIcon 
      Caption         =   "Custom Icon:"
      Height          =   255
      Left            =   7560
      TabIndex        =   25
      Top             =   2240
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CommandButton cmdCustomColor 
      Caption         =   "..."
      Height          =   275
      Left            =   5715
      TabIndex        =   23
      Top             =   3230
      Width           =   275
   End
   Begin VB.CommandButton cmdFontColor 
      Caption         =   "..."
      Height          =   275
      Left            =   5715
      TabIndex        =   16
      Top             =   2530
      Width           =   275
   End
   Begin VB.TextBox txtCornerSize 
      Enabled         =   0   'False
      Height          =   315
      Left            =   6120
      TabIndex        =   15
      Text            =   "12"
      Top             =   1780
      Width           =   1335
   End
   Begin VB.ComboBox cmbButtonShape 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1780
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7560
      TabIndex        =   11
      Top             =   3140
      Width           =   1335
   End
   Begin VB.CommandButton cmdLocateIcon 
      Caption         =   "..."
      Height          =   275
      Left            =   8595
      TabIndex        =   10
      Top             =   2530
      Width           =   275
   End
   Begin VB.ComboBox cmbButtonIcon 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1780
      Width           =   1335
   End
   Begin VB.TextBox txtCaptionValue 
      Height          =   1050
      Left            =   6120
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "frmMain.frx":08CA
      Top             =   340
      Width           =   2775
   End
   Begin VB.ComboBox cmbButtonAlign 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1060
      Width           =   1455
   End
   Begin VB.ComboBox cmbCaptionAlign 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   340
      Width           =   1455
   End
   Begin prjiCatcherLabel.iCatcherLabel iCatcherLabel 
      Height          =   735
      Index           =   0
      Left            =   80
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      Caption         =   "The text in this caption field is justified as ""Centered"" and suports multi-line, text wrapping, and Custom Colors."
      CornerSize      =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontHighlightColor=   8388608
      CustomColor     =   14215660
      Icon            =   "frmMain.frx":0940
      UseCustomColor  =   -1  'True
   End
   Begin VB.TextBox txtFontColor 
      Height          =   315
      Left            =   4560
      TabIndex        =   17
      Text            =   "Locate Color..."
      Top             =   2500
      Width           =   1455
   End
   Begin VB.CommandButton cmdFontHotColor 
      Caption         =   "..."
      Height          =   275
      Left            =   7155
      TabIndex        =   19
      Top             =   2530
      Width           =   275
   End
   Begin VB.TextBox txtFontHotColor 
      Height          =   315
      Left            =   6120
      TabIndex        =   20
      Text            =   "Locate Color..."
      Top             =   2500
      Width           =   1335
   End
   Begin VB.TextBox txtIconLocation 
      Height          =   315
      Left            =   7560
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Locate Icon..."
      Top             =   2500
      Width           =   1335
   End
   Begin VB.CheckBox chkUseCustomColor 
      Caption         =   "Custom Color:"
      Height          =   255
      Left            =   4560
      TabIndex        =   22
      Top             =   2960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtCustomColor 
      Height          =   315
      Left            =   4560
      TabIndex        =   24
      Text            =   "Locate Color..."
      Top             =   3200
      Width           =   1455
   End
   Begin prjiCatcherLabel.iCatcherLabel iCatcherLabel 
      Height          =   735
      Index           =   1
      Left            =   80
      TabIndex        =   31
      Top             =   1080
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      ButtonAlign     =   1
      ButtonIcon      =   1
      Caption         =   "The text in this caption field is justified as ""Right"" and suports multi-line and text wrapping."
      CaptionAlign    =   2
      CornerSize      =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontHighlightColor=   8388608
      Icon            =   "frmMain.frx":095C
   End
   Begin prjiCatcherLabel.iCatcherLabel iCatcherLabel 
      Height          =   735
      Index           =   2
      Left            =   80
      TabIndex        =   32
      Top             =   1920
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      ButtonIcon      =   2
      Caption         =   "The text in this caption field is justified as ""Left"" and suports multi-line and text wrapping."
      CaptionAlign    =   1
      CornerSize      =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontHighlightColor=   8388608
      Icon            =   "frmMain.frx":0978
   End
   Begin prjiCatcherLabel.iCatcherLabel iCatcherLabel 
      Height          =   735
      Index           =   3
      Left            =   80
      TabIndex        =   33
      Top             =   2760
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      ButtonAlign     =   1
      ButtonIcon      =   3
      Caption         =   "The text in this caption field is justified as ""Centered"" and suports multi-line, text wrapping, and Custom Colors and Icons!"
      CornerSize      =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontHighlightColor=   8388608
      Icon            =   "frmMain.frx":0994
      UseCustomColor  =   -1  'True
      UseCustomIcon   =   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "Font Color:"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Button Shape:"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   1545
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Corner Size:"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   1545
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Button Icon:"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   1545
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Caption:"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   105
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Caption Align:"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   105
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Button Align:"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   825
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Font (Hot) Color:"
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      Top             =   2265
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
'       if it were not for these two programmers. Also, I want to thank Paul Paul Turcksin
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
'       23Sep05 - Cleaned up the comments in each sub and added additional
'                 error handling to selected routines.
'
'   Force Declarations
Option Explicit
Dim bLoading        As Boolean      'Loading Flag to prevent unwanted event calls
Dim m_ActiveControl As Long         'Local Cache for the Active iCatcherLabel

Private Sub chkUseCustomColor_Click()
    With Me
        If Not bLoading Then
            '   Set the UseCustonColor Flag
            .iCatcherLabel(m_ActiveControl).UseCustomColor = CBool(.chkUseCustomColor.Value)
        End If
    End With
End Sub

Private Sub chkUseCustomIcon_Click()
    With Me
        If Not bLoading Then
            '   Set the UseCustomIcon Flag
            .iCatcherLabel(m_ActiveControl).UseCustomIcon = CBool(.chkUseCustomIcon.Value)
        End If
    End With
End Sub

Private Sub chkUseGradients_Click()
    With Me
        If Not bLoading Then
            '   Set the UseGradients Flag
            .iCatcherLabel(m_ActiveControl).UseGradient = CBool(.chkUseGradients.Value)
        End If
    End With
End Sub

Private Sub cmbButtonAlign_Click()
    With Me
        If Not bLoading Then
            '   Set the Button Alignment Flag
            .iCatcherLabel(m_ActiveControl).ButtonAlign = .cmbButtonAlign.ListIndex
        End If
    End With
End Sub

Private Sub cmbButtonIcon_Click()
    With Me
        If Not bLoading Then
            '   Set the Button Icon (Default Icons)
            .chkUseCustomIcon.Value = IIf(.cmbButtonIcon.ListIndex = 3, vbChecked, vbUnchecked)
            .iCatcherLabel(m_ActiveControl).ButtonIcon = .cmbButtonIcon.ListIndex
        End If
    End With
End Sub

Private Sub cmbButtonShape_Click()
    With Me
        If Not bLoading Then
            '   Set the Button Shape property
            .iCatcherLabel(m_ActiveControl).ButtonShape = .cmbButtonShape.ListIndex
            If .cmbButtonShape.ListIndex = 1 Then
                .txtCornerSize.Enabled = True
                .Label6.Enabled = True
            Else
                .txtCornerSize.Enabled = False
                .Label6.Enabled = False
            End If
        End If
    End With
End Sub

Private Sub cmbCaptionAlign_Click()
    With Me
        If Not bLoading Then
            '   Set the Caption Aligment Flag
            .iCatcherLabel(m_ActiveControl).CaptionAlign = .cmbCaptionAlign.ListIndex
        End If
    End With
End Sub

Private Sub cmdExit_Click()
    '   Close the test harness
    Form_Terminate
End Sub

Private Sub cmdFontColor_Click()
    Dim ColorVal        As SelectedColor
    
    With Me
        '   Select a Color from a dialog
        If Not bLoading Then
            ColorVal = ShowColor(.hwnd, True)
            If ColorVal.bCanceled = False Then
                '   Set the ForeColor ....Text Color
                .txtFontColor.Text = ColorVal.oSelectedColor
                .txtFontColor.Text = HexColorStr(.txtFontColor.Text)
                .iCatcherLabel(m_ActiveControl).FontColor = CLng(.txtFontColor.Text)
            End If
        End If
    End With
End Sub

Private Sub cmdFontHotColor_Click()
    Dim ColorVal        As SelectedColor
    
    With Me
        '   Select a Color from a dialog
        If Not bLoading Then
            ColorVal = ShowColor(.hwnd, True)
            If ColorVal.bCanceled = False Then
                '   Set the Text Hot (Highlight) Color
                .txtFontHotColor.Text = ColorVal.oSelectedColor
                .txtFontHotColor.Text = HexColorStr(.txtFontHotColor.Text)
                .iCatcherLabel(m_ActiveControl).FontHighlightColor = CLng(.txtFontHotColor.Text)
            End If
        End If
    End With

End Sub

Private Sub cmdCustomColor_Click()
    Dim ColorVal        As SelectedColor
    
    With Me
        '   Select a Color from a dialog
        If Not bLoading Then
            ColorVal = ShowColor(.hwnd, True)
            If ColorVal.bCanceled = False Then
                '   Set the Control Highlighted Color
                .txtCustomColor.Text = ColorVal.oSelectedColor
                .txtCustomColor.Text = HexColorStr(.txtCustomColor.Text)
                .iCatcherLabel(m_ActiveControl).CustomColor = CLng(.txtCustomColor.Text)
            End If
        End If
    End With
End Sub

Private Sub cmdLocateIcon_Click()
    Dim FileName        As SelectedFile
    
    With Me
        '   Select an Icon from a dialog
        If Not bLoading Then
            With FileDialog
                .sFilter = "Icon Files (*.ico)"
                .sDefFileExt = "ico"
                .nFileExt = 1
                .sInitDir = App.Path & "\Graphics"
            End With
            FileName = ShowOpen(.hwnd, True)
            If FileName.bCanceled = False Then
                '   Set the Icon to the Custom Image Control
                .txtIconLocation.Text = FileName.sFiles(1)
                Set .iCatcherLabel(m_ActiveControl).Icon = LoadPicture(.txtIconLocation.Text)
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    With Me
        '   Init the Test Harness with the default properties...
        bLoading = True
        .Caption = "iCatcherLabel Test Harness - v" & App.Major & "." & App.Minor & "." & App.Revision
        With .cmbButtonAlign
            .AddItem "clbLeft"
            .AddItem "clbRight"
            .ListIndex = 1
        End With
        With .cmbButtonIcon
            .AddItem "clbNext"
            .AddItem "clbSuccess"
            .AddItem "clbFailed"
            .AddItem "clbCustom"
            .ListIndex = 3
        End With
        With .cmbButtonShape
            .AddItem "clbEllipse"
            .AddItem "clbRndRect"
            .AddItem "clbRectangle"
            .ListIndex = 0
        End With
        With .cmbCaptionAlign
            .AddItem "clCenter"
            .AddItem "clLeft"
            .AddItem "clRight"
            .AddItem "clTop"
            .AddItem "clBottom"
            .ListIndex = 0
        End With
        m_ActiveControl = 3
        .txtCornerSize = .iCatcherLabel(m_ActiveControl).CornerSize + 8
        .txtFontColor = HexColorStr(&H0)
        .txtFontHotColor = HexColorStr(&H800000)
        .txtCustomColor = HexColorStr(&HFF8080)
        With .iCatcherLabel(0)
            .ButtonAlign = clbLeft
            .ButtonIcon = clbNext
            .ButtonShape = clbEllipse
            .CaptionAlign = clCenter
            .ButtonToolTipText = "This is Button 1"
            .CornerSize = 20
            .CustomColor = &H8000000F '(ButtonFace)
            .UseCustomColor = True
            .UseCustomIcon = False
            .UseGradient = True
        End With
        With .iCatcherLabel(1)
            .UseCustomColor = False
            .UseCustomIcon = False
            .UseGradient = True
            .ButtonAlign = clbRight
            .ButtonIcon = clbSuccess
            .ButtonShape = clbEllipse
            .CaptionAlign = clRight
            .ButtonToolTipText = "This is Button 2"
            .CornerSize = 20
        End With
        With .iCatcherLabel(2)
            .UseCustomColor = False
            .UseCustomIcon = False
            .UseGradient = True
            .ButtonAlign = clbLeft
            .ButtonIcon = clbFailed
            .ButtonShape = clbEllipse
            .CaptionAlign = clLeft
            .ButtonToolTipText = "This is Button 3"
            .CornerSize = 20
        End With
        With .iCatcherLabel(3)
            .UseCustomColor = True
            .UseCustomIcon = True
            .UseGradient = True
            .ButtonAlign = clbRight
            .ButtonIcon = clbCustom
            .ButtonShape = clbRndRect
            .CaptionAlign = clCenter
            .ButtonToolTipText = "This is Button 4"
            .CornerSize = 20
        End With
        .optActive(m_ActiveControl).Value = True
        bLoading = False
        '   Now make sure the correct controls info is updated
        '   on the GUI...
        GetControlSettings
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '   Set our form to nothing
    Set frmMain = Nothing
    '   Unload it...
    Unload Me
    '   Make sure we are done...
    End
End Sub

Private Sub Form_Terminate()
    '   Set our form to nothing
    Set frmMain = Nothing
    '   Unload it...
    Unload Me
    '   Make sure we are done...
    End
End Sub

Private Sub GetControlSettings()
    With Me
        '   This sets the disaply objects to the state of the selected control
        If Not bLoading Then
            '   This flag keeps the local event handlers from logging events
            '   while we update the properties on the test harness...
            bLoading = True
            '   Load all of the selected controls properties and diplay them on
            '   the GUI. This allows for dynamic control modifications with more
            '   than one control in a control array
            .txtCaptionValue.Text = .iCatcherLabel(m_ActiveControl).Caption
            .cmbCaptionAlign.ListIndex = .iCatcherLabel(m_ActiveControl).CaptionAlign
            .cmbButtonAlign.ListIndex = .iCatcherLabel(m_ActiveControl).ButtonAlign
            .cmbButtonShape.ListIndex = .iCatcherLabel(m_ActiveControl).ButtonShape
            If .cmbButtonShape.ListIndex = 1 Then
                .txtCornerSize.Enabled = True
                .Label6.Enabled = True
            Else
                .txtCornerSize.Enabled = False
                .Label6.Enabled = False
            End If
            .txtCornerSize.Text = .iCatcherLabel(m_ActiveControl).CornerSize
            .cmbButtonIcon.ListIndex = .iCatcherLabel(m_ActiveControl).ButtonIcon
            .txtFontColor.Text = HexColorStr(.iCatcherLabel(m_ActiveControl).FontColor)
            .txtFontHotColor.Text = HexColorStr(.iCatcherLabel(m_ActiveControl).FontHighlightColor)
            .txtCustomColor.Text = HexColorStr(.iCatcherLabel(m_ActiveControl).CustomColor)
            .chkUseCustomIcon.Value = IIf(.iCatcherLabel(m_ActiveControl).UseCustomIcon = True, vbChecked, vbUnchecked)
            .chkUseCustomColor.Value = IIf(.iCatcherLabel(m_ActiveControl).UseCustomColor = True, vbChecked, vbUnchecked)
            .chkUseGradients.Value = IIf(.iCatcherLabel(m_ActiveControl).UseGradient = True, vbChecked, vbUnchecked)
            bLoading = False
        End If
    End With
End Sub

Private Sub iCatcherLabel_ButtonClick(Index As Integer)
    MsgBox "You Pressed the iCatcherLabels Button....", vbInformation, "iCatcherLabel"
    Me.optActive(Index).Value = True
End Sub

Private Sub iCatcherLabel_ButtonHover(Index As Integer, X As Single, Y As Single)
    Debug.Print "Hover on iCatcherButton(" & Index & ") at Coordinates..." & X & ", " & Y
End Sub

Private Sub iCatcherLabel_Click(Index As Integer)
    Debug.Print "You Clicked on the Label..."
    Me.optActive(Index).Value = True
    '   Get the Usercontrol Settings and update the GUI
    Call GetControlSettings
End Sub

Private Sub iCatcherLabel_DblClick(Index As Integer)
    Debug.Print "You Double Clicked on the Label..."
End Sub

Private Sub iCatcherLabel_Hover(Index As Integer, X As Single, Y As Single)
    Debug.Print "Hover on iCatcherLabel(" & Index & ") at Coordinates..." & X & ", " & Y
End Sub

Private Sub iCatcherLabel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "MouseDown on Label at Coordinates..." & X & ", " & Y
End Sub

Private Sub iCatcherLabel_MouseEnter(Index As Integer)
    Debug.Print "MouseEnter iCatcherLabel " & Index
End Sub

Private Sub iCatcherLabel_MouseLeave(Index As Integer)
    Debug.Print "MouseLeave iCatcherLabel " & Index
End Sub

Private Sub optActive_Click(Index As Integer)
    '   Set our active control index
    m_ActiveControl = Index
    '   Get the Usercontrol Settings and update the GUI
    Call GetControlSettings
End Sub

Private Sub txtCaptionValue_Change()
    With Me
        If Not bLoading Then
            '   Update the Label with a new caption
            .iCatcherLabel(m_ActiveControl).Caption = txtCaptionValue.Text
            cmbCaptionAlign_Click
        End If
    End With
End Sub

Private Sub txtCornerSize_Change()
    With Me
        If Not bLoading Then
            '   Update the Labels Corner Size....this is for the
            '   buttons backdrop corners and not the control as a whole
            If IsNumeric(.txtCornerSize.Text) Then
                .iCatcherLabel(m_ActiveControl).CornerSize = CLng(.txtCornerSize.Text)
            Else
                MsgBox "Please Select a Valid Corner Size!", vbExclamation, "iCatcherLabel"
            End If
        End If
    End With
End Sub

Private Sub txtFontColor_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If Not bLoading Then
            '   Set the Font color
            If KeyCode = 13 Then
                If IsNumeric(.txtFontColor) Then
                    .txtFontColor.Text = HexColorStr(.txtFontColor.Text)
                    .iCatcherLabel(m_ActiveControl).FontColor = CLng(.txtFontColor.Text)
                Else
                    MsgBox "Please Enter a Valid Color Value!", vbExclamation, "iCatcherLabel"
                End If
            End If
        End If
    End With

End Sub

Private Sub txtFontColor_LostFocus()
    If Not bLoading Then
        '   Deligate the call to the correct event handler...
        Call txtFontColor_KeyDown(13, 0)
    End If
End Sub

Private Sub txtFontHotColor_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If Not bLoading Then
            '   Set the Font Hot (Highlight) color
            If KeyCode = 13 Then
                If IsNumeric(.txtFontColor) Then
                    .txtFontHotColor.Text = HexColorStr(.txtFontHotColor.Text)
                    .iCatcherLabel(m_ActiveControl).FontHighlightColor = CLng(.txtFontHotColor.Text)
                Else
                    MsgBox "Please Enter a Valid Color Value!", vbExclamation, "iCatcherLabel"
                End If
            End If
        End If
    End With
End Sub

Private Sub txtFontHotColor_LostFocus()
    If Not bLoading Then
        '   Deligate the call to the correct event handler...
        Call txtFontHotColor_KeyDown(13, 0)
    End If
End Sub

Private Sub txtCustomColor_KeyDown(KeyCode As Integer, Shift As Integer)
    With Me
        If Not bLoading Then
            If KeyCode = 13 Then
                '   Set the Highlight color
                If IsNumeric(.txtFontColor) Then
                    .txtCustomColor.Text = HexColorStr(.txtCustomColor.Text)
                    .iCatcherLabel(m_ActiveControl).CustomColor = CLng(.txtCustomColor.Text)
                Else
                    MsgBox "Please Enter a Valid Color Value!", vbExclamation, "iCatcherLabel"
                End If
            End If
        End If
    End With
End Sub

Private Sub txtCustomColor_LostFocus()
    If Not bLoading Then
        '   Deligate the call to the correct event handler...
        Call txtCustomColor_KeyDown(13, 0)
    End If
End Sub

Private Function HexColorStr(lValue As Long) As String
    '   Convert a standard color long value into a Hex Equivalant....
    HexColorStr = "&H" & Hex(lValue)
End Function
