VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTF 
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkStyle 
      Caption         =   "No Caption"
      Height          =   225
      Index           =   1
      Left            =   3900
      TabIndex        =   13
      Top             =   4590
      Width           =   1425
   End
   Begin VB.CheckBox chkStyle 
      Caption         =   "No Borders"
      Height          =   210
      Index           =   3
      Left            =   4620
      TabIndex        =   12
      Top             =   3600
      Width           =   1230
   End
   Begin VB.CheckBox chkStyle 
      Caption         =   "Flat Borders"
      Height          =   210
      Index           =   2
      Left            =   4605
      TabIndex        =   9
      Top             =   3825
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   510
      Index           =   0
      Left            =   1020
      TabIndex        =   8
      Top             =   3855
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Erase"
      Height          =   510
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   3855
      Width           =   915
   End
   Begin VB.ComboBox cboAlign 
      Height          =   315
      ItemData        =   "frmTF.frx":0000
      Left            =   1980
      List            =   "frmTF.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4050
      Width           =   1815
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Text"
      Height          =   330
      Index           =   0
      Left            =   3885
      TabIndex        =   5
      ToolTipText     =   "Change Caption Color"
      Top             =   4035
      Width           =   690
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "Border"
      Height          =   330
      Index           =   1
      Left            =   4575
      TabIndex        =   4
      ToolTipText     =   "Change Border Color"
      Top             =   4035
      Width           =   690
   End
   Begin VB.CheckBox chkStyle 
      Caption         =   "Flat Caption"
      Height          =   210
      Index           =   0
      Left            =   3900
      TabIndex        =   3
      Top             =   4365
      Value           =   1  'Checked
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Caption         =   "VB Frame"
      Height          =   1890
      Left            =   150
      TabIndex        =   2
      Top             =   315
      Width           =   1425
   End
   Begin VB.PictureBox Picture1 
      Height          =   1890
      Left            =   1965
      ScaleHeight     =   122
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   233
      TabIndex        =   0
      Top             =   315
      Width           =   3555
      Begin VB.Label Label2 
         Caption         =   "^ Image is on the frame"
         Height          =   270
         Left            =   1275
         TabIndex        =   1
         Top             =   1395
         Width           =   1755
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   165
         Picture         =   "frmTF.frx":006F
         Top             =   375
         Width           =   2100
      End
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   4050
      Top             =   3015
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label is behind the frame"
      Height          =   255
      Left            =   3660
      TabIndex        =   11
      Top             =   1215
      Width           =   2520
   End
   Begin VB.Image Image2 
      Height          =   1260
      Left            =   1770
      Picture         =   "frmTF.frx":1CBB
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "< Image is behind the frame"
      Height          =   285
      Left            =   2565
      TabIndex        =   10
      Top             =   2415
      Width           =   2040
   End
   Begin VB.Image imgBkg 
      Height          =   3840
      Left            =   5445
      Picture         =   "frmTF.frx":BF80
      Top             =   4155
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Menu mnuPic 
      Caption         =   "Picture"
      Begin VB.Menu mnuPicPop 
         Caption         =   "Show Sample Background Image"
         Index           =   0
      End
      Begin VB.Menu mnuPicPop 
         Caption         =   "Remove Sample Background Image"
         Index           =   1
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuColors 
         Caption         =   "Shadow Color"
         Index           =   0
      End
      Begin VB.Menu mnuColors 
         Caption         =   "HighLight Color"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cColors(0 To 3) As Long
Private c_TransFrame As cTransParentFrame

Private Sub chkStyle_Click(Index As Integer)
    If chkStyle(Index).Value = 1 Then
        If Index > 1 Then
            If Index = 2 Then Index = 3 Else Index = 2
            chkStyle(Index).Value = 0
            cmdColor(1).Enabled = (Index <> 2)
        Else
            chkStyle(Abs(Index - 1)).Value = 0
            cmdColor(0).Enabled = (Index = 0)
        End If
    End If
End Sub

Private Sub cmdColor_Click(Index As Integer)

    mnuColors(1).Enabled = (chkStyle(Index).Value = 0)
    mnuPopup.Tag = Index * 2
    If mnuColors(1).Enabled Then
        PopupMenu mnuPopup
    Else
        Call mnuColors_Click(0)
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        
        With c_TransFrame
            If chkStyle(2) = 1 Then
                .BorderStyle = tbpFlat
            ElseIf chkStyle(3) = 1 Then
                .BorderStyle = tbpNone
            Else
                .BorderStyle = tbp3D
            End If
            .FlatCaption = (chkStyle(0).Value = 1)
            .SetTextColors cColors(0), cColors(1)
            .SetBorderColors cColors(2), cColors(3)
            .Align = cboAlign.ListIndex
            
            If chkStyle(1) = 1 Then
                .Caption = vbNullString
            Else
                .Caption = "Transparent Frame"
            End If
            
        End With
        c_TransFrame.Refresh
        
    Else
    
        Set Picture1.Picture = Nothing
        Picture1.Cls
        
    End If
End Sub

Private Sub Form_Load()

    Picture1.BorderStyle = 0
    cboAlign.ListIndex = 0
    cColors(0) = -1     ' initialize to class' Render optional parameters
    cColors(1) = -1
    cColors(2) = -1
    cColors(3) = -1
    
    Set c_TransFrame = New cTransParentFrame
    c_TransFrame.Attach Picture1
    c_TransFrame.Caption = "Transparent Frame"
    
End Sub

Private Sub mnuColors_Click(Index As Integer)

    With dlgColor
        If cColors(Val(mnuPopup.Tag) + Index) = -1 Then
            .Color = Picture1.ForeColor
        Else
            .Color = cColors(Val(mnuPopup.Tag) + Index)
        End If
        .Flags = cdlCCRGBInit
    End With
    On Error GoTo ExitRoutine
    dlgColor.ShowColor
    
    cColors(Val(mnuPopup.Tag) + Index) = dlgColor.Color
    
ExitRoutine:

End Sub

Private Sub mnuPicPop_Click(Index As Integer)
    If Index = 0 Then
        Set Me.Picture = imgBkg.Picture
    Else
        Set Me.Picture = Nothing
    End If
        c_TransFrame.Refresh
End Sub

