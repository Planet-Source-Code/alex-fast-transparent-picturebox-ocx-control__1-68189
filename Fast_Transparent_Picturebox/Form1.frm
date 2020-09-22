VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\AFTPicture.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   DrawWidth       =   20
   LinkTopic       =   "Form1"
   ScaleHeight     =   597
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   StartUpPosition =   3  'Windows Default
   Begin FTPicturebox.FTPicture FTPicture1 
      Height          =   6900
      Left            =   2640
      TabIndex        =   28
      Top             =   960
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   12171
      StretchMode     =   0
      MaskColor       =   16711935
      BorderColor     =   0
      Picture         =   "Form1.frx":0000
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   25
      Top             =   7680
      Width           =   2295
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   7920
      TabIndex        =   24
      Top             =   480
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "StrechMode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   2295
      Begin VB.CommandButton CmdHalftone_Filter_Slow 
         Caption         =   "Halftone_Filter_Slow"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton CmdOriginal_Size 
         Caption         =   "Original_Size"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton CmdNormal_Fastest 
         Caption         =   "Normal_Fastest"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton CmdHalftone_Simple_Fast 
         Caption         =   "Halftone_Simple_Fast"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton CmdHalftone_Filter_Medium 
         Caption         =   "Halftone_Filter_Medium"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7920
      TabIndex        =   12
      Text            =   "50"
      Top             =   120
      Width           =   495
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2640
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   120
      Value           =   50
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Picture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   6840
      Width           =   2295
      Begin VB.CommandButton CmdPicture 
         Caption         =   "Set Picture"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "BorderColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   2295
      Begin VB.PictureBox PicBorder 
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdBordercolor 
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "BorderStyle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
      Begin VB.CommandButton CmdColor_Dash 
         Caption         =   "Color_Dash"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdColor_Solid 
         Caption         =   "Color_Solid"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdNo_Border 
         Caption         =   "No_Border"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "MaskColor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   2295
      Begin VB.PictureBox PicMask 
         Height          =   255
         Left            =   1320
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CmdMaskcolor 
         Caption         =   "Color"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "BackStyle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin VB.CommandButton CmdTransparent 
         Caption         =   "Transparent"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton CmdOpaque 
         Caption         =   "Opaque"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2640
      TabIndex        =   23
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'                       FAST TRANSPARENT PICTUREBOX OCX CONTROL                         '
'                                                                                       '
'***************************************************************************************'
'                                                                                       '
' COPYRIGHT: You can do what you want with this code, but if you are thinking to        '
' improve it, please upload it to "Planet Source Code" as a new version, that way       '
' we all can easily find & use it.                                                      '
'                                                                                       '
' NOTE: If you use "Halftone_Filter_Medium" be sure there aren't similar colors         '
'       like MaskColor in the visible area of the Picture.                              '
'       Other way, take in mind you can use "Halftone_Filter_Slow" a little slowly.     '
'                                                                                       '
' Special Thanks to "BattleStorm" for the "clsTimer Class" very good to test speed.     '
'                                                                                       '
' Made in Spain.                                                                        '
'****************************************************************************************

Option Explicit
Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private CodeTimer As clsTimer

Private Sub Form_Load()
  PicBorder.BackColor = FTPicture1.BorderColor
  PicMask.BackColor = FTPicture1.MaskColor
  FTPicture1.MousePointer = vbSizeAll
End Sub

'---- Resizing ----------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub HScroll1_Change()
  
  If FTPicture1.StretchMode = 0 Then Label1.Caption = " Resized to  (Original_Size)  in :"
  If FTPicture1.StretchMode = 1 Then Label1.Caption = " Stretched with  (Normal_Fastest)  in:"
  If FTPicture1.StretchMode = 2 Then Label1.Caption = " Stretched with  (Halftone_Simple_Fast)  in:"
  If FTPicture1.StretchMode = 3 Then Label1.Caption = " Stretched with  (Halftone_Filter_Medium)  in:"
  If FTPicture1.StretchMode = 4 Then Label1.Caption = " Stretched with  (Halftone_Filter_Slow)  in:"
  
  Text1.Text = HScroll1.Value
  
  Call StartTiming
  FTPicture1.Move FTPicture1.Left, FTPicture1.Top, _
  FTPicture1.PictureWidth * HScroll1.Value / 50, _
  FTPicture1.PictureHeight * HScroll1.Value / 50
  Call StopTiming
  
  'FTPicture1.Width = WI * HScroll1.Value / 50
  'FTPicture1.Height = HE * HScroll1.Value / 50
  
End Sub

'---- Events ---------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label2.Caption = ""
  Label3.Caption = ""
End Sub

Private Sub FTPicture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReleaseCapture
  SendMessage FTPicture1.hWnd, WM_NCLBUTTONDOWN, 2, 0&
  Screen.MousePointer = vbCustom
  Label2.Caption = " MouseMove"
End Sub

Private Sub FTPicture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.Caption = " MouseDown"
End Sub

Private Sub FTPicture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.Caption = " MouseUp"
End Sub

'---- BackStyle ---------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub CmdOpaque_Click()
  FTPicture1.BackStyle = Opaque
  CmdOpaque.Enabled = False
  CmdTransparent.Enabled = True
End Sub

Private Sub CmdTransparent_Click()
  FTPicture1.BackStyle = Transparent
  CmdOpaque.Enabled = True
  CmdTransparent.Enabled = False
End Sub

'---- BorderStyle -------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub CmdNo_Border_Click()
  FTPicture1.BorderStyle = No_Border
  CmdColor_Dash.Enabled = True
  CmdColor_Solid.Enabled = True
  CmdNo_Border.Enabled = False
End Sub

Private Sub CmdColor_Solid_Click()
  FTPicture1.BorderStyle = Color_Solid
  CmdColor_Dash.Enabled = True
  CmdColor_Solid.Enabled = False
  CmdNo_Border.Enabled = True
End Sub

Private Sub CmdColor_Dash_Click()
  FTPicture1.BorderStyle = Color_Dash
  CmdColor_Dash.Enabled = False
  CmdColor_Solid.Enabled = True
  CmdNo_Border.Enabled = True
End Sub

'---- BorderColor -------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub CmdBordercolor_Click()
  CommonDialog1.ShowColor
  FTPicture1.BorderColor = CommonDialog1.Color
  PicBorder.BackColor = CommonDialog1.Color
End Sub

'---- MaskColor ---------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub CmdMaskcolor_Click()
  CommonDialog1.ShowColor
  Label1.Caption = " MaskColor applied in :"
  Call StartTiming
  FTPicture1.MaskColor = CommonDialog1.Color
  Call StopTiming
  PicMask.BackColor = CommonDialog1.Color
End Sub

'---- StretchMode -------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub CmdOriginal_Size_Click()
  Label1.Caption = " Resized to  (Original_Size)  in :"
  Call StartTiming
  FTPicture1.StretchMode = Original_Size
  Call StopTiming
  CmdOriginal_Size.Enabled = False
  CmdNormal_Fastest.Enabled = True
  CmdHalftone_Simple_Fast.Enabled = True
  CmdHalftone_Filter_Medium.Enabled = True
  CmdHalftone_Filter_Slow.Enabled = True
End Sub

Private Sub CmdNormal_Fastest_Click()
  Label1.Caption = " Stretched with  (Normal_Fastest)  in:"
  Call StartTiming
  FTPicture1.StretchMode = Normal_Fastest
  Call StopTiming
  CmdOriginal_Size.Enabled = True
  CmdNormal_Fastest.Enabled = False
  CmdHalftone_Simple_Fast.Enabled = True
  CmdHalftone_Filter_Medium.Enabled = True
  CmdHalftone_Filter_Slow.Enabled = True
End Sub

Private Sub CmdHalftone_Simple_Fast_Click()
  Label1.Caption = " Stretched with  (Halftone_Simple_Fast)  in:"
  Call StartTiming
  FTPicture1.StretchMode = Halftone_Simple_Fast
  Call StopTiming
  CmdOriginal_Size.Enabled = True
  CmdNormal_Fastest.Enabled = True
  CmdHalftone_Simple_Fast.Enabled = False
  CmdHalftone_Filter_Medium.Enabled = True
  CmdHalftone_Filter_Slow.Enabled = True
End Sub

Private Sub CmdHalftone_Filter_Medium_Click()
  Label1.Caption = " Stretched with  (Halftone_Filter_Medium)  in:"
  Call StartTiming
  FTPicture1.StretchMode = Halftone_Filter_Medium
  Call StopTiming
  CmdOriginal_Size.Enabled = True
  CmdNormal_Fastest.Enabled = True
  CmdHalftone_Simple_Fast.Enabled = True
  CmdHalftone_Filter_Medium.Enabled = False
  CmdHalftone_Filter_Slow.Enabled = True
End Sub

Private Sub CmdHalftone_Filter_Slow_Click()
  Label1.Caption = " Stretched with  (Halftone_Filter_Slow)  in:"
  Call StartTiming
  FTPicture1.StretchMode = Halftone_Filter_Slow
  Call StopTiming
  CmdOriginal_Size.Enabled = True
  CmdNormal_Fastest.Enabled = True
  CmdHalftone_Simple_Fast.Enabled = True
  CmdHalftone_Filter_Medium.Enabled = True
  CmdHalftone_Filter_Slow.Enabled = False
End Sub

'----- Picture ----------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub CmdPicture_Click()
 On Error GoTo Error_exit
 Dim TempPic As StdPicture
  
  CommonDialog1.CancelError = True
  CommonDialog1.Filter = "BMP files (*.bmp)|*.BMP|All Files(*.*)|*.*"
  CommonDialog1.InitDir = App.Path
  CommonDialog1.Action = 1

  Set TempPic = LoadPicture(CommonDialog1.FileName)
  
  Label1.Caption = " Set new Picture and applied Mask and Strech in :"
  Call StartTiming
  Set FTPicture1.Picture = TempPic
  Call StopTiming
Error_exit:
End Sub


'---- Timing ------------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub StartTiming()
    Set CodeTimer = New clsTimer
    CodeTimer.StartTimer
End Sub

Private Sub StopTiming()
   CodeTimer.StopTimer
   Text2.Text = CodeTimer.Elasped & " ms"
End Sub




