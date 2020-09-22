VERSION 5.00
Begin VB.UserControl FTPicture 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6435
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   ToolboxBitmap   =   "FTPicture.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2520
      Left            =   1320
      Picture         =   "FTPicture.ctx":0312
      ScaleHeight     =   168
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   248
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderStyle     =   2  'Dash
      Height          =   495
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "FTPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Type BITMAPINFOHEADER
      biSize As Long
      biWidth As Long
      biHeight As Long
      biPlanes As Integer
      biBitCount As Integer
      biCompression As Long
      biSizeImage As Long
      biXPelsPerMeter As Long
      biYPelsPerMeter As Long
      biClrUsed As Long
      biClrImportant As Long
End Type

Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
End Type

Enum PBackStyle
    [Opaque] = 1
    [Transparent] = 0
End Enum

Enum PBorderstyle
    [No_Border] = 0
    [Color_Solid] = 1
    [Color_Dash] = 2
End Enum

Enum PStretch
    [Original_Size] = 0
    [Normal_Fastest] = 1
    [Halftone_Simple_Fast] = 2
    [Halftone_Filter_Medium] = 3
    [Halftone_Filter_Slow] = 4
End Enum

Dim MyBorderStyle As PBorderstyle
Dim MyStretchmode As PStretch

Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'----InitProperties------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub UserControl_InitProperties()
    UserControl.BackStyle = 0
    MyBorderStyle = 1
    MyStretchmode = 0
    BorderStylex
End Sub

'----Resize--------------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub UserControl_Resize()
 If MyBorderStyle <> 0 Then Shape1.Visible = False
 Shape1.Width = UserControl.Width / Screen.TwipsPerPixelX
 Shape1.Height = UserControl.Height / Screen.TwipsPerPixelY
 StrechModex
 If MyBorderStyle <> 0 Then Shape1.Visible = True
End Sub

'----Subs----------------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub StrechModex()
    If MyStretchmode = 0 Then
       UserControl.Width = Picture1.Width * Screen.TwipsPerPixelX
       UserControl.Height = Picture1.Height * Screen.TwipsPerPixelY
       Set UserControl.Picture = Picture1.Picture
       Set UserControl.MaskPicture = Picture1.Picture
    End If
    If MyStretchmode = 1 Then Normal
    If MyStretchmode = 2 Then Simple_Fast
    If MyStretchmode = 3 Then Filter_Medium
    If MyStretchmode = 4 Then Filter_Slow
End Sub

Private Sub BorderStylex()
    If MyBorderStyle = 0 Then
       Shape1.Visible = False
    ElseIf MyBorderStyle = 1 Then
       Shape1.BorderStyle = 1
       Shape1.Visible = True
    ElseIf MyBorderStyle = 2 Then
       Shape1.BorderStyle = 2
       Shape1.Visible = True
    End If
End Sub

'----Events--------------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'----Properties----------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
  MyBorderStyle = PropBag.ReadProperty("BorderStyle", 1)
  MyStretchmode = PropBag.ReadProperty("StretchMode", 1)
  UserControl.MaskColor = PropBag.ReadProperty("MaskColor", 16777215)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
  Shape1.BorderColor = PropBag.ReadProperty("BorderColor", &HFF00FF)
  Set Picture = PropBag.ReadProperty("Picture", Nothing)
  BorderStylex
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
  Call PropBag.WriteProperty("BorderStyle", MyBorderStyle, 1)
  Call PropBag.WriteProperty("StretchMode", MyStretchmode, 1)
  Call PropBag.WriteProperty("MaskColor", UserControl.MaskColor, 16777215)
  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
  Call PropBag.WriteProperty("BorderColor", Shape1.BorderColor, &HFF00FF)
  Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub



Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property



Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property
Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property



Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property



Public Property Get PictureWidth() As Long
    PictureWidth = Picture1.Width
End Property
Public Property Get PictureHeight() As Long
    PictureHeight = Picture1.Height
End Property



Public Property Get BorderColor() As OLE_COLOR
    BorderColor = Shape1.BorderColor
End Property
Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    Shape1.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property



Public Property Get MaskColor() As OLE_COLOR
    MaskColor = UserControl.MaskColor
End Property
Public Property Let MaskColor(ByVal New_MaskColor As OLE_COLOR)
    UserControl.MaskColor() = New_MaskColor
    PropertyChanged "MaskColor"
End Property



Public Property Get BorderStyle() As PBorderstyle
    BorderStyle = MyBorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As PBorderstyle)
    MyBorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    BorderStylex
End Property



Public Property Get StretchMode() As PStretch
Attribute StretchMode.VB_UserMemId = -520
    StretchMode = MyStretchmode
End Property
Public Property Let StretchMode(ByVal New_Stretch As PStretch)
    MyStretchmode = New_Stretch
    PropertyChanged "StretchMode"
    StrechModex
End Property



Public Property Get BackStyle() As PBackStyle
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property
Public Property Let BackStyle(ByVal New_BackStyle As PBackStyle)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property



Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property
Public Property Set Picture(ByVal New_Picture As Picture)
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set UserControl.Picture = New_Picture
    Set UserControl.MaskPicture = New_Picture
    Set Picture1.Picture = New_Picture
    StrechModex
    PropertyChanged "Picture"

End Property



'----Stretch----------------------------------------------------------------------------
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

Private Sub Normal()

UserControl.PaintPicture Picture1.Picture, _
0, 0, UserControl.Width / Screen.TwipsPerPixelX, UserControl.Height / Screen.TwipsPerPixelY, _
0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight
Set UserControl.Picture = UserControl.Image
Set UserControl.MaskPicture = UserControl.Image

End Sub

Private Sub Simple_Fast()

Dim BmpInfo As BITMAPINFO
Dim SPixels() As Byte
Dim UCWidth As Long, UCHeight As Long

UCWidth = UserControl.Width / Screen.TwipsPerPixelX
UCHeight = UserControl.Height / Screen.TwipsPerPixelY

With BmpInfo.bmiHeader
    .biSize = 40
    .biWidth = Picture1.ScaleWidth
    .biHeight = -Picture1.ScaleHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    .biSizeImage = 0
End With
      
ReDim SPixels(1 To 4, 1 To Picture1.ScaleWidth, 1 To Picture1.ScaleHeight)
GetDIBits Picture1.hDC, Picture1.Image, 0, Picture1.ScaleHeight, SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS

Call SetStretchBltMode(UserControl.hDC, 4)
Call StretchDIBits(UserControl.hDC, 0, 0, UCWidth, UCHeight, _
                              0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight _
                              , SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS, vbSrcCopy)

Set UserControl.Picture = UserControl.Image
Set UserControl.MaskPicture = UserControl.Image

End Sub


Private Sub Filter_Medium()

Dim SPixels() As Byte
Dim BmpInfo As BITMAPINFO
Dim Y As Long, X As Long, VhDC1 As Long, VMap1 As Long
Dim Rs As Long, Gs As Long, Bs As Long, Tot As Long
Dim iRed As Long, iGreen As Long, iBlue As Long
Dim UCWidth As Long, UCHeight As Long

UCWidth = UserControl.Width / Screen.TwipsPerPixelX
UCHeight = UserControl.Height / Screen.TwipsPerPixelY

With BmpInfo.bmiHeader
    .biSize = 40
    .biWidth = Picture1.ScaleWidth
    .biHeight = -Picture1.ScaleHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    .biSizeImage = 0
End With
      
ReDim SPixels(1 To 4, 1 To Picture1.ScaleWidth, 1 To Picture1.ScaleHeight)

GetDIBits Picture1.hDC, Picture1.Image, 0, Picture1.ScaleHeight, SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS

With BmpInfo.bmiHeader
    .biWidth = UCWidth
    .biHeight = -UCHeight
End With

VhDC1 = CreateCompatibleDC(UserControl.hDC)
VMap1 = CreateDIBSection(VhDC1, BmpInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
SelectObject VhDC1, VMap1

With BmpInfo.bmiHeader
    .biWidth = Picture1.ScaleWidth
    .biHeight = -Picture1.ScaleHeight
End With

Call SetStretchBltMode(VhDC1, 4)
Call StretchDIBits(VhDC1, 0, 0, UCWidth, UCHeight, _
                              0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight _
                              , SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS, vbSrcCopy)

With BmpInfo.bmiHeader
    .biWidth = UCWidth
    .biHeight = -UCHeight
End With

ReDim SPixels(1 To 4, 1 To UCWidth, 1 To UCHeight)

GetDIBits VhDC1, VMap1, 0, UCHeight, SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS

iRed = UserControl.MaskColor Mod 256
iGreen = ((UserControl.MaskColor And &HFF00) / 256&) Mod 256&
iBlue = (UserControl.MaskColor And &HFF0000) / 65536

For Y = 1 To UCHeight
    For X = 1 To UCWidth
      Rs = Abs(SPixels(3, X, Y) - iRed)
      Gs = Abs(SPixels(2, X, Y) - iGreen)
      Bs = Abs(SPixels(1, X, Y) - iBlue)
      Tot = Rs + Gs + Bs
      If Tot < 200 Then
         SPixels(3, X, Y) = iRed
         SPixels(2, X, Y) = iGreen
         SPixels(1, X, Y) = iBlue
      End If
    Next X
Next Y

SetDIBits UserControl.hDC, UserControl.Image, 0, UCHeight, SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS

DeleteDC VhDC1
DeleteObject VMap1

Set UserControl.Picture = UserControl.Image
Set UserControl.MaskPicture = UserControl.Image


End Sub

Private Sub Filter_Slow()

Dim SPixels() As Byte, DPixels() As Byte
Dim BmpInfo As BITMAPINFO
Dim Y As Long, X As Long, VhDC1 As Long, VMap1 As Long, VhDC2 As Long, VMap2 As Long
Dim iRed As Long, iGreen As Long, iBlue As Long
Dim UCWidth As Long, UCHeight As Long

UCWidth = UserControl.Width / Screen.TwipsPerPixelX
UCHeight = UserControl.Height / Screen.TwipsPerPixelY

With BmpInfo.bmiHeader
    .biSize = 40
    .biWidth = Picture1.ScaleWidth
    .biHeight = -Picture1.ScaleHeight
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    .biSizeImage = 0
End With
      
ReDim SPixels(1 To 4, 1 To Picture1.ScaleWidth, 1 To Picture1.ScaleHeight)
ReDim DPixels(1 To 4, 1 To Picture1.ScaleWidth, 1 To Picture1.ScaleHeight)
      
GetDIBits Picture1.hDC, Picture1.Image, 0, Picture1.ScaleHeight, SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS
      
iRed = UserControl.MaskColor Mod 256
iGreen = ((UserControl.MaskColor And &HFF00) / 256&) Mod 256&
iBlue = (UserControl.MaskColor And &HFF0000) / 65536

For Y = 1 To Picture1.ScaleHeight
    For X = 1 To Picture1.ScaleWidth
      If SPixels(3, X, Y) = iRed Then
      If SPixels(2, X, Y) = iGreen Then
      If SPixels(1, X, Y) = iBlue Then
         DPixels(3, X, Y) = SPixels(3, X, Y)
         DPixels(2, X, Y) = SPixels(2, X, Y)
         DPixels(1, X, Y) = SPixels(1, X, Y)
      End If
      End If
      End If
    Next X
Next Y

With BmpInfo.bmiHeader
    .biWidth = UCWidth
    .biHeight = -UCHeight
End With
VhDC1 = CreateCompatibleDC(UserControl.hDC)
VMap1 = CreateDIBSection(VhDC1, BmpInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)

VhDC2 = CreateCompatibleDC(UserControl.hDC)
VMap2 = CreateDIBSection(VhDC2, BmpInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)

SelectObject VhDC1, VMap1
SelectObject VhDC2, VMap2

With BmpInfo.bmiHeader
    .biWidth = Picture1.ScaleWidth
    .biHeight = -Picture1.ScaleHeight
End With
Call SetStretchBltMode(VhDC1, 4)
Call StretchDIBits(VhDC1, 0, 0, UCWidth, UCHeight, _
                              0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight _
                              , SPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS, vbSrcCopy)
If UCHeight > Picture1.Height Then
   Call SetStretchBltMode(VhDC2, 4)
Else
   Call SetStretchBltMode(VhDC2, 2)
End If
Call StretchDIBits(VhDC2, 0, 0, UCWidth, UCHeight, _
                              0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight _
                              , DPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS, vbSrcCopy)


Dim XPixels() As Byte, YPixels() As Byte, ZPixels() As Byte
With BmpInfo.bmiHeader
    .biWidth = UCWidth
    .biHeight = -UCHeight
End With

ReDim XPixels(1 To 4, 1 To UCWidth, 1 To UCHeight)
ReDim YPixels(1 To 4, 1 To UCWidth, 1 To UCHeight)
ReDim ZPixels(1 To 4, 1 To UCWidth, 1 To UCHeight)

GetDIBits VhDC1, VMap1, 0, UCHeight, XPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS
GetDIBits VhDC2, VMap2, 0, UCHeight, YPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS

For Y = 1 To UCHeight
    For X = 1 To UCWidth
      If YPixels(3, X, Y) = iRed Then
      If YPixels(2, X, Y) = iGreen Then
      If YPixels(1, X, Y) = iBlue Then
         ZPixels(3, X, Y) = iRed
         ZPixels(2, X, Y) = iGreen
         ZPixels(1, X, Y) = iBlue
         GoTo GoOn
      End If
      End If
      End If
      ZPixels(3, X, Y) = XPixels(3, X, Y)
      ZPixels(2, X, Y) = XPixels(2, X, Y)
      ZPixels(1, X, Y) = XPixels(1, X, Y)
GoOn:
    Next X
Next Y

SetDIBits UserControl.hDC, UserControl.Image, 0, UCHeight, ZPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS

    DeleteDC VhDC1
    DeleteObject VMap1
    DeleteDC VhDC2
    DeleteObject VMap2

Set UserControl.Picture = UserControl.Image
Set UserControl.MaskPicture = UserControl.Image

End Sub


'  ;)
