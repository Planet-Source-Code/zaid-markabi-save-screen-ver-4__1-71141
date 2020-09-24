VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Save Screen V.4"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   0
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      Begin VB.Shape L 
         BorderColor     =   &H00C0C000&
         BorderStyle     =   2  'Dash
         FillColor       =   &H0000FFFF&
         FillStyle       =   7  'Diagonal Cross
         Height          =   9000
         Left            =   0
         Top             =   0
         Width           =   12000
      End
      Begin VB.Shape LL 
         BorderColor     =   &H00C00000&
         BorderStyle     =   3  'Dot
         FillColor       =   &H00FF8080&
         FillStyle       =   4  'Upward Diagonal
         Height          =   9000
         Left            =   0
         Top             =   0
         Width           =   12000
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "800*600*800*600"
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1935
      End
   End
   Begin VB.PictureBox Pi 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3135
      ScaleWidth      =   4095
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   3000
         Top             =   720
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   2400
         Top             =   720
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1800
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
         Begin VB.FileListBox File1 
            Height          =   870
            Left            =   1200
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Timer Timer1 
            Left            =   600
            Top             =   120
         End
         Begin VB.Timer Timer2 
            Interval        =   100
            Left            =   120
            Top             =   120
         End
         Begin VB.Label LDD 
            Caption         =   "OFF"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label8 
            Caption         =   "ON"
            Height          =   255
            Left            =   600
            TabIndex        =   13
            Tag             =   "0"
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Image PowerButton2 
         Height          =   255
         Left            =   2960
         Top             =   400
         Width           =   375
      End
      Begin VB.Image PowerButton1 
         Height          =   255
         Left            =   2600
         Top             =   400
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   2520
         Picture         =   "Form1.frx":15162
         Top             =   360
         Width           =   885
      End
      Begin VB.Image PowerButton3 
         Height          =   450
         Left            =   240
         Picture         =   "Form1.frx":163EC
         Top             =   2400
         Width           =   2310
      End
      Begin VB.Image PowerButton4 
         Height          =   450
         Left            =   2520
         Picture         =   "Form1.frx":19A8E
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Options :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Tag             =   "Start TOP"
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Tag             =   "Start TOP"
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Tag             =   "Start LEFT"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "600"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "Long Pic D"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "800"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "Long Pic >"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3840
         TabIndex        =   7
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Pictures In Sec. :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808080&
         Caption         =   " SaveScreen V.4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Programmed By [ Zaid Markabi ]
' ___________________________________________________________________________________________________
'|                                                                                                   |\_______________________
'|  ###############        ###         #####   ######                ######    #####                 |                        |\0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'| ##############         #####         ###     ##   ##               ######  #####                  |      Zaid Markabi      |=\ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|         ####          ### ###        ###     ##    ##              ##  ## ##  ##                  |                        |==\0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'|       ###            ###   ###       ###     ##     ##    #####    ##   ###   ##                  | zaidmarkabi@yahoo.com  |===\ 1 __________________________________  0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 1 0 0 0 1
'|     ###             ###########      ###     ##     ##   ####      ##    #    ##                  |                        |====|>| Development For Our Digital Life | 1 1 0 0 1 1 1 0 1 0 0 1 0 0 0 1 1 0 1 0
'|   ###              #############     ###     ##    ##              ##         ##      A R K A B I | VisualBasic Programmer |===/ 1|__________________________________| 0 1 1 0 1 0 0 0 1 1 1 0 1 0 1 1 0 1 0 0
'| ##############    ###         ###    ###     ##   ##               ##         ##     ############ |                        |==/0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'| ###############   ###         ###   #####   ######                ####       ####   ### 2008 ###  |Syria(Arab Area)-Tartuse|=/ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|                                                                                    ############   | _______________________|/0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'|___________________________________________________________________________________________________|/
' Em@il Me Please : zaidmarkabi@yahoo.com
' I hope to hear from you soon ,

Option Explicit
Option Base 0
Dim picht As Integer
Dim picwt As Integer
Dim clflag As Boolean

Dim PictureNumbers32 As Integer

Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   TOp As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Dim sFile As String
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Sub LoadNewDoc()
Picture1.Picture = LoadPicture(sFile)
End Sub
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long
   Dim pic As PicBmp
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   With pic
      .Size = Len(pic)
      .Type = vbPicTypeBitmap
      .hBmp = hBmp
      .hPal = hPal
   End With
   r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
   Set CreateBitmapPicture = IPic
End Function
  Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim r As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE
   If Client Then
      hDCSrc = GetDC(hWndSrc)
   Else
      hDCSrc = GetWindowDC(hWndSrc)
   End If
   hDCMemory = CreateCompatibleDC(hDCSrc)
   hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
   hBmpPrev = SelectObject(hDCMemory, hBmp)
   RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
   HasPaletteScrn = RasterCapsScrn And RC_PALETTE
   PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      LogPal.palVersion = &H300
      LogPal.palNumEntries = 256
      r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
      hPal = CreatePalette(LogPal)
      hPalPrev = SelectPalette(hDCMemory, hPal, 0)
      r = RealizePalette(hDCMemory)
   End If
   r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
   hBmp = SelectObject(hDCMemory, hBmpPrev)
   If HasPaletteScrn And (PaletteSizeScrn = 256) Then
      hPal = SelectPalette(hDCMemory, hPalPrev, 0)
   End If
   r = DeleteDC(hDCMemory)
   r = ReleaseDC(hWndSrc, hDCSrc)
   Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CaptureScreen() As Picture
  Dim hWndScreen As Long
   hWndScreen = GetDesktopWindow()
   Set CaptureScreen = CaptureWindow(hWndScreen, False, Label3.Caption, Label4.Caption, Label1.Caption, Label2.Caption)
End Function


Private Sub Form_Load()
Me.Hide
If Screen.Width = 12000 Then
Label1.Caption = 800
Label2.Caption = 600
Else
Label1.Caption = 1023
Label2.Caption = 767
End If
Label9.Caption = "10"
FramePerSeconds32 = 100
End Sub

Private Sub Form_Resize()
If Label8.Tag = "1" Then
MsgBox "You Have " + Format(PictureNumbers32) + " Pictures !"
Form3.Show
Timer1.Enabled = False
Timer2.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Label8.Caption = "ON"
Me.Hide
End If
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label10.Visible = False
End Sub

Private Sub Label12_Click()
End
End Sub

Private Sub Label9_Change()
FramePerSeconds32 = 1000 / Int(Label9.Caption)
End Sub

Private Sub Picture2_KeyPress(KeyAscii As Integer)
Picture2.Visible = False
Pi.Visible = True
Label4.Caption = L.TOp
Label3.Caption = L.Left
Label1.Caption = L.Width
Label2.Caption = L.Height
Me.WindowState = 0
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo 5
L.Visible = False
L.Left = x
L.TOp = y
LL.Left = x
LL.TOp = y
LL.Width = 0
LL.Height = 0
LDD.Caption = "ON"
5:
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo 5
If LDD.Caption = "ON" Then
LL.Width = x - LL.Left
LL.Height = y - LL.TOp
End If
5:
Label10.Visible = True
Label10.Caption = Str(LL.Left) + "*" + Str(LL.TOp) + "*" + Str(LL.Width) + "*" + Str(LL.Height)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo 5
LDD.Caption = "OFF"
L.Width = x - L.Left
L.Height = y - L.TOp
L.Visible = True
5:
End Sub

Private Sub PowerButton1_Click()
Label9.Caption = Val(Label9.Caption) + 1
End Sub

Private Sub PowerButton1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer4.Enabled = True
End Sub

Private Sub PowerButton1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer4.Enabled = False
Timer4.Interval = 250
End Sub

Private Sub PowerButton2_Click()
If Val(Label9.Caption) > 1 Then
Label9.Caption = Val(Label9.Caption) - 1
End If
End Sub

Private Sub PowerButton2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer5.Enabled = True
End Sub

Private Sub PowerButton2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Timer5.Enabled = False
Timer5.Interval = 250
End Sub

Private Sub PowerButton3_Click()
On Error Resume Next
File1.Path = App.Path + "\Temp\"
Dim i As Integer
For i = 0 To File1.ListCount
Kill File1.Path + "\" + File1.List(i)
Next
MsgBox "SaveScreen Will Be Started After 5 Sec."
FramePerSeconds32 = 1000 / Int(Label9.Caption)
Timer1.Interval = 5000
Me.WindowState = 1
Label8.Tag = "1"
End Sub

Private Sub PowerButton4_Click()
Me.Hide
Form3.Show
End Sub

Private Sub Timer1_Timer()
If Label8.Caption = "ON" Then
Timer1.Interval = FramePerSeconds32
PictureNumbers32 = PictureNumbers32 + 1
Set Picture1.Picture = CaptureScreen()
Dim XN As String
XN = Format(PictureNumbers32)
If Len(XN) = 1 Then XN = "000" + XN
If Len(XN) = 2 Then XN = "00" + XN
If Len(XN) = 3 Then XN = "0" + XN
SavePicture Picture1.Picture, App.Path + "\Temp\" + XN + ".BMP"
Else
Timer1.Interval = 0
End If
End Sub

Private Sub Timer2_Timer()
Set Picture2.Picture = CaptureScreen()
Me.Show
Picture2.Width = Me.Width
Picture2.Height = Me.Height
Me.WindowState = 2
L.Width = Label1.Caption
L.Height = Label2.Caption
LL.Width = Label1.Caption
LL.Height = Label2.Caption
Timer2.Interval = 0
End Sub

Private Sub Timer4_Timer()
Timer4.Interval = 100
Label9.Caption = Val(Label9.Caption) + 1
End Sub

Private Sub Timer5_Timer()
If Val(Label9.Caption) > 1 Then
Timer5.Interval = 100
Label9.Caption = Val(Label9.Caption) - 1
End If
End Sub
