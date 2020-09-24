VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SaveScreen 4"
   ClientHeight    =   4575
   ClientLeft      =   3105
   ClientTop       =   3345
   ClientWidth     =   7320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAVICreator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Compress Video"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtAVIFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   6
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton cmdPick 
      Caption         =   "..."
      Height          =   315
      Left            =   6840
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox txtFrameDuration 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2820
      TabIndex        =   4
      Text            =   "100"
      Top             =   1500
      Width           =   3975
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Export.."
      Height          =   555
      Left            =   5160
      TabIndex        =   3
      Top             =   3720
      Width           =   1755
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4380
      Left            =   80
      ScaleHeight     =   4380
      ScaleWidth      =   1500
      TabIndex        =   0
      Top             =   80
      Width           =   1500
      Begin VB.Image Image4 
         Height          =   1680
         Left            =   -120
         Picture         =   "frmAVICreator.frx":2A4A
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1680
      End
      Begin VB.Image Image3 
         Height          =   1800
         Left            =   -240
         Picture         =   "frmAVICreator.frx":EA8C
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   1680
      End
   End
   Begin VB.Label LoadingLbl 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ready"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "File name :"
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   9
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   " Frame (ms) :"
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   8
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "You Have To Select File Name And Frame (ms) Video ."
      Height          =   255
      Index           =   12
      Left            =   1860
      TabIndex        =   7
      Top             =   600
      Width           =   5235
   End
   Begin VB.Image Image8 
      Height          =   1920
      Left            =   1800
      Picture         =   "frmAVICreator.frx":1AACE
      Top             =   2520
      Visible         =   0   'False
      Width           =   1920
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
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   " SaveScreen V.4 - AVI Creator"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cAVI As New ZaidMarkabi_Avi
Private m_cVH As New cVideoHandlers

Private Function pbGetImage(ByVal iImageIndex As Long) As cBmp
Static sDir As String
Static sBaseDir As String
Static cellWidth As Long
Static cellHeight As Long
Static xCell As Long
Static yCell As Long

   If (iImageIndex = 1) Then

         sBaseDir = App.Path + "\Temp\"
         If (Right(sBaseDir, 1) <> "\") Then sBaseDir = sBaseDir & "\"
         sDir = Dir(sBaseDir & "*.bmp")
   
   Else

         sDir = Dir
   
   End If

   Dim cb As New cBmp
      If Len(sDir) > 0 Then
         cb.Load sBaseDir & sDir
         Set pbGetImage = cb
      End If
End Function

Private Sub cmdCreate_Click()
m_cAVI.Filename = txtAVIFile.Text
m_cAVI.Name = txtAVIFile.Text
m_cAVI.FrameDuration = CLng(txtFrameDuration.Text)
m_cAVI.bitsPerPixel = 24
If Check1.Value = 1 Then m_cAVI.VideoHandlerFourCC = m_cVH.Handler(1).FourCC
FileCopy App.Path + "\DAT.Zaid", App.Path + "\Temp\" + "9999999999.BMP"

Dim cb As cBmp
Dim iImageIndex As Long
Dim bStreamOpen As Boolean

iImageIndex = 1
 Set cb = pbGetImage(iImageIndex)
   If Not (cb Is Nothing) Then
      m_cAVI.StreamCreate cb
      bStreamOpen = True
      Me.Enabled = False
      Do
      If iImageIndex = AllFiles Then
      LoadingLbl.Caption = "Complete"
      Else
      LoadingLbl.Caption = Format(iImageIndex + 1) + " of " + Format(AllFiles)
      End If
      DoEvents
         iImageIndex = iImageIndex + 1
         Set cb = pbGetImage(iImageIndex)
         If Not (cb Is Nothing) Then
            m_cAVI.StreamAdd cb
         End If
      Loop While Not (cb Is Nothing)
      Me.Enabled = True
      m_cAVI.StreamClose
      bStreamOpen = False
   End If
   
   MsgBox "AVI created successfully.", vbInformation
   End
End Sub

Private Sub cmdPick_Click()
On Error GoTo 1
CommonDialog1.ShowSave
txtAVIFile.Text = CommonDialog1.Filename + ".Avi"
1:
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Label12_Click()
End
End Sub
