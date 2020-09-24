VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8205
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   4620
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image PowerButton3 
      Height          =   465
      Left            =   120
      Picture         =   "Form3.frx":2A4A
      Top             =   3960
      Width           =   2400
   End
   Begin VB.Image PowerButton1 
      Height          =   465
      Left            =   2520
      Picture         =   "Form3.frx":64AC
      Top             =   3960
      Width           =   2385
   End
   Begin VB.Image PowerButton2 
      Height          =   465
      Left            =   5520
      Picture         =   "Form3.frx":9F0E
      Top             =   3960
      Width           =   2520
   End
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   4080
      Picture         =   "Form3.frx":DC58
      Stretch         =   -1  'True
      Top             =   495
      Width           =   4005
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
      Left            =   7920
      TabIndex        =   0
      Top             =   0
      Width           =   255
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
      TabIndex        =   1
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub File1_Click()
On Error Resume Next
Image1.Picture = LoadPicture(File1.Path + "\" + File1.List(File1.ListIndex))
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\Temp\"
End Sub

Private Sub Label12_Click()
End
End Sub

Private Sub PowerButton1_Click()
Dim i As Integer
For i = 0 To File1.ListCount - 1
On Error Resume Next
Kill File1.Path + "\" + File1.List(i)
Next
File1.Refresh
End Sub

Private Sub PowerButton2_Click()
If File1.ListCount > 0 Then
Form4.Show
Form4.txtFrameDuration.Text = Format(FramePerSeconds32)
AllFiles = File1.ListCount
Unload Me
Else
MsgBox "You don't have any picture to export as AVI file !", vbOKOnly + vbCritical, "Error"
End If
End Sub

Private Sub PowerButton3_Click()
Dim i As Integer
For i = 0 To File1.ListCount - 1
On Error Resume Next
If File1.Selected(i) = True Then
Kill File1.Path + "\" + File1.List(i)
End If
Next
File1.Refresh
End Sub
