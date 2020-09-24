VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Open GIF Image"
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save GIF Frame"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   240
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Index           =   4
      Left            =   1440
      TabIndex        =   9
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "GIF SPLITTER,AUTHOR:VANJA FUCKAR,EMAIL:INGA@VIP.HR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const sFilter = "Gif (*.gif)" & vbNullChar & "*.gif"
Dim sFile As String
Dim sPath As String
Private HH As New Splitter
Dim cnt As Long
Dim Playcnt As Long
Private Sub Command1_Click()
If cnt = 0 Then Exit Sub
cnt = cnt - 1
Set Image1.Picture = HH.GetFrame(cnt)
ShowCord
End Sub
Private Sub Command2_Click()
If cnt = HH.GetFramesCount - 1 Or HH.GetFramesCount = 0 Then Exit Sub
cnt = cnt + 1
Set Image1.Picture = HH.GetFrame(cnt)
ShowCord
End Sub
Private Sub Command3_Click()
aa = GetOpenFilePath(hWnd, sFilter, 0, sFile, "", "Open GIF Image", sPath)
If aa = False Then Exit Sub
gg = HH.LoadGif(sPath)
If gg = False Then MsgBox "Single Frame GIF!", vbInformation, "Information"
cnt = 0
Set Image1.Picture = HH.GetFrame(cnt)
Label1(2) = "Frames Count:" & HH.GetFramesCount
ShowCord
End Sub

Private Sub Command4_Click()
aa = GetSaveFilePath(hWnd, sFilter, 0, sFilter, "", "", "Save Database", sPath)
If aa = False Then Exit Sub
HH.SaveFrame sPath, cnt
End Sub

Private Sub ShowCord()
Label1(0) = "X:" & HH.GetFrameDimenzion(cnt).x
Label1(1) = "Y:" & HH.GetFrameDimenzion(cnt).y
Label1(3) = "Frame:" & cnt
Label1(4) = "Frame Wait:" & HH.FrameWait(cnt)
End Sub




