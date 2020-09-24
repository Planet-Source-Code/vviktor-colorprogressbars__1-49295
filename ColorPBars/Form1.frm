VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   5670
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   3000
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2400
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Color progressbars:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Standard progressbar:"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Any) As Long
Private Const WM_USER = &H400
Private Const CCM_FIRST       As Long = &H2000&
Private Const PBM_SETBARCOLOR As Long = (WM_USER + 9)
Private Const CCM_SETBKCOLOR  As Long = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR  As Long = CCM_SETBKCOLOR

Private Sub Form_Load()

'- Red progressbar
SendMessage ProgressBar1(1).hwnd, PBM_SETBARCOLOR, 0, vbRed
SendMessage ProgressBar1(1).hwnd, PBM_SETBKCOLOR, 0, RGB(200, 0, 0)

'- Green progressbar
SendMessage ProgressBar1(2).hwnd, PBM_SETBARCOLOR, 0, vbGreen
SendMessage ProgressBar1(2).hwnd, PBM_SETBKCOLOR, 0, RGB(0, 200, 0)

'- Black-white progressbar
SendMessage ProgressBar1(3).hwnd, PBM_SETBARCOLOR, 0, vbWhite
SendMessage ProgressBar1(3).hwnd, PBM_SETBKCOLOR, 0, vbBlack




End Sub

Private Sub Timer1_Timer()

'The animation
For I = 0 To ProgressBar1.Count - 1
If ProgressBar1(I).Value = ProgressBar1(I).Max Then
ProgressBar1(I).Value = 0
Else
ProgressBar1(I).Value = ProgressBar1(I).Value + 1
End If
Next I
End Sub
