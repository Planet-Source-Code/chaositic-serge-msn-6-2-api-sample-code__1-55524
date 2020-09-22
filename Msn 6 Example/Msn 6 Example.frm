VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Msn 6 Example"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command11 
      Caption         =   "View Profile"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Open IM"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Signout"
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Out to Lunch"
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "On The Phone"
      Height          =   195
      Left            =   2880
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Busy"
      Height          =   195
      Left            =   1920
      TabIndex        =   7
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Away"
      Height          =   195
      Left            =   2880
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Brb"
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Offline"
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Online"
      Height          =   195
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4080
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Name"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Email Here"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Email here"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Name Here"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents msn As Messenger
Attribute msn.VB_VarHelpID = -1
Private Sub Command1_Click()
msn.OptionsPages 0, MOPT_GENERAL_PAGE
SendKeys Text1.Text
Sleep (1000)
SendKeys "{ENTER}"
End Sub

Private Sub Command10_Click()
msn.InstantMessage Text2.Text
End Sub

Private Sub Command11_Click()
msn.ViewProfile Text3.Text
End Sub

Private Sub Command2_Click()
msn.MyStatus = MISTATUS_ONLINE
End Sub

Private Sub Command3_Click()
msn.MyStatus = MISTATUS_INVISIBLE
End Sub

Private Sub Command4_Click()
msn.MyStatus = MISTATUS_BE_RIGHT_BACK
End Sub

Private Sub Command5_Click()
msn.MyStatus = MISTATUS_AWAY
End Sub

Private Sub Command6_Click()
msn.MyStatus = MISTATUS_BUSY
End Sub

Private Sub Command7_Click()
msn.MyStatus = MISTATUS_ON_THE_PHONE
End Sub

Private Sub Command8_Click()
msn.MyStatus = MISTATUS_OUT_TO_LUNCH
End Sub

Private Sub Command9_Click()
msn.Signout
End Sub

Private Sub Form_Load()
Set msn = New Messenger
End Sub







Private Sub Timer1_Timer()
If msn.MyStatus = MISTATUS_OFFLINE Then
msn.AutoSignin
End If
End Sub
