VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbInterval 
      Height          =   315
      ItemData        =   "frmArielTimerTest.frx":0000
      Left            =   1320
      List            =   "frmArielTimerTest.frx":002B
      TabIndex        =   2
      Top             =   780
      Width           =   1395
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   180
      Width           =   1095
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Tick Counter"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   360
      Width           =   915
   End
   Begin VB.Label lblTimer 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "Interval (ms)"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   840
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents ArTimer As ArielTimer
Attribute ArTimer.VB_VarHelpID = -1

Private Sub ArTimer_OnTimer()
lblTimer = Format(Timer, "#,###.00")

End Sub


Private Sub cmbInterval_Change()
ArTimer.Interval = cmbInterval.Text

End Sub

Private Sub cmbInterval_Click()
ArTimer.Interval = cmbInterval.Text

End Sub


Private Sub cmdStart_Click()
ArTimer.Enabled = True

End Sub

Private Sub cmdStop_Click()
ArTimer.Enabled = False
End Sub


Private Sub Form_Load()
Set ArTimer = New ArielTimer
ArTimer.Interval = 2000
cmbInterval.Text = ArTimer.Interval

End Sub


