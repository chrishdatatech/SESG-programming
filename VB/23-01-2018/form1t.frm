VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   2460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Hilsen"
      Enabled         =   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Tjek"
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   525
      Left            =   1320
      TabIndex        =   7
      Text            =   "Speed"
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change speed"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop for helvede"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   1320
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Dakkedak"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Swipp"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hansen"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Text            =   "Test"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hvem er du?"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Chris"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = "Hansen"
End Sub

Private Sub Command2_Click()
Label1.Caption = Text1.Text
End Sub

Private Sub Command3_Click()
Timer1.Enabled = True
End Sub

Private Sub Command4_Click()
Timer2.Enabled = False
Timer2.Enabled = False
Label1.Visible = True
End Sub

Private Sub Command5_Click()
Timer1.Interval = Text2.Text
Timer2.Interval = Text2.Text
End Sub

Private Sub Command6_Click()

If Text3.Text = "1234" Then
    Label1.Caption = "Rigtigt"
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command7.Enabled = True
    Text1.Enabled = True
    Form1.BackColor = &HC000&
Else
    Label1.Caption = "Fejl"
    Form1.BackColor = 255
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command7.Enabled = False
    Text1.Enabled = False
End If

End Sub

Private Sub Command7_Click()

Select Case Text1.Text
    Case "Chris"
        Label2.Caption = "Hej Chris"
    Case "Leif"
        Label2.Caption = "Hej Leif"
    Case "Jens"
        Label2.Caption = "Hej Jens"
    Case "Per"
        Label2.Caption = "Hej Per"
    Case Else
        Label2.Caption = "Hej, dig kender jeg ikke"
End Select
    
    
    
End Sub

Private Sub Timer1_Timer()
Label1.Visible = False
Timer2.Enabled = True
Label1.BackColor = &H8000000D
Label1.ForeColor = 255
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label1.Visible = True
Timer2.Enabled = False
Timer1.Enabled = True
End Sub
