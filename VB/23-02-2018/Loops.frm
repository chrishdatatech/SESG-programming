VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Do While"
      Height          =   855
      Left            =   3480
      TabIndex        =   14
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Do Until"
      Height          =   855
      Left            =   3480
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Loop While"
      Height          =   855
      Left            =   3480
      TabIndex        =   12
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Loop Until"
      Height          =   855
      Left            =   3480
      TabIndex        =   11
      Top             =   6480
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   5400
      TabIndex        =   10
      Text            =   "Max"
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   5400
      TabIndex        =   9
      Text            =   "Step"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Step"
      Height          =   855
      Left            =   3720
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Loop Until"
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Loop While"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   5520
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Do Until"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do While"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "For Next"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Do...Loops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1080
      TabIndex        =   7
      Top             =   2880
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "For...Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   6
      Top             =   1080
      Width           =   5715
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   360
      Y1              =   120
      Y2              =   7320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Loops"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Command1_Click()
    For i = 1 To 10 Step 1
        Print i
    Next i
End Sub

Private Sub Command10_Click()
i = 10
    Do While i >= 1
    Print i
    i = i - 1
Loop
End Sub

Private Sub Command2_Click()

i = 1
    Do While i <= 10
    Print i
    i = i + 1
Loop
End Sub

Private Sub Command3_Click()
i = 0
    Do Until i = 10
    Print i
    i = i + 1
Loop
End Sub

Private Sub Command4_Click()
i = 0
Do
    Print i
    i = i + 1
    Loop While i <= 10
End Sub

Private Sub Command5_Click()
i = 0

Do
    Print i
    i = i + 1
    Loop Until i = 10
End Sub

Private Sub Command6_Click()
    For i = 1 To Text2.Text Step Text1.Text
        Print i
Next i

End Sub

Private Sub Command7_Click()
i = 10

Do
    Print i
    i = i - 1
    Loop Until i = 0
End Sub

Private Sub Command8_Click()
i = 10
Do
    Print i
    i = i - 1
    Loop While i >= 0
End Sub

Private Sub Command9_Click()
i = 10
    Do Until i = 0
    Print i
    i = i - 1
Loop
End Sub

Private Sub Text1_GotFocus()
    Text1.Text = ""
End Sub

Private Sub Text2_GotFocus()
    Text2.Text = ""
End Sub

Private Sub Text3_GotFocus()
    Text3.Text = ""
End Sub

Private Sub Text4_GotFocus()
    Text4.Text = ""
End Sub
