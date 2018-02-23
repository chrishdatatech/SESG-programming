VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Loop Until"
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   6360
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Loop While"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   5400
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Do Until"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do While"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   3480
      Width           =   3015
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
      Left            =   600
      TabIndex        =   7
      Top             =   2760
      Width           =   2955
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
      Width           =   2955
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

Private Sub Command2_Click()

i = 0
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

Private Sub Form_Load()

End Sub
