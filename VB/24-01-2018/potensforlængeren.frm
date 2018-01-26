VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   525
      Left            =   600
      TabIndex        =   4
      Text            =   "n"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   1320
      TabIndex        =   2
      Text            =   "n"
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Beregn"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2400
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Omregn n i potens til decimaltal"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Long
Dim b As Long

Private Sub Command1_Click()
    a = Val(Text3.Text)
    b = Val(Text1.Text)
    Sum = (a ^ b)
    Label3.Caption = Sum
    
End Sub

Private Sub Text1_GotFocus()
    Text1.Text = ""
End Sub

Private Sub Text3_GotFocus()
    Text3.Text = ""
End Sub




