VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   5385
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   735
      Left            =   840
      TabIndex        =   0
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HIGHSCORE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
 Unload Form2
 Load Form1
 Form1.Show
End Sub

