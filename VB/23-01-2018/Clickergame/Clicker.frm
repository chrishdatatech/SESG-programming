VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   3720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Text            =   "NAME"
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Highscore"
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enter"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Text            =   "Enter seconds"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   855
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   4440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click me"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Time left"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Number of clicks"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public klikCount As Integer
Public name As String
Public time As Integer
Public score As Integer

Private Sub Command1_Click()
    Timer1.Enabled = True
    klikCount = klikCount + 1
    Label2.Caption = klikCount
End Sub

Private Sub Command2_Click() 'Overfører tekst fra text1
    Label1.Caption = Text1.Text
        
End Sub

Private Sub Command3_Click() 'Sikrer at label1-2 er 0 ved tryk på Reset, samt command1 enables.
    Label1.Caption = 0
    Label2.Caption = 0
    Label1.Caption = Text1.Text
    klikCount = 0
    Command1.Enabled = True
    Timer1.Enabled = False
    
End Sub

Private Sub Command4_Click()
    Load Form2
    Form2.Show
End Sub

Private Sub Text1_GotFocus()
    Text1.Text = "" 'Rydder text1 on click.
End Sub

Private Sub Text2_GotFocus()
    Text2.Text = ""

End Sub

Private Sub Text2_Change()
        Text2.Enabled = True
        If Text2.Text = "" Then
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Text1.Enabled = False
        Else
            Command1.Enabled = True
            Command2.Enabled = True
             Command3.Enabled = True
            Text1.Enabled = True
        End If
            
            
End Sub

Private Sub Timer1_Timer()  'Timer der bestemmer om command1 er enabled eller ej
    Label1.Caption = Label1.Caption - 1
        If Label1.Caption = 0 Then
            
            
            Load Form2
            Form2.Show
            Timer1.Enabled = False
            Command1.Enabled = False
        End If
        
        
End Sub

