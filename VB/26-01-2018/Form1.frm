VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11535
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Se facit"
      Height          =   495
      Left            =   480
      TabIndex        =   25
      Top             =   2040
      Width           =   855
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9360
      Top             =   1440
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   7920
      Top             =   1560
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   4320
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   3240
      TabIndex        =   11
      Text            =   "C:/Users/cdh66/Desktop/Programmering/VB/"
      Top             =   6120
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hent Spil"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   6960
      Width           =   4455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OK"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Game Info"
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Gem resultat i en fil"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Passer Tallet"
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "End"
      Height          =   375
      Left            =   9480
      TabIndex        =   1
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nyt Tal (PC)"
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   6360
      Picture         =   "Form1.frx":0000
      Top             =   2160
      Visible         =   0   'False
      Width           =   4995
   End
   Begin VB.Label Label12 
      Caption         =   "Indtast password:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label11 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   23
      Top             =   2760
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   22
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Indtast filnavn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   21
      Top             =   6120
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "Indtast dit navn:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Antal spil:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   17
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Antal gæt:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   16
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   15
      Top             =   5400
      Width           =   11055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Gæt et tal mellem 1 og 100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   14
      Top             =   120
      Width           =   10095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Programmør: Chris D. Hansen"
      Height          =   255
      Left            =   9240
      TabIndex        =   13
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   8880
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   855
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tal As Integer
Public gætialt As Integer
Public antalspil As Integer
Dim gættal As Integer
Dim gæt As Integer
Dim pctal As Integer

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Check1.Visible = False
        Label12.Visible = True
        Text3.Visible = True
        Command6.Visible = True
        Text3.SetFocus
    Else
        Text1.Visible = False
    End If
End Sub

Private Sub Command1_Click()
    Randomize Timer
    pctal = Int(Rnd * 100) + 1
    Text1.Text = pctal
    
    Command1.Enabled = False
    Command3.Enabled = True
    HScroll1.Enabled = True
    Text2.Enabled = True
    
    Text2.Text = ""
    Label3.Caption = ""
    antalspil = antalspil + 1
    Label10.Caption = antalspil
    gæt = 0
    Label11.Caption = 0
    Image1.Visible = False
End Sub

Private Sub Command2_Click()
Dim valg As Integer

    valg = MsgBox("Afslutter programmet :)", vbOKCancel + vbExclamation, "Gæt et tal mellem 1 og 100")
    If valg = vbOK Then
    valg = MsgBox("Er du sikker på du vil afslutte programmet?", vbYesNo + vbInformation, "Gæt et tal mellem 1 og 100")
    If valg = vbNo Then
        GoTo 200
    Else
        End
    End If
Else
End If
200

End Sub

Private Sub Command3_Click()
    gættal = Val(Text2.Text)
    
    If gættal = 0 Or gættal > 100 Then
        MsgBox "Du skal indtaste et tal mellem 1 og 100"
        Text2.Text = ""
    Else
        If gættal > pctal + 5 Then
            Label3.Caption = "Du gættede alt for højt"
    Else
        If gættal > pctal Then
        Label3.Caption = "Nu er du tæt på, men stadig for højt"
        End If
    End If
    
    If gættal = pctal Then
        Label3.Caption = "Tillykke " + Text4.Text + ", du gættede rigtigt"
        Command1.Enabled = True
        Command3.Enabled = False
        Image1.Visible = True
        Text2.Enabled = False
    End If
    
    If gættal < pctal - 5 Then
        Label3.Caption = "Du gættede for lavt"
    Else
        If gættal < pctal Then
        Label3.Caption = "Nu er du tæt på, men stadig for lavt"
        End If
    End If
    
        gæt = gæt + 1
        gætialt = gætialt + 1
        Label11.Caption = gæt
    End If
End Sub

Private Sub Command4_Click()
    Open Text5.Text For Output As #1
        Print #1, Text4.Text; ": "; "Antal gæt i alt"; " "; gætialt; " "; "antal Spil"; " "; Label10.Caption
        Print #1, gætialt
        Print #1, antalspil
        Print #1, Text4.Text
    Close
End Sub

Private Sub Command5_Click()
    Load Form2
    Form2.Show
    Form2.Label2.Caption = gætialt
    Form2.Label4.Caption = antalspil
    
    If Label10.Caption > 0 Then
        Form2.Label7.Caption = gætialt / antalspil
        tal = gætialt / antalspil
    End If
    
    'skriver tekst på form2
    
    Select Case tal
        Case 0
            Form2.Label8.Caption = "Du har ikke gættet tallet endnu!"
        Case 1 To 2
            Form2.Label8.Caption = "Super ekspert. Har du gættet passwordet?"
        Case 3, 4
            Form2.Label8.Caption = "Du er ekspert!"
        Case 5, 6
            Form2.Label8.Caption = "OK, men det kunne være bedre!"
        Case Else
            Form2.Label8.Caption = "Dårligt! - Prøv med halveringsmetoden :)"
    End Select
End Sub

Private Sub Command6_Click()
    If Text3.Text = "chris" Then
        Text1.Visible = True
        Label12.Visible = False
        Text3.Visible = False
        Command6.Visible = False
        Check1.Visible = True
    Else
        Text1.Visible = False
        Label12.Visible = False
        Text3.Visible = False
        Command6.Visible = False
        Check1.Visible = True
        Check1.Value = 0
    End If
End Sub

Private Sub Command7_Click()
    Open Text5.Text For Input As #1
        Line Input #1, inputdata 'Read line of data (1. linie skal ikke bruges)
        Line Input #1, inputdata 'Read line of data (antal gæt i alt)
        gætialt = inputdata
        Line Input #1, inputdata 'Read line of data (Antal spil)
        antalspil = inputdata
        Line Input #1, inputdata 'Read line of data (Navnet på spiller)
        Text4.Text = inputdata
        Label10.Caption = antalspil
        Label11.Caption = gætialt
    Close
End Sub

Private Sub HScroll1_Scroll()
    Text2.Text = HScroll1.Value
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call Command3_Click
    End If
End Sub

Private Sub Timer1_Timer()
    Label6.Visible = False
    Label7.Visible = True
    Label7.Caption = Int(Rnd * 100)
    Shape1.Visible = False
    Shape2.Visible = True
    Timer1.Enabled = False
    Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
    Label7.Visible = False
    Label6.Visible = True
    Label6.Caption = Int(Rnd * 100)
    Shape2.Visible = False
    Shape1.Visible = True
    Timer2.Enabled = False
    Timer1.Enabled = True
End Sub

