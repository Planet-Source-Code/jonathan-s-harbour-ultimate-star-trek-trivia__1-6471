VERSION 5.00
Begin VB.Form frmConfirm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3105
   ClientLeft      =   -15
   ClientTop       =   -15
   ClientWidth     =   5340
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmConfirm.frx":0000
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton3 
      Caption         =   "Button3"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1935
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdButton2 
      Caption         =   "Button2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3195
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdButton1 
      Caption         =   "Button1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   675
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2100
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblTitle1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TITLE 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   5100
   End
   Begin VB.Label lblTitle3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TITLE 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5070
   End
   Begin VB.Label lblTitle2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TITLE 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   1035
      Width           =   5070
   End
End
Attribute VB_Name = "frmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim reply As Boolean
Dim status As Boolean

Public Function Finished() As Boolean
    Finished = status
End Function

Private Sub cmdButton1_Click()
        Button_Off
        frmConfirm.Hide
        reply = True
End Sub

Private Sub cmdButton2_Click()
        Button_Off
        frmConfirm.Hide
        reply = False
End Sub

Private Sub cmdButton3_Click()
    Button_Off
    frmConfirm.Hide
    reply = True
    status = True
End Sub

Public Function Confirm(ByVal title1$, ByVal title2$, ByVal title3$, button1$, button2$, ByVal MODAL As Boolean) As Boolean
    status = False
    frmConfirm.ScaleMode = vbPixels
    lblTitle1.Caption = title1$
    lblTitle2.Caption = title2$
    lblTitle3.Caption = title3$
    
    If button1$ = "" And button2$ = "" Then
        cmdButton1.Visible = False
        cmdButton2.Visible = False
        cmdButton3.Visible = True
        cmdButton3.Caption = "&Okay"
    ElseIf button1$ <> "" And button2$ <> "" Then
        cmdButton1.Visible = True
        cmdButton1.Caption = button1$
        cmdButton2.Visible = True
        cmdButton2.Caption = button2$
        cmdButton3.Visible = False
    ElseIf button1$ = "" And button2$ <> "" Then
        cmdButton1.Visible = False
        cmdButton2.Visible = False
        cmdButton3.Visible = True
        cmdButton3.Caption = button2$
    ElseIf button1$ <> "" And button2$ = "" Then
        cmdButton2.Visible = False
        cmdButton1.Visible = False
        cmdButton3.Visible = True
        cmdButton3.Caption = button1$
    End If
    
    If MODAL Then
        frmConfirm.Show vbModal
    Else
        frmConfirm.Show
    End If
    Confirm = reply
    status = True
End Function

Public Function DisplayNextPlayer(ByVal title1$, ByVal title2$, ByVal title3$, button1$, button2$) As Boolean
    status = False
    frmConfirm.ScaleMode = vbPixels
    lblTitle1.Caption = title1$
    lblTitle2.Caption = title2$
    lblTitle3.Caption = title3$
        
    cmdButton1.Visible = False
    cmdButton2.Visible = False
    cmdButton3.Visible = False
        
    frmConfirm.Show
    frmConfirm.Refresh
    
    'Speak "NEXT PLAYER"
    Speak "CATEGORY"
    Speak title3$
    
    reply = True
    status = True
    frmConfirm.Hide
    DisplayNextPlayer = reply
End Function

Public Function GainPoints(ByVal points&, ByVal msg1$, ByVal msg2$) As Boolean
    Dim msg$
    
    If points > 1 Then
        msg$ = "YOU GAINED " & points & " POINTS"
    Else
        msg$ = "YOU GAINED " & points & " POINT"
    End If
    
    status = False
    frmConfirm.ScaleMode = vbPixels
    lblTitle1.Caption = msg$
    lblTitle2.Caption = msg1$
    lblTitle3.Caption = msg2$
        
    cmdButton1.Visible = False
    cmdButton2.Visible = False
    cmdButton3.Visible = False
        
    frmConfirm.Show
    frmConfirm.Refresh
        
    Speak "YOU GAINED"
    SpeakNumber points
    Speak "POINTS"
    
    reply = True
    status = True
    frmConfirm.Hide
    GainPoints = reply
End Function

Public Sub SpeakNumber(ByVal number&)
    Dim num&
    num = number
       
    If num = 1000 Then
        Play_Sound "1.wav"
        Play_Sound "1000.wav"
        num = 0
    End If
    If num > 1000 Then
        Play_Sound Int(num \ 1000) & ".wav"
        Play_Sound "1000.wav"
        num = num - Int(num \ 1000) * 1000
    End If
    If num = 100 Then
        Play_Sound "1.wav"
    End If
    If num > 100 Then
        Play_Sound Int(num \ 100) & ".wav"
        Play_Sound "100.wav"
        num = num - Int(num \ 100) * 100
    End If
    If num > 20 Then
        Play_Sound Int(num \ 10) * 10 & ".wav"
        num = num - Int(num \ 10) * 10
    End If
    If num > 0 Then Play_Sound num & ".wav"
End Sub

Private Sub Speak(ByVal msg$)
    Select Case Trim$(UCase$(msg$))
        Case "YOU GAINED":          Play_Sound "yougained.wav"
        Case "POINTS":              Play_Sound "points.wav"
        Case "POINT":               Play_Sound "point.wav"
        Case "NEXT PLAYER":         Play_Sound "nextplayer.wav"
        Case "CATEGORY":            Play_Sound "category.wav"
        Case "CHARACTERS":          Play_Sound "characters.wav"
        Case "CITIES":              Play_Sound "cities.wav"
        Case "CODES/NUMBERS":       Play_Sound "codes.wav"
        Case "COLORS":              Play_Sound "colors.wav"
        Case "DATES":               Play_Sound "dates.wav"
        Case "EPISODES/SHOWS":      Play_Sound "episodes.wav"
        Case "FIRST NAMES":         Play_Sound "firstnames.wav"
        Case "FOOD":                Play_Sound "food.wav"
        Case "LAST NAMES":          Play_Sound "lastnames.wav"
        Case "LETTERS":             Play_Sound "letters.wav"
        Case "MIDDLE NAMES":        Play_Sound "middlenames.wav"
        Case "MISCELLANEOUS":       Play_Sound "misc.wav"
        Case "NICKNAMES":           Play_Sound "nicknames.wav"
        Case "NUMBERS":             Play_Sound "numbers.wav"
        Case "PLANETS/WORLDS":      Play_Sound "planets.wav"
        Case "RANKS":               Play_Sound "ranks.wav"
        Case "REAL NAMES":          Play_Sound "realnames.wav"
        Case "SPECIES/RACES":       Play_Sound "species.wav"
        Case "SEASONS":             Play_Sound "seasons.wav"
        Case "STARDATES":           Play_Sound "stardates.wav"
        Case "TECHNOLOGY":          Play_Sound "technology.wav"
        Case "TRUE OR FALSE":       Play_Sound "trueorfalse.wav"
        Case "VESSELS/VEHICLES":    Play_Sound "vessels.wav"
        Case "YEARS":               Play_Sound "years.wav"
        Case "YES OR NO":           Play_Sound "yesorno.wav"
    End Select
End Sub

Private Sub Form_Load()
    Me.MouseIcon = frmMain.MouseIcon
    status = False
End Sub
