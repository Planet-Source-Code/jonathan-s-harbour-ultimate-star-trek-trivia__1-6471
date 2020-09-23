VERSION 5.00
Begin VB.Form frmIntro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3825
   ClientLeft      =   195
   ClientTop       =   30
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Intro.frx":0000
   ScaleHeight     =   255
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   423
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H8000000A&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3300
      Picture         =   "Intro.frx":4F844
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2700
      Width           =   2295
   End
   Begin VB.CommandButton cmdContinue 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   660
      Picture         =   "Intro.frx":5075A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2700
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1999 by Jonathan S. Harbour"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2220
      Width           =   4995
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Start_Intro()
    frmMidi.Play 1
    Me.Refresh
    Play_Sound "SOUND4.WAV"
    Me.MouseIcon = frmMain.MouseIcon
End Sub

Private Sub Form_Load()
    Randomize
    Start_Intro
End Sub

Private Sub cmdContinue_Click()
    Button_Off
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Button_Off
    End
End Sub

