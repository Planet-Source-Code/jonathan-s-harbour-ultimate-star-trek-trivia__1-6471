VERSION 5.00
Begin VB.Form frmNumPlayers 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Players.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   225
      Picture         =   "Players.frx":2076
      ScaleHeight     =   1395
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   1170
      Width           =   3555
      Begin VB.OptionButton optPlayers 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   450
         Width           =   690
      End
      Begin VB.OptionButton optPlayers 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   450
         Width           =   690
      End
      Begin VB.OptionButton optPlayers 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   690
      End
      Begin VB.OptionButton optPlayers 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   450
         Value           =   -1  'True
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2220
      Width           =   1200
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   3915
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "How many players?"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   4995
   End
End
Attribute VB_Name = "frmNumPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim numPlayers&

Private Sub cmdPlay_Click()
    Button_Off
    Game.CurrentPlayer = 0
    Game.numPlayers = numPlayers
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Button_Off
    End
End Sub

Private Sub Form_Load()
'    Play_Sound "SOUND1.WAV"
    numPlayers = 1
    Me.MouseIcon = frmMain.MouseIcon
End Sub

Private Sub optPlayers_Click(Index As Integer)
    numPlayers = Index + 1
End Sub
