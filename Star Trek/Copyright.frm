VERSION 5.00
Begin VB.Form frmCopyright 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COPYRIGHTS AND LEGAL STUFF"
   ClientHeight    =   4380
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Copyright.frx":0000
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Copyright.frx":19B89
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ULTIMATE TREK TRIVIA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1155
      Left            =   1260
      TabIndex        =   0
      Top             =   120
      Width           =   4635
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1140
      Left            =   120
      Picture         =   "Copyright.frx":19CEC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "frmCopyright"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
    frmCopyright.Hide
End Sub

Private Sub Form_Load()
    Me.MouseIcon = frmMain.MouseIcon

End Sub
