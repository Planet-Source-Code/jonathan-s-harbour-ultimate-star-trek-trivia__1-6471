VERSION 5.00
Begin VB.Form frmNames 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3105
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   5355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Names.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&OK"
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
      Height          =   405
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1170
   End
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   315
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1305
      Width           =   4725
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player #1, please type"
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
      Height          =   510
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   4995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "your name here:"
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
      Height          =   510
      Left            =   180
      TabIndex        =   0
      Top             =   630
      Width           =   5010
   End
End
Attribute VB_Name = "frmNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdContinue_Click()
        Button_Off
        frmNames.Hide
End Sub

Private Sub Form_Load()
    Me.MouseIcon = frmMain.MouseIcon
    frmNames.Show
    frmNames.txtInput.Text = ""
    frmNames.txtInput.SetFocus
    frmNames.Hide
End Sub

Public Function Input_Name(ByVal num As Integer) As String
        frmNames.lblTitle.Caption = "Player #" + Str$(num) + ", please type"
        Play_Sound "SOUND1.WAV"
        frmNames.Show vbModal
        Input_Name = frmNames.txtInput.Text
        Unload frmNames
End Function

