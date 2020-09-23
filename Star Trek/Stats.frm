VERSION 5.00
Begin VB.Form frmStats 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   30
   ClientTop       =   -15
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Stats.frx":0000
   ScaleHeight     =   3120
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label cmdOkay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Return"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3915
      TabIndex        =   20
      Top             =   2790
      Width           =   915
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Label5"
      Height          =   375
      Left            =   3780
      TabIndex        =   19
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Label lblTotalPoints 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2025
      TabIndex        =   18
      Top             =   1605
      Width           =   555
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TOTAL POINTS:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   330
      TabIndex        =   17
      Top             =   1605
      Width           =   1275
   End
   Begin VB.Label lblFavorite 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   5
      Left            =   4500
      TabIndex        =   16
      Top             =   2085
      Width           =   555
   End
   Begin VB.Label lblFavorite 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   4
      Left            =   4500
      TabIndex        =   15
      Top             =   1785
      Width           =   555
   End
   Begin VB.Label lblFavorite 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   3
      Left            =   4500
      TabIndex        =   14
      Top             =   1485
      Width           =   555
   End
   Begin VB.Label lblFavorite 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   2
      Left            =   4500
      TabIndex        =   13
      Top             =   1185
      Width           =   555
   End
   Begin VB.Label lblFavorite 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Index           =   1
      Left            =   4500
      TabIndex        =   12
      Top             =   885
      Width           =   555
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Movies"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   5
      Left            =   3165
      TabIndex        =   11
      Top             =   2115
      Width           =   510
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Voyager"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   4
      Left            =   3150
      TabIndex        =   10
      Top             =   1785
      Width           =   570
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Deep Space 9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   3
      Left            =   3165
      TabIndex        =   9
      Top             =   1485
      Width           =   930
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Next Generation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   2
      Left            =   3165
      TabIndex        =   8
      Top             =   1185
      Width           =   1155
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Star Trek"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   1
      Left            =   3150
      TabIndex        =   7
      Top             =   855
      Width           =   660
   End
   Begin VB.Label lblOverallScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2025
      TabIndex        =   6
      Top             =   1965
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "OVERALL SCORE:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   330
      TabIndex        =   5
      Top             =   1965
      Width           =   1365
   End
   Begin VB.Label lblCorrectAnswers 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2025
      TabIndex        =   4
      Top             =   1245
      Width           =   555
   End
   Begin VB.Label lblTotalQuestions 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2025
      TabIndex        =   3
      Top             =   885
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "CORRECT ANSWERS:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   330
      TabIndex        =   2
      Top             =   1245
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TOTAL QUESTIONS:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   330
      TabIndex        =   1
      Top             =   885
      Width           =   1590
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER   NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5010
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOkay_Click()
    Button_Off
    Unload Me
End Sub

Public Sub Display_Stats()
    Dim OverallScore As String
    Dim n&
    
    Load frmStats
    lblPlayer.Caption = UCase$(Trim$(Players(Game.CurrentPlayer).Name))
    lblTotalQuestions.Caption = Format$(Players(Game.CurrentPlayer).Total_Questions, "000")
    lblCorrectAnswers.Caption = Format$(Players(Game.CurrentPlayer).Correct_Answers, "000")
    lblTotalPoints.Caption = Format$(Players(Game.CurrentPlayer).Score, "000")
    If Players(Game.CurrentPlayer).Total_Questions > 0 Then
        OverallScore = Format$(100 * (Players(Game.CurrentPlayer).Correct_Answers / Players(Game.CurrentPlayer).Total_Questions), "0")
    Else
        OverallScore = "100"
    End If
    lblOverallScore.Caption = OverallScore + "%"
    
    ' display favorite categories
    For n = 1 To 5
        lblFavorite(n).Caption = Format$(Players(Game.CurrentPlayer).Favorites(n), "000")
    Next
    
    frmStats.Show vbModal
End Sub

Private Sub Form_Load()
    Me.MouseIcon = frmMain.MouseIcon

End Sub
