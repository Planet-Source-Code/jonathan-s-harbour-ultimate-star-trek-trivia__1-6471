VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Colors.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picColor 
      BackColor       =   &H00FF0000&
      FillColor       =   &H000000FF&
      Height          =   1032
      Index           =   4
      Left            =   3945
      ScaleHeight     =   975
      ScaleWidth      =   825
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   888
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BLUE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FF00&
      FillColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1032
      Index           =   3
      Left            =   2805
      ScaleHeight     =   975
      ScaleWidth      =   825
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1350
      Width           =   888
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "GREEN"
         Height          =   252
         Index           =   3
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H0000FFFF&
      FillColor       =   &H000000FF&
      Height          =   1032
      Index           =   2
      Left            =   1605
      ScaleHeight     =   975
      ScaleWidth      =   825
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1350
      Width           =   888
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "YELLOW"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   60
         TabIndex        =   7
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.PictureBox picColor 
      BackColor       =   &H000000FF&
      FillColor       =   &H000000FF&
      Height          =   1032
      Index           =   1
      Left            =   405
      ScaleHeight     =   975
      ScaleWidth      =   825
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1350
      Width           =   888
      Begin VB.Label lblColor 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RED"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR GAME COLOR:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   510
      Left            =   105
      TabIndex        =   1
      Top             =   675
      Width           =   5010
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYER #1, CHOOSE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   4965
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Chosen_Color As Integer
Dim n&

Public Function Choose_Color(ByVal num As Integer) As Long
        Static Chosen(1 To 4) As Boolean
        Load frmColors
        For n = 1 To 4
                If Chosen(n) = True Then picColor(n).Visible = False
        Next
        
        frmColors.lblTitle.Caption = "PLAYER #" + Str$(num) + ", CHOOSE"
        frmColors.Show vbModal
        Chosen(Chosen_Color) = True
        Choose_Color = picColor(Chosen_Color).BackColor
        Unload frmColors
End Function

Private Sub Form_Load()
    Me.MouseIcon = frmMain.MouseIcon
End Sub

Private Sub lblColor_Click(Index As Integer)
        Button_Click
        Chosen_Color = Index
        frmColors.Hide
End Sub

Private Sub picColor_Click(Index As Integer)
        Button_Click
        Chosen_Color = Index
        frmColors.Hide
End Sub
