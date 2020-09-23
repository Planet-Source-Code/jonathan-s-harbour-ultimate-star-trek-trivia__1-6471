VERSION 5.00
Begin VB.Form frmProjects 
   Caption         =   "TRIVIA PROJECTS"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6405
   Icon            =   "frmProject.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox filename 
      Height          =   285
      Left            =   3195
      TabIndex        =   8
      Top             =   3555
      Width           =   3120
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&CANCEL"
      Height          =   420
      Left            =   4995
      TabIndex        =   7
      Top             =   4050
      Width           =   1320
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&LOAD"
      Height          =   420
      Left            =   3240
      TabIndex        =   6
      Top             =   4050
      Width           =   1320
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   270
      Width           =   2940
   End
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   3195
      Pattern         =   "*.trv"
      TabIndex        =   1
      Top             =   270
      Width           =   3120
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   90
      TabIndex        =   0
      Top             =   990
      Width           =   2940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DATA FILES:"
      Height          =   195
      Index           =   2
      Left            =   3195
      TabIndex        =   5
      Top             =   45
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "DRIVES:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FOLDERS:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   765
      Width           =   795
   End
End
Attribute VB_Name = "frmProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    filename.Text = ""
    Me.Hide
End Sub

Private Sub cmdLoad_Click()
    Me.Hide
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    filename.Text = File1.List(File1.ListIndex)
End Sub

Private Sub File1_DblClick()
    File1_Click
    cmdLoad_Click
End Sub
