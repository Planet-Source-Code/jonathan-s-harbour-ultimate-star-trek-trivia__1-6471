VERSION 5.00
Begin VB.Form frmScores 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "High Scores"
   ClientHeight    =   3885
   ClientLeft      =   2850
   ClientTop       =   1515
   ClientWidth     =   4185
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Score2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3885
   ScaleWidth      =   4185
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3420
      Top             =   3240
   End
   Begin VB.CommandButton btnNewScore 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   5
      Top             =   3240
      Width           =   1185
   End
   Begin VB.TextBox txtScore 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2130
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "Text2"
      Top             =   2820
      Width           =   1965
   End
   Begin VB.CommandButton btnOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   3480
      Width           =   1185
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2820
      Width           =   1905
   End
   Begin VB.ListBox lstScores 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1965
   End
   Begin VB.ListBox lstNames 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1155
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1905
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SCORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2160
      TabIndex        =   8
      Top             =   840
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PLAYERS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1875
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      Picture         =   "Score2.frx":030A
      Top             =   5400
      Width           =   4230
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      Picture         =   "Score2.frx":1E8C
      Top             =   3840
      Width           =   4230
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   0
      Picture         =   "Score2.frx":3A0E
      Top             =   4620
      Width           =   4230
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Congatulations!  You've just achieved a New High Score!  Enter your name below:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   180
      TabIndex        =   6
      Top             =   2400
      Width           =   3765
   End
   Begin VB.Image imgMain 
      Appearance      =   0  'Flat
      Height          =   576
      Left            =   0
      Picture         =   "Score2.frx":5590
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4164
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
' SCORE1.FRM
' This form is dependent on file SCORE1.BAS.
'------------------------------------------------------------


Private Sub btnNewScore_Click()
'------------------------------------------------------------
' When this button is pressed, save the new player name and
' score, then hide the text boxes and button used to enter
' the player's name, and resize the form.
'------------------------------------------------------------

    ' Save all high scores back to the .INI file.
    AddScoreAndSave txtName, txtScore

    DisplayScores

    SetForDisplay
End Sub

Private Sub btnOK_Click()
'------------------------------------------------------------
' Close the frmScores window when this button is pushed.
'------------------------------------------------------------

    Timer1.Enabled = False
    Timer1.Interval = 0
    DoEvents

    Unload Me
End Sub

Private Sub DisplayScores()
'------------------------------------------------------------
' Display the scores and player names from the Hi() array
' into the form's list controls.
'------------------------------------------------------------
Dim i As Integer
                                      
    If Num_HiScores > 0 Then
        ' Empty the lists.
        lstNames.Clear
        lstScores.Clear

        ' Display the high scores in the list boxes.
        For i = 1 To Num_HiScores
            lstNames.AddItem Hi(i).Name
            lstScores.AddItem Format$(Hi(i).Score)
        Next
    End If

End Sub

Private Sub Form_Load()
'------------------------------------------------------------
' When the form is loaded, center it and display the current
' high scores.
'------------------------------------------------------------
Dim rc As Long

    ' Center the form on the screen.
    Me.Left = (Screen.Width - Me.Width) \ 2
    Me.Top = (Screen.Height - Me.Height) \ 2

    Me.Caption = gGameTitle

    ' Display current high scores.
    DisplayScores

    If gDisplayOnly Then
        SetForDisplay
    Else
        
        btnOK.Visible = False
        btnNewScore.Default = True
    
        ' Clear text field to let player enter their name.
        txtName = ""
        txtName.MaxLength = 15
        txtScore = Format$(gNewScore)
    
        'rc = SendMessage(txtScore.hWnd, EM_SETREADONLY, 1, 0)
        txtScore.Locked = True
    End If
End Sub

Private Sub Form_Paint()
    Make3D Me, lstNames, 1
    Make3D Me, lstScores, 1
    If txtName.Visible Then
        Make3D Me, txtName, 0
        Make3D Me, txtScore, 0
    End If
End Sub

Private Sub Make3D(pic As Form, ctl As Control, ByVal BorderStyle As Integer)
'--------------------------------------------------
' Wrap a 3D effect around a control on a form.
'--------------------------------------------------
Dim AdjustX As Integer, AdjustY As Integer
Dim RightSide As Single
Dim BW As Integer, BorderWidth As Integer
Dim LeftTopColor As Long, RightBottomColor As Long
Dim i As Integer
' Color Constants
Const DARK_GRAY = &H808080
Const WHITE = &HFFFFFF
Const BLACK = &H0


    If Not ctl.Visible Then Exit Sub

    AdjustX = Screen.TwipsPerPixelX
    AdjustY = Screen.TwipsPerPixelY

    BorderWidth = 1

    Select Case BorderStyle
        Case 0: ' Inset
            LeftTopColor = DARK_GRAY
            RightBottomColor = WHITE
        Case 1: ' Raised
            LeftTopColor = WHITE
            RightBottomColor = DARK_GRAY
    End Select
    

    ' Set the top shading line.
    For BW = 1 To BorderWidth
        ' Top
        pic.CurrentX = ctl.Left - (AdjustX * BW)
        pic.CurrentY = ctl.Top - (AdjustY * BW)
        pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top - (AdjustY * BW)), LeftTopColor
        ' Right
        pic.Line -(ctl.Left + ctl.Width + (AdjustX * (BW - 1)), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
        ' Bottom
        pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top + ctl.Height + (AdjustY * (BW - 1))), RightBottomColor
        ' Left
        pic.Line -(ctl.Left - (AdjustX * BW), ctl.Top - (AdjustY * BW)), LeftTopColor
    Next
End Sub

Private Sub SetForDisplay()

    ' Hide "New Hi" controls...
    lblInfo.Visible = False
    txtName.Visible = False
    txtScore.Visible = False
    btnNewScore.Visible = False

    ' Adjust the OK button position and Window Height.
    btnOK.Visible = True
    btnOK.Top = lblInfo.Top
    btnOK.Default = True
    Me.Height = btnNewScore.Top + 45
    Me.Refresh

End Sub

Private Sub Timer1_Timer()
Static InSub As Integer
Dim StartTime As Single
Dim i As Integer

    If InSub Then Exit Sub
    InSub = True

    For i = 1 To 4
        imgMain.Picture = Image2.Picture
        StartTime = Timer
        Do While (Timer - StartTime) < 0.1
            DoEvents
        Loop
    
        imgMain.Picture = Image3.Picture
        StartTime = Timer
        Do While (Timer - StartTime) < 0.1
            DoEvents
        Loop
    Next

    imgMain.Picture = Image1.Picture
    InSub = False
End Sub

