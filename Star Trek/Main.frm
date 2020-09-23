VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ULTIMATE TREK TRIVIA"
   ClientHeight    =   7245
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Main.frx":030A
   MousePointer    =   99  'Custom
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQuestion 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1410
      Left            =   345
      MultiLine       =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Main.frx":0BD4
      Top             =   1200
      Width           =   3730
   End
   Begin VB.TextBox txtHolder 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   135
      Top             =   7290
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2220
      Left            =   4800
      Picture         =   "Main.frx":0C30
      ScaleHeight     =   2190
      ScaleWidth      =   3990
      TabIndex        =   1
      Top             =   4200
      Width           =   4020
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Player:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblScore2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   660
         Width           =   120
      End
      Begin VB.Label lblCategory 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   420
         Width           =   930
      End
      Begin VB.Label lblPlayer 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Player"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   180
         Width           =   645
      End
      Begin VB.Label lblScore1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Score:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   240
         TabIndex        =   5
         Top             =   660
         Width           =   645
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "T - 60"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   600
         Left            =   1365
         TabIndex        =   3
         ToolTipText     =   "Click to toggle pause"
         Top             =   1200
         Width           =   1365
      End
   End
   Begin VB.PictureBox picBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   4800
      Picture         =   "Main.frx":2114
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   264
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   780
      Width           =   3990
      Begin VB.PictureBox picImage 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   150
         Picture         =   "Main.frx":39E9
         ScaleHeight     =   164
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   244
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   150
         Width           =   3660
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3030
      Left            =   60
      Picture         =   "Main.frx":68D1
      ScaleHeight     =   3030
      ScaleWidth      =   9150
      TabIndex        =   12
      Top             =   0
      Width           =   9150
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   60
      Picture         =   "Main.frx":7F1F
      ScaleHeight     =   4140
      ScaleWidth      =   9180
      TabIndex        =   13
      Top             =   3060
      Width           =   9180
      Begin VB.TextBox txtAnswer 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Index           =   1
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Main.frx":8B1F
         Top             =   585
         Width           =   3150
      End
      Begin VB.TextBox txtAnswer 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Index           =   2
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Main.frx":8B80
         Top             =   1185
         Width           =   3150
      End
      Begin VB.TextBox txtAnswer 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Index           =   3
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Main.frx":8BDE
         Top             =   1800
         Width           =   3150
      End
      Begin VB.TextBox txtAnswer 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Index           =   4
         Left            =   870
         MultiLine       =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Main.frx":8C3C
         Top             =   2355
         Width           =   3150
      End
      Begin VB.CommandButton cmdAnswer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Index           =   1
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   540
         Width           =   490
      End
      Begin VB.CommandButton cmdAnswer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Index           =   2
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1125
         Width           =   490
      End
      Begin VB.CommandButton cmdAnswer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Index           =   3
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1710
         Width           =   490
      End
      Begin VB.CommandButton cmdAnswer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   490
         Index           =   4
         Left            =   300
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2295
         Width           =   490
      End
      Begin VB.Label lblMusic 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Music"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5700
         TabIndex        =   26
         ToolTipText     =   "Turn Music On/Off"
         Top             =   3780
         Width           =   690
      End
      Begin VB.Label cmdHelp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   6600
         TabIndex        =   23
         ToolTipText     =   "Player statistics"
         Top             =   3765
         Width           =   675
      End
      Begin VB.Label cmdInfo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Info"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   7500
         TabIndex        =   24
         ToolTipText     =   "Program information"
         Top             =   3765
         Width           =   495
      End
      Begin VB.Label lblDebug 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "DEBUG"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   4500
         TabIndex        =   22
         ToolTipText     =   "DEBUG MODE ACTIVE"
         Top             =   3765
         Width           =   960
      End
      Begin VB.Label cmdQuit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   8220
         TabIndex        =   25
         ToolTipText     =   "Exit the program"
         Top             =   3760
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ULTIMATE TREK TRIVIA
' COPYRIGHT 1993-1999 BY JONATHAN SCOTT HARBOUR

Option Explicit

Const TIME_DELAY = 61
Const MAX_ERROR = 3000

Dim Database As GameData
Dim iTimer%
Dim NL$
Dim DisplayPicture As Boolean
Dim History(1 To 3000) As Boolean
Const TOTAL_CORRECT_WAVES = 5
Const TOTAL_WRONG_WAVES = 5
Const TOTAL_ERROR_WAVES = 5
Dim configFile$, sound_file$
Dim n&, num&

Private Sub Fill_History_Array()
    For n = 1 To Game.LastRecord
        History(n) = False
    Next
End Sub

Private Sub cmdQuit_Click()
    Button_Click
    Change_Colors vbWhite
    End_Program
    Change_Colors Players(Game.CurrentPlayer).Color
End Sub

Private Sub Form_Load()
    Dim fav&
    
    txtQuestion.Text = ""
    txtAnswer(1).Text = ""
    txtAnswer(2).Text = ""
    txtAnswer(3).Text = ""
    txtAnswer(4).Text = ""
    
    Me.ScaleMode = 3
    Me.ScaleWidth = 640
    Me.ScaleHeight = 480
    
    Randomize
    NL = Chr$(10) & Chr$(13)
    configFile = App.Path & "\trek.trv"
    Game.dataFile = App.Path & "\" & LoadProfile("trivia", "datafile", configFile)
    If Mid$(UCase$(LoadProfile("trivia", "debugmode", configFile)), 1, 4) = "TRUE" Then
        DEBUG_MODE = True
    Else
        DEBUG_MODE = False
    End If
    Game.CurrentPlayer = 1
    PlayMusic = True
    
    frmIntro.Show vbModal
    pause
    If DEBUG_MODE Then
        lblDebug.Visible = True
        Game.numPlayers = 1
        Players(1).Name = "PLAYTESTER"
        Players(1).Color = vbYellow
        Players(1).Correct_Answers = 0
        Players(1).Total_Questions = 0
        Players(1).Score = 0
        Players(1).Multiplier = 1
        For fav = 1 To 15
            Players(1).Favorites(fav) = 0
        Next
    Else
        lblDebug.Visible = False
        frmNumPlayers.Show vbModal
        pause
        For num = 1 To Game.numPlayers
            Players(num).Name = Input_Name(num)
            pause
            Players(num).Color = Choose_Color(num)
            pause
            Players(num).Correct_Answers = 0
            Players(num).Total_Questions = 0
            Players(num).Score = 0
            Players(num).Multiplier = 1
            For fav = 1 To 15
                Players(num).Favorites(fav) = 0
            Next
        Next
    End If
    
    On Error GoTo errorhandler:
    
    ' calculate the length of a record
    Game.RecordLen = Len(Database)
    
    ' calculate number of records in file
    Game.LastRecord = FileLen(Game.dataFile) / Game.RecordLen

    If Game.LastRecord = 0 Then
        Error_Message UCase$(Game.dataFile) + " NOT FOUND!"
        End
    End If
    
    If Game.LastRecord < 10 Then
        Error_Message UCase$(Game.dataFile) + " MUST HAVE MORE DATA...IT'S TOO SMALL!"
        End
    End If
    
    Me.Show
    pause
    Reset_Timer
    Get_Next_Question
    Exit Sub
    
    Fill_History_Array
    
errorhandler:
    Error_Message "DISK I/O ERROR IN FORM_LOAD"
    Resume Next
End Sub

Private Sub cmdHelp_Click()
    Button_Click
    Timer2.Enabled = False
    Hide_Data
    frmStats.Display_Stats
    Timer2.Enabled = True
    Show_Data
End Sub

Private Sub Hide_Data()
    ' hide the question and answers so players can't cheat while in pause mode
    txtQuestion.Visible = False
    txtAnswer(1).Visible = False
    txtAnswer(2).Visible = False
    txtAnswer(3).Visible = False
    txtAnswer(4).Visible = False
End Sub

Private Sub Show_Data()
    ' re-enables the question and answer data
    txtQuestion.Visible = True
    txtAnswer(1).Visible = True
    txtAnswer(2).Visible = True
    txtAnswer(3).Visible = True
    txtAnswer(4).Visible = True
End Sub

Private Sub Random_Success_Message(ByRef msg1$, ByRef msg2$)
    Dim Message$, temp$
    Static Max&, textfile&, randmsg&
    Static first_time As Boolean
    
    textfile = FreeFile
    
    If Not first_time Then  ' this is backwards logic since starts as false
        first_time = True
        Open App.Path + "\Success.txt" For Input As textfile
        Do Until EOF(textfile)
            Line Input #textfile, Message$
            Max = Max + 1
        Loop
        Close textfile
    End If
    
    Open App.Path + "\Success.txt" For Input As textfile
    randmsg = Int((Rnd * Max) + 1)
    For num = 1 To randmsg
            Line Input #textfile, Message$
    Next
    Close textfile

    temp$ = Trim$(Message$)
    If Len(temp$) > 35 Then
        msg1$ = Trim$(Mid$(temp$, 1, 35))
        msg2$ = Trim$(Mid$(temp$, 36, Len(temp$)))
    Else
        msg1$ = Trim$(temp$)
        msg2$ = ""
    End If
End Sub

Private Sub Inc(ByRef num As Variant, Value As Variant)
        num = num + Value
End Sub

Private Sub Dec(ByRef num As Variant, Value As Variant)
        num = num - Value
End Sub

Private Sub Display_Correct_Answer()
    Dim msg$, msg1$, msg2$, sound_file$
            
    Inc Players(Game.CurrentPlayer).Correct_Answers, 1
    Inc Players(Game.CurrentPlayer).Total_Questions, 1
    Inc Players(Game.CurrentPlayer).Score, Players(Game.CurrentPlayer).Multiplier
    Select Case Game.Favorite
        Case "STAR TREK":       Inc Players(Game.CurrentPlayer).Favorites(1), 1
        Case "NEXT GENERATION": Inc Players(Game.CurrentPlayer).Favorites(2), 1
        Case "DEEP SPACE NINE": Inc Players(Game.CurrentPlayer).Favorites(3), 1
        Case "VOYAGER":         Inc Players(Game.CurrentPlayer).Favorites(4), 1
        Case Else:              Inc Players(Game.CurrentPlayer).Favorites(5), 1
    End Select
    
    sound_file$ = "correct" & Trim$(Int((TOTAL_CORRECT_WAVES * Rnd) + 1)) + ".wav"
    Play_Sound sound_file$
    
    Random_Success_Message msg1$, msg2$
    YouGainedPoints Players(Game.CurrentPlayer).Multiplier, msg1$, msg2$
    pause
    Inc Players(Game.CurrentPlayer).Multiplier, 1
End Sub

Private Sub Display_Wrong_Answer()
    Dim num As Integer: num = Game.CurrentPlayer
    Dim iPoints&
    Dim msg$
    
    iPoints = Players(num).Multiplier
    Inc Players(num).Total_Questions, 1
    Dec Players(num).Score, iPoints
    Players(num).Multiplier = 1
    
    pause
    sound_file$ = "wrong" + Trim$(Int((TOTAL_WRONG_WAVES * Rnd) + 1)) + ".wav"
    Play_Sound sound_file$
    
    If iPoints = 1 Then msg$ = " POINT!" Else msg$ = " POINTS!"
    MessageBox "YOU LOST " + Str$(iPoints) + msg$, "The correct answer was:", Game.CurrentAnswer, "&Blast", ""
    pause
End Sub

Private Sub cmdAnswer_Click(Index As Integer)
    Timer2.Enabled = False
    frmMain.picImage.Visible = False
    Toggle_Buttons
    If Index = Game.CorrectAnswer Then
        Display_Correct_Answer
    Else
        Display_Wrong_Answer
    End If
    Reset_Timer
    pause
    Get_Next_Question
    Toggle_Buttons
End Sub

Private Sub Toggle_Buttons()
    For n = 1 To 4
        cmdAnswer(n).Enabled = Not cmdAnswer(n).Enabled
    Next n
End Sub

Public Sub Reset_Timer()
    iTimer = TIME_DELAY
End Sub

Private Sub Change_Colors(ByVal col As Long)
    lblScore1.ForeColor = col
    lblScore2.ForeColor = col
    frmConfirm.lblTitle1.ForeColor = col
    frmConfirm.lblTitle2.ForeColor = col
    frmConfirm.lblTitle3.ForeColor = col
    lblPlayer.ForeColor = col
    lblTimer.ForeColor = col
    lblCategory.ForeColor = col
    Label1.ForeColor = col
    Label3.ForeColor = col
End Sub

Public Function Get_Unique_Answer(ByVal a1$, ByVal a2$, ByVal a3$) As String
    Dim iError&
    Dim temp$
    
    iError = 0
    On Error GoTo error
    Get_Random_Answer
    Do Until Trim(Database.Minor_Category) = Game.CurrentCategory _
        And temp$ <> a1$ And temp$ <> a2$ And temp$ <> a3$ And temp$ <> Game.CurrentAnswer
        Get_Random_Answer
        temp$ = Trim(Database.Answer)
        iError = iError + 1
        If iError > MAX_ERROR Then Program_Stuck
    Loop
    Get_Unique_Answer = temp$
    GoTo done
error:
    MsgBox "GET_UNIQUE_ANSWER: UNABLE TO FIND RANDOM ANSWER"
    Resume Next
done:
End Function

Public Sub Get_Next_Question()
    Dim iError
    Dim temp$
    Dim answers&, questions&, percent&
    
    Clear_Form
    Timer2.Enabled = False
    
    ' clear the tooltip hints for debug mode
    If DEBUG_MODE Then
        cmdAnswer(1).ToolTipText = ""
        cmdAnswer(2).ToolTipText = ""
        cmdAnswer(3).ToolTipText = ""
        cmdAnswer(4).ToolTipText = ""
    End If
    
    On Error GoTo Error_Open
    Open Game.dataFile For Random As #1 Len = Game.RecordLen

    ' choose an answer spot to put the "correct" answer into
    Game.CorrectAnswer = Int((4 * Rnd) + 1)
    
    On Error GoTo Error_Misc
    Get_Random_Question
    Game.CurrentQuestion = Trim$(Database.Question)
    Game.ImageFile = Trim$(Database.ImageFile)
    Game.SoundFile = Trim$(Database.SoundFile)
    Game.CurrentAnswer = Trim$(Database.Answer)
    Game.CurrentCategory = Trim$(Database.Minor_Category)
    Game.Favorite = Trim$(Database.Major_Category)

    ' next player's turn
    Game.CurrentPlayer = Game.CurrentPlayer + 1
    If Game.CurrentPlayer > Game.numPlayers Then
        Game.CurrentPlayer = 1
    End If
    
    ' change the colors
    Change_Colors Players(Game.CurrentPlayer).Color

'    Delay 1000
    AnnounceNextTurn UCase$(Trim$(Players(Game.CurrentPlayer).Name)) + "'S TURN", "CATEGORY: ", Game.CurrentCategory, "&Continue", ""
    
    frmMain.Caption = "Ultimate Trek Trivia"
    lblPlayer.Caption = UCase(Trim(Players(Game.CurrentPlayer).Name))
    lblCategory.Caption = Game.CurrentCategory
    
    ' display the score
    answers = Players(Game.CurrentPlayer).Correct_Answers
    questions = Players(Game.CurrentPlayer).Total_Questions
    If questions > 0 Then
        percent = (answers \ questions) * 100
    Else
        percent = 100
    End If
    lblScore2.Caption = Players(Game.CurrentPlayer).Score
    lblScore2.Caption = lblScore2.Caption & " (" & percent & "%)"
    
    txtAnswer(1) = ""
    txtAnswer(2) = ""
    txtAnswer(3) = ""
    txtAnswer(4) = ""
    
    If Game.CurrentCategory = "TRUE OR FALSE" Then
        If Game.CurrentAnswer = "TRUE" Then
            Game.CorrectAnswer = 1
            If DEBUG_MODE Then cmdAnswer(1).ToolTipText = "CORRECT"
        Else
            Game.CorrectAnswer = 2
            If DEBUG_MODE Then cmdAnswer(2).ToolTipText = "CORRECT"
        End If
        
        txtAnswer(1) = "TRUE"
        txtAnswer(2) = "FALSE"
                
        cmdAnswer(3).Visible = False
        cmdAnswer(4).Visible = False
    
    ElseIf Game.CurrentCategory = "YES OR NO" Then
        If Game.CurrentAnswer = "YES" Then
            Game.CorrectAnswer = 1
            If DEBUG_MODE Then cmdAnswer(1).ToolTipText = "CORRECT"
        Else
            Game.CorrectAnswer = 2
            If DEBUG_MODE Then cmdAnswer(2).ToolTipText = "CORRECT"
        End If
        
        txtAnswer(1) = "YES"
        txtAnswer(2) = "NO"
                
        cmdAnswer(3).Visible = False
        cmdAnswer(4).Visible = False
    Else
        
        ' get the wrong answers
        txtAnswer(1) = Get_Unique_Answer("", "", "")
        txtAnswer(2) = Get_Unique_Answer(txtAnswer(1).Text, "", "")
        txtAnswer(3) = Get_Unique_Answer(txtAnswer(1).Text, txtAnswer(2).Text, "")
        txtAnswer(4) = Get_Unique_Answer(txtAnswer(1).Text, txtAnswer(2).Text, txtAnswer(3).Text)
        
    End If
    
    On Error GoTo Error_Misc
    
    'display the question
    txtQuestion.Text = Trim(Game.CurrentQuestion)
    
    ' display the correct answer over one of the random answers
    txtAnswer(Game.CorrectAnswer).Text = Game.CurrentAnswer
    If DEBUG_MODE Then cmdAnswer(Game.CorrectAnswer).ToolTipText = "CORRECT"
    
    Close #1
    Timer2.Enabled = True
    
    ' display a picture
    temp$ = App.Path + "\pictures\" + Trim$(Game.ImageFile)
    If right$(temp$, 4) = ".JPG" Then
        picImage.Picture = LoadPicture(temp$)
        If DEBUG_MODE Then picImage.ToolTipText = temp$
    Else
        picImage.Picture = LoadPicture(App.Path + "\pictures\" + "PIC_000.JPG")
    End If
            
    frmMain.picImage.Visible = True
    frmMain.Refresh
    
    ' speak the question from the attached wave
    temp$ = Trim$(Game.SoundFile)
    If Len(temp$) > 0 Then
        Play_Sound "question.wav"
        Play_Sound temp$
    End If
    
    Exit Sub

Error_Misc:
    MsgBox error
    Resume Next
Error_Open:
    MsgBox "GET_NEXT_QUESTION: ERROR OPENING FILE"
    Resume Next
Error_Question:
    MsgBox "GET_NEXT_QUESTION: ERROR READING QUESTION"
    Resume Next
End Sub

Private Sub Get_Random_Question()
    Dim iCounter%, done%, iRandom%
    Dim FileExt$, GraphicFile$
    
    iCounter = 0
    done = False
    On Error GoTo error
    Do Until done Or iCounter > 1000
        Do
            iRandom = Int((Game.LastRecord * Rnd) + 1)
        Loop Until History(iRandom) = False
        
        ' has this question been used before?
        Get #1, iRandom, Database
        iCounter = iCounter + 1         ' avoid an infinite loop
        If Trim(Database.Question) <> "" Then done = True
    Loop
    If iCounter > MAX_ERROR Then Program_Stuck
    
    ' set question flag to true so it won't be asked again
    History(iRandom) = True
    Exit Sub

error:
    MsgBox error
    Resume Next
End Sub

Private Sub Get_Random_Answer()
    Dim iCounter&, iRandom&
    Dim done%
    
    iCounter = 0
    done = False
    On Error GoTo error
    Do
        iRandom = Int((Game.LastRecord * Rnd) + 1)
        Get #1, iRandom, Database
        iCounter = iCounter + 1         ' avoid an infinite loop
        If Trim$(Database.Answer) <> "" Then done = True
    Loop Until done Or iCounter > MAX_ERROR
    If iCounter > MAX_ERROR Then Program_Stuck
    Exit Sub

error:
    Error_Message "DATA ERROR IN GET_RANDOM_ANSWER"
    Resume Next
End Sub

Private Sub Program_Stuck()
    Dim msg$
    
    msg$ = "THERE'S AN ERROR IN THE DATABASE!" + vbCrLf
    msg$ = msg$ + "The category called '" + Trim$(Database.Minor_Category) + "' must" + vbCrLf
    msg$ = msg$ + "be used at LEAST 10 times so there will be enough multiple-" + vbCrLf
    msg$ = msg$ + "choice answers.  At present, the database is insufficient to" + vbCrLf
    msg$ = msg$ + "generate challenging trivia questions!  I'm going to skip the" + vbCrLf
    msg$ = msg$ + "category for now, but you might want to notify the author" + vbCrLf
    msg$ = msg$ + "if this problem persists!" + vbCrLf
    MsgBox msg$, vbCritical, "WE HAVE A PROBLEM . . ."
    
    Clear_Form
    Reset_Timer
    Get_Next_Question
End Sub

Private Sub Clear_Form()
    txtQuestion = ""
    
    txtAnswer(1).Text = ""
    txtAnswer(2).Text = ""
    txtAnswer(3).Text = ""
    txtAnswer(4).Text = ""
    
    cmdAnswer(3).Visible = True
    cmdAnswer(4).Visible = True
End Sub

Private Sub cmdInfo_Click()
    Button_Click
    Hide_Data
    About_Program
    Show_Data
End Sub

Private Function Get_High_Score() As Integer
    Dim FinalScore As Integer
    Dim n As Integer
    Dim highscore As Integer: highscore = 0
    Dim highplayer As Integer: highplayer = 0
    Dim GreaterThanZero As Boolean: GreaterThanZero = False
    Dim TieScore As Boolean: TieScore = False
    
    If Game.numPlayers > 1 Then
        ' find the high score
        For n = 1 To Game.numPlayers
            If Players(n).Score > 0 Then
                GreaterThanZero = True
                If Players(n).Score > highscore Then
                    highscore = Players(n).Score
                    highplayer = n
                ElseIf Players(n).Score = highscore Then
                    TieScore = True
                    Exit For
                End If
            End If
        Next
        
        If TieScore And GreaterThanZero Then
            FinalScore = 0
        Else
            FinalScore = highplayer     ' we have a winner!
        End If
    Else
        ' there's only one player this game
        FinalScore = Players(Game.CurrentPlayer).Score
    End If
    
    Get_High_Score = FinalScore
End Function

Private Sub End_Program()
    Dim num As Integer
    
    Hide_Data
    If Not DEBUG_MODE Then
        Timer2.Enabled = False
        Dim reply As Boolean
        reply = MessageBox("ULTIMATE TREK TRIVIA", "Are you sure", "you want to quit?", "&Yes", "&NO!")
        pause
        If reply = True Then
            frmMain.Hide
            PlayMusic = False
            frmMidi.Stop_Playing
            num = Get_High_Score()
            
            ' check for a tie score
            If num > 0 Then                ' nope
                ShowHiScores Players(Game.CurrentPlayer).Score, "Ultimate Trek Trivia", App.Path & "\Trivia.ini", False
                pause
            End If
            
            frmCopyright.Show vbModal
            pause
            End
        End If
        Timer2.Enabled = True
        Show_Data
    Else
        End
    End If
End Sub

Private Sub About_Program()
    Timer2.Enabled = False
    frmCopyright.Show vbModal
    pause
    Timer2.Enabled = True
End Sub

Private Sub lblMusic_Click()
    PlayMusic = Not PlayMusic
    If PlayMusic = True Then
        lblMusic.ForeColor = &H0&
        frmMidi.Play_Next
    Else
        lblMusic.ForeColor = &H80&
        frmMidi.Stop_Playing
    End If
    
End Sub

Private Sub lblTimer_Click()
    Static status As Boolean
    
    Button_Click
    status = Not status
    If status Then
        Show_Data
        Timer2.Enabled = True
        cmdAnswer(1).Enabled = True
        cmdAnswer(2).Enabled = True
        cmdAnswer(3).Enabled = True
        cmdAnswer(4).Enabled = True
    Else
        Hide_Data
        Timer2.Enabled = False
        cmdAnswer(1).Enabled = False
        cmdAnswer(2).Enabled = False
        cmdAnswer(3).Enabled = False
        cmdAnswer(4).Enabled = False
    End If
End Sub

Private Sub Timer2_Timer()
    iTimer = iTimer - 1
    lblTimer.Caption = "T - " & Trim(Str$(iTimer))
    If iTimer < 1 Then cmdAnswer_Click 0
End Sub

Public Sub pause()
    Delay 500
End Sub

Private Sub txtAnswer_GotFocus(Index As Integer)
    txtHolder.SetFocus
End Sub

Private Sub txtHolder_Change()
    Dim s$
    If Len(s$) > 0 Then
        s$ = left$(txtHolder.Text, 1)
        If s$ = "1" Then cmdAnswer_Click 1
        If s$ = "2" Then cmdAnswer_Click 2
        If s$ = "3" Then cmdAnswer_Click 3
        If s$ = "4" Then cmdAnswer_Click 4
    End If
    txtHolder.Text = ""
End Sub

Private Sub txtHolder_GotFocus()
    txtHolder.Text = ""
End Sub

Private Sub txtQuestion_GotFocus()
    txtHolder.SetFocus
End Sub

Private Sub txtScore_GotFocus()
    txtHolder.SetFocus
End Sub

