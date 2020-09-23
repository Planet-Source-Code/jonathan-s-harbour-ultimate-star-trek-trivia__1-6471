VERSION 5.00
Begin VB.Form frmEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trivia Editor"
   ClientHeight    =   5700
   ClientLeft      =   30
   ClientTop       =   645
   ClientWidth     =   8145
   ForeColor       =   &H000080FF&
   Icon            =   "Edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Edit.frx":0442
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "SELECT..."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2520
      TabIndex        =   24
      TabStop         =   0   'False
      ToolTipText     =   "save answer to history popup list"
      Top             =   4020
      Width           =   1380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PLAY WAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   180
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "save answer to history popup list"
      Top             =   4020
      Width           =   1320
   End
   Begin VB.CommandButton cmdSelectImage 
      Caption         =   "SELECT..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7020
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "save answer to history popup list"
      Top             =   1440
      Width           =   960
   End
   Begin VB.PictureBox picImageFile 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      Height          =   2520
      Left            =   4260
      ScaleHeight     =   164
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   244
      TabIndex        =   21
      Top             =   1740
      Width           =   3720
   End
   Begin VB.CommandButton cmdEditCategory 
      Caption         =   "EDIT CATEGORIES"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2220
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "save answer to history popup list"
      Top             =   2700
      Width           =   1440
   End
   Begin VB.CommandButton cmdEditSubject 
      Caption         =   "EDIT SUBJECTS"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "save answer to history popup list"
      Top             =   1740
      Width           =   1440
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Quit"
      Height          =   476
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "End Program"
      Top             =   4920
      Width           =   732
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Last"
      Height          =   476
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Last Record"
      Top             =   4920
      Width           =   732
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Next"
      Height          =   476
      Left            =   4020
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Next Record"
      Top             =   4920
      Width           =   732
   End
   Begin VB.CommandButton cmdPrev 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Prev"
      Height          =   476
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Prev Record"
      Top             =   4920
      Width           =   732
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&First"
      Height          =   476
      Left            =   1860
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "First Record"
      Top             =   4920
      Width           =   732
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&New"
      Height          =   476
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "New Record"
      Top             =   4920
      Width           =   732
   End
   Begin VB.CommandButton cmdRemember 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6960
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "save answer to history popup list"
      Top             =   180
      Width           =   1020
   End
   Begin VB.TextBox txtSoundFile 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   1500
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "FILENAME.WAV"
      Top             =   3705
      Width           =   2415
   End
   Begin VB.TextBox txtImageFile 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   5580
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "FILENAME.JPG"
      Top             =   1425
      Width           =   1332
   End
   Begin VB.ComboBox cboMainCat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "Edit.frx":19FCB
      Left            =   120
      List            =   "Edit.frx":19FCD
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "PRIMARY (MAJOR) CATEGORIES"
      Top             =   1980
      Width           =   3825
   End
   Begin VB.TextBox txtAnswer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4140
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   3852
   End
   Begin VB.TextBox txtQuestion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   972
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   420
      Width           =   3852
   End
   Begin VB.ComboBox cboCategory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "SECONDARY CATEGORIES"
      Top             =   2940
      Width           =   3795
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1999 Jonathan S. Harbour"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   6105
      TabIndex        =   26
      Top             =   5520
      Width           =   1995
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trivia Editor v3.1"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   60
      TabIndex        =   25
      Top             =   5520
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SOUND FILE:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   150
      TabIndex        =   18
      Top             =   3705
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMAGE FILE:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4230
      TabIndex        =   10
      Top             =   1425
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "right-click for quick list"
      Top             =   2700
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ANSWER:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4155
      TabIndex        =   6
      ToolTipText     =   "right-click for history popup list"
      Top             =   135
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUESTION:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUBJECT:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   150
      TabIndex        =   4
      ToolTipText     =   "right-click for quick list"
      Top             =   1710
      Width           =   960
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Project"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Project"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export..."
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "&Quit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuFirst 
         Caption         =   "First"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "Previous"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuNext 
         Caption         =   "Next"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuLast 
         Caption         =   "Last"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuFind 
         Caption         =   "Find String"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuStatistics 
      Caption         =   "&Statistics"
      Begin VB.Menu mnuSubject 
         Caption         =   "Subject Category"
      End
      Begin VB.Menu mnuQuestion 
         Caption         =   "Question Category"
      End
   End
   Begin VB.Menu mnuMajorPopup 
      Caption         =   "Major Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu Major_Popup 
         Caption         =   "Major Popup"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMinorPopup 
      Caption         =   "Category Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu Minor_Popup 
         Caption         =   "Minor Popup"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAnswerList 
      Caption         =   "Answer List Popup Menu"
      Visible         =   0   'False
      Begin VB.Menu Answer_List_Popup 
         Caption         =   "Answer List Popup"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ULTIMATE TREK TRIVIA EDITOR
' COPYRIGHT 1997 BY JONATHAN SCOTT HARBOUR

Option Explicit

Public dataFile$, dataPath$, configFile$
Dim sCategory$(1 To 30)
Dim iCategory_Count%
Dim dbInfo As GameData
Dim dbRecordLen As Long
Dim dbCurrentRecord As Long
Public dbLastRecord As Long
Dim iTotal(0 To 5) As Integer
Dim MAX_OPTIONS As Integer
Dim Current_Option As Integer
Dim Search_Text As String
Dim Search_Count As Integer

Public Sub LoadMajorCategories()
    Dim n&, cat$
    NUM_MAJORS = 0
    For n = 0 To 99
        cat$ = LoadProfile("major categories", "cat" & n, configFile)
        If Len(cat$) > 0 Then
            NUM_MAJORS = NUM_MAJORS + 1
            cboMainCat.AddItem Trim$(cat$)      ' load the dropdown list with the data
            Load Major_Popup(NUM_MAJORS)        ' load the popup menu with the data
            Major_Popup(NUM_MAJORS).Caption = Trim$(cat$)
        End If
    Next n
    Major_Popup(0).Visible = False
End Sub

Public Sub SaveMajorCategories()
    Dim n&, ans$
    For n = 0 To frmEdit.cboMainCat.ListCount - 1
        ans$ = frmEdit.cboMainCat.List(n)
        SaveProfile "major categories", "cat" & n, ans, configFile
    Next
    For n = frmEdit.cboMainCat.ListCount To 99
        SaveProfile "major categories", "cat" & n, "", configFile
    Next n
End Sub

Public Sub LoadMinorCategories()
    Dim n&, cat$
    NUM_CATS = 0
    For n = 0 To 99
        cat$ = LoadProfile("minor categories", "cat" & n, configFile)
        If Len(cat$) > 0 Then
            NUM_CATS = NUM_CATS + 1
            cboCategory.AddItem Trim$(cat$)   ' load the dropdown list with the data
            Load Minor_Popup(NUM_CATS)        ' load the popup menu with the data
            Minor_Popup(NUM_CATS).Caption = Trim$(cat$)
        End If
    Next n
    Minor_Popup(0).Visible = False
End Sub

Public Sub SaveMinorCategories()
    Dim n&, ans$
    For n = 0 To frmEdit.cboCategory.ListCount - 1
        ans$ = frmEdit.cboCategory.List(n)
        SaveProfile "minor categories", "cat" & n, ans, configFile
    Next
    For n = frmEdit.cboCategory.ListCount To 99
        SaveProfile "minor categories", "cat" & n, "", configFile
    Next n
End Sub

Public Sub LoadAnswerHistory()
    Dim n&, cat$
    NUM_ANSWERS = 0
    For n = 0 To 99
        cat$ = LoadProfile("answer history", "answer" & n, configFile)
        If Len(cat$) > 0 Then
            NUM_ANSWERS = NUM_ANSWERS + 1
            Load Answer_List_Popup(NUM_ANSWERS) ' load the popup menu with the data
            Answer_List_Popup(NUM_ANSWERS).Caption = Trim$(cat$)
        End If
    Next n
    Answer_List_Popup(0).Visible = False
End Sub

Private Sub Save_Answer_History()
    Dim n&, ans$
    For n = 1 To NUM_ANSWERS
        ans$ = Answer_List_Popup(n).Caption
        SaveProfile "answer history", "answer" & n, ans, configFile
    Next
End Sub

Private Sub cmdEditCategory_Click()
    frmCategories.Show vbModal
End Sub

Private Sub cmdEditSubject_Click()
    frmSubjects.Show vbModal
End Sub

Private Sub mnuFileExport_Click()
    frmExport.Show vbModal
End Sub

Private Sub mnuFileOpen_Click()
    frmProjects.Show vbModal
    If Len(frmProjects.FileName) > 0 Then
        dataPath = frmProjects.Dir1.Path
        configFile = dataPath & "\" & frmProjects.FileName
        dataFile = dataPath & "\" & LoadProfile("trivia", "datafile", configFile)
        LoadMajorCategories
        LoadMinorCategories
        LoadAnswerHistory
        Open_Data_File
    End If
End Sub

Private Sub Form_Load()
    Me.ScaleMode = 3
    Me.ScaleWidth = 640
    Me.ScaleHeight = 480
    MAX_OPTIONS = 0             ' set up the Save Answer button and options
    Current_Option = 0
    Show
    mnuFileOpen_Click
End Sub

Private Sub cmdFirst_Click()
    SaveCurrentRecord
    dbCurrentRecord = 1
    ShowCurrentRecord
End Sub

Private Sub cmdLast_Click()
    SaveCurrentRecord
    dbCurrentRecord = dbLastRecord
    ShowCurrentRecord
End Sub

Private Sub cmdNew_Click()
    Static last_major$, last_cat$
    
     SaveCurrentRecord
     
    ' reset multimedia filenames
    txtImageFile.Text = ""
    txtSoundFile.Text = ""
    
    ' save last categories
    last_major$ = dbInfo.Major_Category
    last_cat$ = dbInfo.Minor_Category
    
     ' add a new blank record
     dbLastRecord = dbLastRecord + 1
     dbInfo.Question = ""
     dbInfo.Answer = ""
     dbInfo.Major_Category = last_major$
     dbInfo.Minor_Category = last_cat$
     'dbInfo.Relation = OTHER
     
     Put #1, dbLastRecord, dbInfo
     
     ' update current record
     dbCurrentRecord = dbLastRecord
     
     ' display the record that was just created
     ShowCurrentRecord
     
     txtQuestion.SetFocus
End Sub

Private Sub cmdNext_Click()
    If dbCurrentRecord = dbLastRecord Then
        Beep
    Else
        SaveCurrentRecord
        dbCurrentRecord = dbCurrentRecord + 1
        ShowCurrentRecord
    End If
End Sub

Private Sub cmdPrev_Click()
    If dbCurrentRecord = 1 Then
        Beep
    Else
        SaveCurrentRecord
        dbCurrentRecord = dbCurrentRecord - 1
        ShowCurrentRecord
    End If
End Sub

Private Sub cmdQuit_Click()
    Save_Answer_History
    SaveMajorCategories
    SaveMinorCategories
    End
End Sub

Public Sub Open_Data_File()
    ' calculate the length of a record
    dbRecordLen = Len(dbInfo)
    
    ' open the file for random access
    ' if file does not exist it will be created
    Open dataFile For Random As #1 Len = dbRecordLen
    
    ' update current record
    dbCurrentRecord = 1
    
    ' find the last record number in the file
    dbLastRecord = FileLen(dataFile) / dbRecordLen
    
    ' if file was just created then set lastrecord to 0
    If dbLastRecord = 0 Then
        dbLastRecord = 1
    End If
    
    ' display the current record
    ShowCurrentRecord

End Sub

Public Sub ShowCurrentRecord()
    Dim temp As String
    Dim s As String
    Dim pos As Long
    
    ' read a record into dbInfo
    Get #1, dbCurrentRecord, dbInfo
    
    'display the record
    txtQuestion.Text = Trim(dbInfo.Question)
    txtAnswer.Text = Trim(dbInfo.Answer)
    
    ' set up the category for this record
    cboMainCat.Text = dbInfo.Major_Category
    cboCategory.Text = dbInfo.Minor_Category
    
    ' set multimedia file field
    txtImageFile.Text = Trim$(UCase$(dbInfo.ImageFile))
    On Error Resume Next
    picImageFile.Picture = LoadPicture(App.Path & "\pictures\" & txtImageFile.Text)
    txtSoundFile.Text = Trim$(UCase$(dbInfo.SoundFile))
    
    ' set caption to reflect current record
    frmEdit.Caption = "TRIVIA EDITOR (REC " + _
    Str(dbCurrentRecord) + " / " + _
    Str(dbLastRecord) + ")"
                
End Sub

Public Sub SaveCurrentRecord()
    Dim temp$
    
    ' save question field
    dbInfo.Question = Trim(UCase(txtQuestion.Text))
    
    ' save answer field
    dbInfo.Answer = Trim(UCase(txtAnswer.Text))
    
    ' save category fields
    dbInfo.Major_Category = Trim(UCase(cboMainCat.Text))
    dbInfo.Minor_Category = Trim(UCase(cboCategory.Text))
    
    ' save multimedia file field
    temp$ = Trim$(UCase$(txtImageFile.Text))
    If temp$ = ".JPG" Then temp$ = ""
    dbInfo.ImageFile = temp$
    temp$ = Trim$(UCase$(txtSoundFile.Text))
    If temp$ = ".WAV" Then temp$ = ""
    dbInfo.SoundFile = temp$
    
    ' write record to disk
    Put #1, dbCurrentRecord, dbInfo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' save the current record
    SaveCurrentRecord
    
    ' close the file
    Close #1
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button = 2 Then
                PopupMenu mnuMajorPopup
        End If
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuAnswerList
    End If
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuMinorPopup
    End If
End Sub

Private Sub Major_Popup_Click(Index As Integer)
    cboMainCat.Text = Major_Popup(Index).Caption
End Sub

Private Sub Minor_Popup_Click(Index As Integer)
    cboCategory.Text = Minor_Popup(Index).Caption
End Sub

Private Sub mnuFind_Click()
    Dim p, n As Integer
    
    SaveCurrentRecord
    Search_Count = 0
    Search_Text = InputBox("PLEASE ENTER A SEARCH STRING:", "HIDE AND SEEK")
    If Search_Text = "" Then Exit Sub
    Search_Text = UCase$(Trim$(Search_Text))

    For n = 1 To dbLastRecord
        Get #1, n, dbInfo
        
        ' search the question field
        p = InStr(1, dbInfo.Question, Search_Text, vbTextCompare)
        If p > 0 Then
            dbCurrentRecord = n
            ShowCurrentRecord
            txtQuestion.SetFocus
            txtQuestion.SelStart = p - 1
            txtQuestion.SelLength = Len(Search_Text)
            Search_Count = Search_Count + 1
            Exit Sub
        End If
    
        ' search the answer field
        p = InStr(1, dbInfo.Answer, Search_Text, vbTextCompare)
        If p > 0 Then
            dbCurrentRecord = n
            ShowCurrentRecord
            txtAnswer.SetFocus
            txtAnswer.SelStart = p - 1
            txtAnswer.SelLength = Len(Search_Text)
            Search_Count = Search_Count + 1
            Exit Sub
        End If
    Next

    MsgBox Search_Text + vbCrLf + "SEARCH COUNT: " + Format$(Search_Count), vbInformation, "HIDE AND SEEK"
End Sub

Private Sub mnuFindNext_Click()
    Dim n As Integer
    Dim p
    
    If Search_Text = "" Then Exit Sub
    SaveCurrentRecord
    For n = dbCurrentRecord + 1 To dbLastRecord
        Get #1, n, dbInfo
        
        ' search the question field
        p = InStr(1, dbInfo.Question, Search_Text, vbTextCompare)
        If p > 0 Then
            dbCurrentRecord = n
            ShowCurrentRecord
            txtQuestion.SetFocus
            txtQuestion.SelStart = p - 1
            txtQuestion.SelLength = Len(Search_Text)
            Search_Count = Search_Count + 1
            Exit Sub
        End If
    
        ' search the answer field
        p = InStr(1, dbInfo.Answer, Search_Text, vbTextCompare)
        If p > 0 Then
            dbCurrentRecord = n
            ShowCurrentRecord
            txtAnswer.SetFocus
            txtAnswer.SelStart = p - 1
            txtAnswer.SelLength = Len(Search_Text)
            Search_Count = Search_Count + 1
            Exit Sub
        End If
    Next

    MsgBox Search_Text + vbCrLf + "SEARCH COUNT: " + Format$(Search_Count), vbInformation, "HIDE AND SEEK"
    
End Sub

Private Sub mnuFirst_Click()
    cmdFirst_Click
End Sub

Private Sub mnuLast_Click()
    cmdLast_Click
End Sub

Private Sub mnuNew_Click()
    cmdNew_Click
End Sub

Private Sub mnuNext_Click()
    cmdNext_Click
End Sub

Private Sub mnuPrevious_Click()
    cmdPrev_Click
End Sub

Private Sub mnuQuit_Click()
        cmdQuit_Click
End Sub

Private Function blanks(num As Integer) As String
    Dim s$, n
    s$ = ""
    For n = 1 To num
        s$ = s$ + " "
    Next
    blanks$ = s$
End Function

Private Sub mnuSubject_Click()
    Dim count(1 To 50) As Integer
    Dim n, majors
    
    SaveCurrentRecord

    For n = 1 To dbLastRecord
        Get #1, n, dbInfo
        For majors = 1 To NUM_MAJORS
            If Trim$(dbInfo.Major_Category) = Major_Popup(majors).Caption Then
                count(majors) = count(majors) + 1
            End If
        Next
    Next

    For n = 1 To frmStats.lblCount.UBound
        frmStats.lblCount(n).Visible = False
        frmStats.lblName(n).Visible = False
    Next
    
    For n = 1 To NUM_MAJORS
        frmStats.lblCount(n - 1).Visible = True
        frmStats.lblName(n - 1).Visible = True
        frmStats.lblCount(n - 1).Caption = Format$(count(n))
        frmStats.lblName(n - 1).Caption = Major_Popup(n).Caption
    Next

    frmStats.Caption = "SUBJECT CATEGORY STATISTICS"
    frmStats.Show vbModal
End Sub

Private Sub mnuQuestion_Click()
    Dim count(1 To 50) As Integer
    Dim n, minors
    
    SaveCurrentRecord

    For n = 1 To dbLastRecord
        Get #1, n, dbInfo
        For minors = 1 To NUM_CATS
            If Trim$(dbInfo.Minor_Category) = Minor_Popup(minors).Caption Then
                count(minors) = count(minors) + 1
            End If
        Next
    Next

    For n = 1 To frmStats.lblCount.UBound
        frmStats.lblCount(n).Visible = False
        frmStats.lblName(n).Visible = False
    Next
    
    For n = 1 To NUM_CATS
        frmStats.lblCount(n - 1).Visible = True
        frmStats.lblName(n - 1).Visible = True
        frmStats.lblCount(n - 1).Caption = Format$(count(n))
        frmStats.lblName(n - 1).Caption = Minor_Popup(n).Caption
    Next

    frmStats.Caption = "QUESTION CATEGORY STATISTICS"
    frmStats.Show vbModal
End Sub

Private Sub cmdRemember_Click()
        If Len(Trim$(txtAnswer.Text)) < 41 Then
                Current_Option = (Current_Option + 1) Mod NUM_ANSWERS
                Answer_List_Popup(Current_Option).Visible = True
                Answer_List_Popup(Current_Option).Caption = UCase$(Trim$(txtAnswer.Text))
        End If
End Sub

Private Sub Answer_List_Popup_Click(Index As Integer)
    txtAnswer.Text = Answer_List_Popup(Index).Caption
    txtAnswer.SetFocus
End Sub

Private Sub txtImageFile_Change()
    On Error Resume Next
    picImageFile.Picture = LoadPicture(App.Path & "\pictures\" & txtImageFile.Text)
End Sub

Private Sub txtImageFile_Click()
    txtImageFile.SelStart = 0
    txtImageFile.SelLength = Len(txtImageFile.Text)
End Sub

Private Sub txtSoundFile_Click()
    txtSoundFile.SelStart = 0
    txtSoundFile.SelLength = Len(txtSoundFile.Text)
End Sub

Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtQuestion.SetFocus
    End If
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtAnswer.SetFocus
    End If
End Sub

