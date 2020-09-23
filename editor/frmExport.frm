VERSION 5.00
Begin VB.Form frmExport 
   Caption         =   "Export Trivia Database"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPicture 
      Caption         =   "Skip Picture Questions"
      Height          =   195
      Left            =   3240
      TabIndex        =   18
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CheckBox chkCategory 
      Caption         =   "Include Category"
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   2700
      Width           =   1515
   End
   Begin VB.OptionButton optComma 
      Caption         =   "COMMA"
      Height          =   195
      Left            =   5220
      TabIndex        =   16
      Top             =   2700
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optTab 
      Caption         =   "TAB"
      Height          =   195
      Left            =   5220
      TabIndex        =   15
      Top             =   2400
      Width           =   915
   End
   Begin VB.CheckBox chkSubject 
      Caption         =   "Include Subject"
      Height          =   195
      Left            =   3240
      TabIndex        =   14
      Top             =   2400
      Width           =   1515
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   4620
      Width           =   3255
   End
   Begin VB.CommandButton cmdExportAll 
      Caption         =   "Export all subjects"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "Export selected subjects"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3780
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4740
      TabIndex        =   9
      Text            =   "trivia.txt"
      Top             =   1920
      Width           =   1755
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Export filename:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   8
      Top             =   1020
      Width           =   45
   End
   Begin VB.Label lblTotalQuestions 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4740
      TabIndex        =   7
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total questions:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total records:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lblRecords 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4740
      TabIndex        =   4
      Top             =   240
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total subjects:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3240
      TabIndex        =   2
      Top             =   540
      Width           =   1200
   End
   Begin VB.Label lblCount 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4740
      TabIndex        =   1
      Top             =   540
      Width           =   1755
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Subjects:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim db As GameData
Dim n, majors, totalRecords, totalSubjects
Dim category_count(1 To 50) As Integer

Private Sub Form_Load()
    Dim cat$
    For n = 0 To frmEdit.cboMainCat.ListCount
        cat$ = frmEdit.cboMainCat.List(n)
        If Len(cat$) > 0 Then
            List1.AddItem Trim$(cat$)
        End If
    Next n
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    
    totalRecords = frmEdit.dbLastRecord
    lblRecords.Caption = totalRecords
    totalSubjects = List1.ListCount
    lblCount = totalSubjects
    
    For n = 1 To totalRecords
        Get #1, n, db
        For majors = 1 To totalSubjects
            If Trim$(db.Major_Category) = frmEdit.Major_Popup(majors).Caption Then
                category_count(majors) = category_count(majors) + 1
            End If
        Next majors
    Next n
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    totalRecords = 0
    totalSubjects = 0
End Sub

Private Sub List1_Click()
    Dim total, multiple
    
    total = 0
    multiple = 0
    For n = 0 To List1.ListCount - 1
        If List1.Selected(n) Then
            total = total + category_count(n + 1)
            multiple = multiple + 1
        End If
    Next n
    
    If multiple > 1 Then
        lblSubject.Caption = "MULTIPLE SUBJECTS"
    Else
        lblSubject.Caption = List1.List(List1.ListIndex)
    End If
    
    lblTotalQuestions.Caption = total
End Sub


Private Sub cmdExport_Click()
    Dim fileNum, total, skipped, line$, sep$, majors, n
    
    If optComma Then
        sep$ = ","
    Else
        sep$ = Chr$(9)
    End If
    
    fileNum = FreeFile
    total = 0
    skipped = 0
        
    For majors = 0 To totalSubjects - 1
        If List1.Selected(majors) Then
            For n = 1 To totalRecords
                Get #1, n, db
                If Trim$(db.Major_Category) = List1.List(majors) Then
                    If Trim$(db.Question) <> "" And Trim$(db.Answer) <> "" Then
                        If chkPicture.Value = 1 And IsPictureQuestion(Trim$(db.Question)) Then
                            'MsgBox "Skipped picture question #" & n & NL & Trim$(Mid$(db.Question, 1, 50))
                            skipped = skipped + 1
                        Else
                            total = total + 1
                        End If
                    Else
                        skipped = skipped + 1
                    End If
                End If
            Next n
        End If
    Next majors
    
    Open Text1.Text For Output As fileNum
    Print #fileNum, "TRIVIA EDITOR EXPORT FILE" & sep$ & Now() & sep$ & "Copyright 1999 Jonathan S. Harbour"
    Print #fileNum, ""
    Print #fileNum, "Total records" & sep$ & total
    line$ = "Subjects exported" & sep$
    For n = 0 To List1.ListCount - 1
        If List1.Selected(n) Then
            line$ = line$ & List1.List(n) & sep$
        End If
    Next n
    Print #fileNum, line$
    Print #fileNum, " "
    
    For majors = 0 To totalSubjects - 1
        If List1.Selected(majors) Then
            For n = 1 To totalRecords
                Get #1, n, db
                If Trim$(db.Major_Category) = List1.List(majors) Then
                    If Trim$(db.Question) <> "" And Trim$(db.Answer) <> "" Then
                        If chkPicture.Value = 1 And IsPictureQuestion(Trim$(db.Question)) Then
                            'MsgBox "Skipped picture question #" & n & NL & Trim$(Mid$(db.Question, 1, 50))
                        Else
                            line$ = ""
                            If chkSubject.Value = 1 Then line$ = line$ & Trim$(db.Major_Category) & sep$
                            If chkCategory.Value = 1 Then line$ = line$ & Trim$(db.Minor_Category) & sep$
                            line$ = line$ & Trim$(db.Question) & sep$ & Trim$(db.Answer)
                            Print #fileNum, line$
                        End If
                    End If
                End If
            Next n
        End If
    Next majors
    
    Close fileNum
    MsgBox "Export of " & Format$(total, "#,##0") & " records complete." & NL & NL & _
        "Filename: " & Text1.Text & ", Size: " & Format$(FileLen(Text1.Text), "#,##0") & NL & NL & _
        "Questions skipped: " & Format$(skipped, "#,##0")

End Sub

Private Sub cmdExportAll_Click()
    Dim fileNum, total, skipped, line$, sep$
    
    If optComma Then
        sep$ = ","
    Else
        sep$ = Chr$(9)
    End If
    
    fileNum = FreeFile
    total = 0
    skipped = 0
        
    For majors = 0 To totalSubjects - 1
        For n = 1 To totalRecords
            Get #1, n, db
            If Trim$(db.Major_Category) = List1.List(majors) Then
                If Trim$(db.Question) <> "" And Trim$(db.Answer) <> "" Then
                    If chkPicture.Value = 1 And IsPictureQuestion(Trim$(db.Question)) Then
                        'MsgBox "Skipped picture question #" & n & NL & Trim$(Mid$(db.Question, 1, 50))
                        skipped = skipped + 1
                    Else
                        total = total + 1
                    End If
                Else
                    skipped = skipped + 1
                End If
            End If
        Next n
    Next majors
    
    Open Text1.Text For Output As fileNum
        Print #fileNum, "TRIVIA EDITOR EXPORT FILE" & sep$ & Now() & sep$ & "Copyright 1999 Jonathan S. Harbour"
        Print #fileNum, ""
        Print #fileNum, "Total records" & sep$ & total
        line$ = "Subjects exported" & sep$
        For n = 0 To List1.ListCount - 1
            line$ = line$ & List1.List(n) & sep$
        Next n
        Print #fileNum, line$
        Print #fileNum, " "
        
        For majors = 0 To totalSubjects - 1
            For n = 1 To totalRecords
                Get #1, n, db
                If Trim$(db.Major_Category) = List1.List(majors) Then
                    If Trim$(db.Question) <> "" And Trim$(db.Answer) <> "" Then
                        If chkPicture.Value = 1 And IsPictureQuestion(Trim$(db.Question)) Then
                            'MsgBox "Skipped picture question #" & n & NL & Trim$(Mid$(db.Question, 1, 50))
                        Else
                            line$ = ""
                            If chkSubject.Value = 1 Then line$ = line$ & Trim$(db.Major_Category) & sep$
                            If chkCategory.Value = 1 Then line$ = line$ & Trim$(db.Minor_Category) & sep$
                            line$ = line$ & Trim$(db.Question) & sep$ & Trim$(db.Answer)
                            Print #fileNum, line$
                        End If
                    End If
                End If
            Next n
        Next majors
    
    Close fileNum
    MsgBox "Export of " & Format$(total, "#,##0") & " records complete." & NL & NL & _
        "Filename: " & Text1.Text & ", Size: " & Format$(FileLen(Text1.Text), "#,##0") & NL & NL & _
        "Questions skipped: " & Format$(skipped, "#,##0")
End Sub

Public Function IsPictureQuestion(ByVal s$) As Boolean
    Dim Result As Boolean
    
    Result = False
    If InStr(1, s$, "IS SHOWN") Then Result = True
    If InStr(1, s$, "THIS PICTURE") Then Result = True
    If InStr(1, s$, "THIS SCENE") Then Result = True
    If InStr(1, s$, "THIS SHIP") Then Result = True
    If InStr(1, s$, "THIS CHARACTER") Then Result = True
    If InStr(1, s$, "THIS PERSON") Then Result = True
    If InStr(1, s$, "THIS CREW") Then Result = True
    
    IsPictureQuestion = Result
End Function

Private Sub cmdExit_Click()
    Me.Hide
    Unload Me
End Sub

Public Function NL()
    NL = Chr$(13) & Chr$(10)
End Function

