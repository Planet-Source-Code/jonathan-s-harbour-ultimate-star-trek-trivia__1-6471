VERSION 5.00
Begin VB.Form frmSubjects 
   Caption         =   "SUBJECT EDITOR"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      Height          =   555
      Left            =   3600
      TabIndex        =   6
      Top             =   3465
      Width           =   2220
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   135
      TabIndex        =   5
      Top             =   4455
      Width           =   3255
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&CANCEL"
      Height          =   555
      Left            =   3600
      TabIndex        =   4
      Top             =   4230
      Width           =   2220
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&DELETE SUBJECT"
      Height          =   555
      Left            =   3600
      TabIndex        =   3
      Top             =   1125
      Width           =   2220
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD SUBJECT"
      Height          =   555
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   2220
   End
   Begin VB.ListBox lstSubjects 
      Height          =   3960
      Left            =   135
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SUBJECTS:"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   870
   End
End
Attribute VB_Name = "frmSubjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
    lstSubjects.AddItem "<NEW>"
    lstSubjects.ListIndex = 0
    txtSubject.SelStart = 0
    txtSubject.SelLength = Len(txtSubject.Text)
    txtSubject.SetFocus
End Sub

Private Sub cmdDelete_Click()
    lstSubjects.RemoveItem lstSubjects.ListIndex
    txtSubject.Text = ""
    lstSubjects.ListIndex = 0
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim cat$, n&
    frmEdit.cboMainCat.Clear
    
    For n = 0 To NUM_MAJORS
        'Unload frmEdit.Major_Popup(n)
    Next n
    
    NUM_MAJORS = 0
    For n = 0 To lstSubjects.ListCount
        cat = lstSubjects.List(n)
        If Len(cat) > 0 Then
            NUM_MAJORS = NUM_MAJORS + 1
            frmEdit.cboMainCat.AddItem cat
            'Load frmEdit.Major_Popup(NUM_MAJORS)
            'frmEdit.Major_Popup(NUM_MAJORS).Caption = Trim$(cat)
        End If
    Next n
    'frmEdit.Major_Popup(0).Visible = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim cat$
    For n = 0 To frmEdit.cboMainCat.ListCount
        cat$ = frmEdit.cboMainCat.List(n)
        If Len(cat$) > 0 Then
            lstSubjects.AddItem Trim$(cat$)
        End If
    Next n
    If lstSubjects.ListCount > 0 Then
        lstSubjects.ListIndex = 0
    End If
End Sub

Private Sub lstSubjects_Click()
    txtSubject.Text = lstSubjects.List(lstSubjects.ListIndex)
End Sub

Private Sub txtSubject_Change()
    lstSubjects.List(lstSubjects.ListIndex) = UCase$(txtSubject.Text)
End Sub
