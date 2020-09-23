VERSION 5.00
Begin VB.Form frmStats 
   Caption         =   "Statistics"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   Picture         =   "Statistics.frx":0000
   ScaleHeight     =   4350
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOkay 
      Caption         =   "&Continue"
      Default         =   -1  'True
      Height          =   432
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   3645
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   25
      Left            =   4140
      TabIndex        =   52
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   25
      Left            =   2580
      TabIndex        =   51
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   24
      Left            =   4140
      TabIndex        =   50
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   24
      Left            =   2580
      TabIndex        =   49
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   23
      Left            =   4140
      TabIndex        =   47
      Top             =   2535
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   22
      Left            =   4140
      TabIndex        =   46
      Top             =   2295
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   21
      Left            =   4140
      TabIndex        =   45
      Top             =   2055
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   20
      Left            =   4140
      TabIndex        =   44
      Top             =   1815
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   19
      Left            =   4140
      TabIndex        =   43
      Top             =   1575
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   18
      Left            =   4140
      TabIndex        =   42
      Top             =   1335
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   17
      Left            =   4140
      TabIndex        =   41
      Top             =   1095
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   16
      Left            =   4140
      TabIndex        =   40
      Top             =   855
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   15
      Left            =   4140
      TabIndex        =   39
      Top             =   615
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   14
      Left            =   4140
      TabIndex        =   38
      Top             =   375
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   13
      Left            =   4140
      TabIndex        =   37
      Top             =   135
      Width           =   495
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   12
      Left            =   1665
      TabIndex        =   36
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   23
      Left            =   2580
      TabIndex        =   35
      Top             =   2535
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   22
      Left            =   2580
      TabIndex        =   34
      Top             =   2295
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   21
      Left            =   2580
      TabIndex        =   33
      Top             =   2055
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   20
      Left            =   2580
      TabIndex        =   32
      Top             =   1815
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   19
      Left            =   2580
      TabIndex        =   31
      Top             =   1575
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   18
      Left            =   2580
      TabIndex        =   30
      Top             =   1335
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   17
      Left            =   2580
      TabIndex        =   29
      Top             =   1095
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   16
      Left            =   2580
      TabIndex        =   28
      Top             =   855
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   15
      Left            =   2580
      TabIndex        =   27
      Top             =   615
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   14
      Left            =   2580
      TabIndex        =   26
      Top             =   375
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   13
      Left            =   2580
      TabIndex        =   25
      Top             =   135
      Width           =   1575
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   11
      Left            =   1680
      TabIndex        =   23
      Top             =   2760
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   2760
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   10
      Left            =   1680
      TabIndex        =   21
      Top             =   2520
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   20
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   9
      Left            =   1680
      TabIndex        =   19
      Top             =   2280
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   9
      Left            =   120
      TabIndex        =   18
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   8
      Left            =   1680
      TabIndex        =   17
      Top             =   2040
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   7
      Left            =   1680
      TabIndex        =   15
      Top             =   1800
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   6
      Left            =   1680
      TabIndex        =   13
      Top             =   1560
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   5
      Left            =   1680
      TabIndex        =   11
      Top             =   1320
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   5
      Left            =   120
      TabIndex        =   10
      Top             =   1320
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   4
      Left            =   1680
      TabIndex        =   9
      Top             =   1080
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   3
      Left            =   1680
      TabIndex        =   7
      Top             =   840
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   2
      Left            =   1680
      TabIndex        =   5
      Top             =   600
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   492
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   192
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1572
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORY NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   1575
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   192
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   492
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    Me.left = (Screen.Width - Width) / 2
    Me.top = (Screen.Height - Height) / 2
End Sub

