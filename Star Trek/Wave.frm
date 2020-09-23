VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmWave 
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MCI.MMControl mciWave 
      Height          =   495
      Left            =   60
      TabIndex        =   0
      Top             =   855
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   873
      _Version        =   393216
      Shareable       =   -1  'True
      DeviceType      =   "WaveAudio"
      FileName        =   ""
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "WAVE PLAYER FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   432
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3768
   End
End
Attribute VB_Name = "frmWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim done As Boolean

Public Sub PlayAll(ByVal FileName$)
    mciWave.Command = "Close"
    mciWave.Wait = True
    
    On Error Resume Next
    done = False
    mciWave.FileName = FileName$
    mciWave.Command = "Open"
    mciWave.Command = "Play"
    
End Sub

Public Sub Play(sound_file As String)
    mciWave.Command = "Close"
    
    ' Force the multimedia MCI control to complete before returning
    mciWave.Wait = True
    
    On Error Resume Next
    mciWave.FileName = sound_file
    mciWave.Command = "Open"
    mciWave.Command = "Play"
End Sub

Private Sub Form_Load()
    done = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mciWave.Command = "Close"
End Sub

Private Sub mciWave_Done(NotifyCode As Integer)
    done = True
End Sub

