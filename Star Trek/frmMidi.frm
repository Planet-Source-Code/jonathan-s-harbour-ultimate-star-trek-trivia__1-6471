VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmMidi 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MCI.MMControl mmMidi 
      Height          =   375
      Left            =   675
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   "Sequencer"
      FileName        =   ""
   End
End
Attribute VB_Name = "frmMidi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public currentMidi&

Public Sub Play(ByVal sound&)
    
    If PlayMusic = True Then
        ' wait for previous sound to finish
        If Not mmMidi.Mode = vbMCIModeNotOpen Then
            mmMidi.Command = "Close"
        End If
        
        ' Force the multimedia MCI control to complete before returning
        mmMidi.Wait = True
        
        On Error Resume Next
        mmMidi.FileName = App.Path & "\sounds\" & "song" & sound & ".mid"
        mmMidi.Command = "Open"
        mmMidi.Command = "Play"
    
        currentMidi = sound
    End If
End Sub

Private Sub Form_Load()
    currentMidi = 1
End Sub

Public Sub mmMidi_Done(NotifyCode As Integer)
    Play_Next
End Sub

Public Sub Play_Next()
    currentMidi = currentMidi + 1
    If currentMidi > 11 Then currentMidi = 1
    Play currentMidi
End Sub

Public Sub Stop_Playing()
    mmMidi.Command = "Stop"
End Sub

