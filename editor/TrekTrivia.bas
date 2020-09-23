Attribute VB_Name = "TrekTrivia"
' ULTIMATE TREK TRIVIA
' COPYRIGHT 1997 BY JONATHAN SCOTT HARBOUR


Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long

Global Const DEBUG_MODE = True

Const BUTTON_OK = 1
Const BUTTON_CANCEL = 2
Const BUTTON_ABORT = 3
Const BUTTON_RETRY = 4
Const BUTTON_IGNORE = 5
Const BUTTON_YES = 6
Const BUTTON_NO = 7

' define the record format for the questions (currently 384 bytes)
Type GameData
    Question As String * 140
    Answer As String * 70
    Major_Category As String * 40
    Minor_Category As String * 40
    Relation As String * 40
    ImageFile As String * 12
    SoundFile As String * 12
    Filler As String * 30
End Type

Type PlayerData
    Name As String * 20
    Color As Long
    Total_Questions As Integer
    Correct_Answers As Integer
    Score As Long
    Multiplier As Integer
    Favorites(1 To 15) As Integer
End Type

Type GameStatus
    dataFile As String
    RecordLen As Long
    LastRecord As Long
    numPlayers As Integer
    CurrentPlayer As Integer
    ImageFile As String
    SoundFile As String
    CurrentQuestion As String
    Favorite As String
    CurrentCategory As String
    CurrentAnswer As String
    CurrentRecord As Long
    CurrentRelation As Integer
    CorrectAnswer As Integer
End Type

Dim iTimeDelay%
Public Players(1 To 4) As PlayerData
Public Game As GameStatus
Public NUM_MAJORS, NUM_CATS, NUM_ANSWERS As Integer

Public Sub Error_Message(msg$)
    myvalue = Int((TOTAL_ERROR_WAVES * Rnd) + 1)
    sound_file$ = "ERROR" + Trim$(Format$(myvalue)) + ".WAV"
    Debug.Print "Error sound: ", sound_file$
    Play_Sound (sound_file$)
    MsgBox msg, vbExclamation, "ERROR"
End Sub

Public Sub Play_Sound(ByVal sound_file As String)
'    frmWave.Play (App.Path & "\sounds\" & sound_file)
    ret = sndPlaySound(App.Path & "\sounds\" & sound_file, 2)
End Sub

Public Sub Button_Click()
    Play_Sound "SOUND2.WAV"
End Sub

Public Sub Button_Off()
    Play_Sound "SOUND3.WAV"
End Sub

Public Sub Button_On()
    Play_Sound "SOUND2.WAV"
End Sub

Public Function Choose_Color(ByVal num As Integer) As Long
    Choose_Color = frmColors.Choose_Color(num)
End Function

Public Function Input_Name(ByVal num As Integer) As String
    Input_Name = Trim(frmNames.Input_Name(num))
End Function

Public Function MessageBox(title$, msg1$, msg2$, accept$, decline$)
    MessageBox = frmConfirm.Confirm(title$, msg1$, msg2$, accept$, decline$, True)
End Function

Public Function AnnounceNextTurn(title$, msg1$, msg2$, accept$, decline$)
    frmConfirm.DisplayNextPlayer title$, msg1$, msg2$, accept$, decline$
    While Not frmConfirm.Finished
        DoEvents
    Wend
End Function

Public Function YouGainedPoints(ByVal points&, ByVal msg1$, ByVal msg2$)
    frmConfirm.GainPoints points, msg1$, msg2$
    While Not frmConfirm.Finished
        DoEvents
    Wend
End Function

Public Sub Delay(ByVal ms&)
    Dim start&, time&
    
    start = GetTickCount()
    time = 0
    While time < start + ms
        time = GetTickCount()
        DoEvents
    Wend
End Sub
