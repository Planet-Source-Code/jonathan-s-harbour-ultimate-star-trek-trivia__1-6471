Attribute VB_Name = "HiScore"
Option Explicit
'------------------------------------------------------------
' HISCORE.BAS
'------------------------------------------------------------

Const MODAL = 1
Const MAX_HISCORES = 5
Const Section = "HiScores"
Const ENTRY = "Score"

Type tScores
    Name As String
    Score As Long
End Type

Global Hi(1 To 6) As tScores
Global Num_HiScores As Integer
Global gNewScore As Long
Global gINIFile As String
Global gDisplayOnly As Integer
Global gGameTitle As String

Sub AddScoreAndSave(ByVal NewName As String, ByVal NewScore As Long)
'------------------------------------------------------------
' Add this new score to the list of high scores and save
' everything back to the .INI file.
'------------------------------------------------------------
Dim i As Integer
Dim j As Integer
Dim temp As tScores

    ' Add the new score to the end of the Hi() array.
    Hi(Num_HiScores + 1).Name = NewName
    Hi(Num_HiScores + 1).Score = NewScore

    ' Bubble-sort the scores in descending order (highest first) ...
    For j = 1 To Num_HiScores + 1
        For i = 2 To Num_HiScores + 1
            If Hi(i).Score > Hi(i - 1).Score Then
                temp = Hi(i - 1)
                Hi(i - 1) = Hi(i)
                Hi(i) = temp
            End If
        Next
    Next

    If Num_HiScores < MAX_HISCORES Then Num_HiScores = Num_HiScores + 1

    ' Write the scores back to the .INI file.
    For i = 1 To Num_HiScores
        WriteScore gINIFile, i, Format$(Hi(i).Score) & ";" & Trim$(Hi(i).Name)
    Next

End Sub

Sub GetScores()
    Dim i As Integer
    Dim rc As Integer
    Dim pos As Integer
    Dim AString$, DefValue$, configFile$

    configFile = App.Path & "\trek.trv"
    For i = 1 To MAX_HISCORES
        AString = LoadProfile(Section, ENTRY & Format$(i), configFile)
        pos = InStr(AString, ";")
        If pos > 0 Then
            Hi(i).Score = left(AString, pos - 1)
            Hi(i).Name = Mid$(AString, pos + 1)
        End If
    Next
    Num_HiScores = MAX_HISCORES
End Sub

Function IsAHiScore(NewScore As Long) As Integer
'------------------------------------------------------------
' Returns True if NewScore is a High Score, False otherwise.
'------------------------------------------------------------
Dim i As Integer

    IsAHiScore = False
    If Num_HiScores > 0 Then
        If Num_HiScores = MAX_HISCORES And (NewScore = Hi(Num_HiScores).Score) Then
            Exit Function
        End If
    End If

    If Num_HiScores < MAX_HISCORES Then
        IsAHiScore = True
        Exit Function
    End If

    For i = 1 To Num_HiScores
        If Hi(i).Score <= NewScore Then
            IsAHiScore = True
            Exit For
        End If
    Next

End Function

Sub ShowHiScores(ByVal NewScore As Long, ByVal GameTitle As String, ByVal INIFile As String, ByVal DisplayOnly As Integer)

    gINIFile = INIFile
    gGameTitle = GameTitle
    GetScores

    If DisplayOnly Then
        gDisplayOnly = True
        frmScores.Show MODAL
    Else
        If IsAHiScore(NewScore) Then
            gNewScore = NewScore
            frmScores.Show MODAL
        End If
    End If
End Sub

Sub WriteScore(ByVal FileName As String, EntryNum As Integer, ByVal AString As String)
    Dim configFile$
    
    configFile = App.Path & "\trek.trv"
    SaveProfile Section, ENTRY & Format$(EntryNum), AString, configFile
End Sub


