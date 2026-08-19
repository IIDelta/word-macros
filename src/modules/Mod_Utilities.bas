Attribute VB_Name = "Mod_Utilities"
Option Explicit

' ==============================================================================
' CORE UTILITY ENGINE
' ==============================================================================

' Safely start optimization
Public Sub StartOptimization()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.Options.AutoFormatAsYouTypeReplaceQuotes = False
    Application.Options.AutoFormatAsYouTypeReplaceFractions = False
    Application.Options.AutoFormatAsYouTypeReplaceHyperlinks = False
    Application.Options.AutoFormatAsYouTypeReplaceOrdinals = False
    Application.Options.AutoFormatAsYouTypeReplaceSymbols = False
    ActiveDocument.TrackRevisions = False
    On Error GoTo 0
End Sub

' End optimization
Public Sub EndOptimization()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ETA and Progress Tracker
Public Sub UpdateProgress(ByVal completedRevisions As Long, ByVal totalRevisions As Long, ByVal startTime As Single, ByVal actionName As String, Optional ByVal updateInterval As Long = 50)
    If totalRevisions = 0 Then Exit Sub
    
    If completedRevisions Mod updateInterval = 0 Or completedRevisions = totalRevisions Then
        Dim elapsedSeconds As Double
        Dim estimatedRemainingSeconds As Double
        Dim percentComplete As Double
        Dim timeStr As String
        
        elapsedSeconds = Timer - startTime
        If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400 ' Handle midnight rollover
        
        If completedRevisions > 0 Then
            estimatedRemainingSeconds = (elapsedSeconds / completedRevisions) * (totalRevisions - completedRevisions)
        Else
            estimatedRemainingSeconds = 0
        End If
        
        percentComplete = (completedRevisions / totalRevisions) * 100
        
        If estimatedRemainingSeconds > 3600 Then
            timeStr = Format(estimatedRemainingSeconds / 3600, "0.0") & " hrs"
        ElseIf estimatedRemainingSeconds > 60 Then
            timeStr = Format(estimatedRemainingSeconds / 60, "0.0") & " mins"
        Else
            timeStr = Format(estimatedRemainingSeconds, "0") & " secs"
        End If
        
        Application.StatusBar = actionName & ": " & Format(percentComplete, "0.0") & "% complete. ETA: " & timeStr
        
        ' Yield to OS to allow UI updates and interaction with other docs
        DoEvents
    End If
End Sub

Public Function GetElapsedSeconds(ByVal startTime As Single) As Double
    Dim elapsedSeconds As Double
    elapsedSeconds = Timer - startTime
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400
    GetElapsedSeconds = elapsedSeconds
End Function

Public Function FormatDuration(ByVal totalSeconds As Double) As String
    Dim roundedSeconds As Long, hours As Long, minutes As Long, seconds As Long
    roundedSeconds = CLng(totalSeconds)
    hours = roundedSeconds \ 3600
    minutes = (roundedSeconds Mod 3600) \ 60
    seconds = roundedSeconds Mod 60
    If hours > 0 Then 
        FormatDuration = hours & "h " & minutes & "m" 
    ElseIf minutes > 0 Then 
        FormatDuration = minutes & "m " & seconds & "s" 
    Else 
        FormatDuration = seconds & "s"
    End If
End Function
