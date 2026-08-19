Attribute VB_Name = "Mod_Redlines"
Option Explicit

' ==============================================================================
' 1. YELLOW HIGHLIGHT REDLINE
' ==============================================================================
Public Sub MW_YellowHighlightRedline()
    Const PROGRESS_INTERVAL As Long = 25
    Const DELETED_TEXT_STYLE_NAME As String = "DC PleaseReview Deleted Text"
    Const HIGHLIGHT_RESTORED_DELETIONS As Boolean = True

    Dim doc As Document
    Dim storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim originalScreenUpdating As Boolean, originalTrackRevisions As Boolean, settingsCaptured As Boolean
    Dim totalRevisions As Long, completedRevisions As Long
    Dim insertionCount As Long, deletionCount As Long, movedToCount As Long, movedFromCount As Long
    Dim cellInsertionCount As Long, cellDeletionCount As Long, otherAcceptedCount As Long
    Dim startTime As Single, elapsedSeconds As Double
    Dim processingSucceeded As Boolean, errorNumber As Long, errorDescription As String

    On Error GoTo CleanFail

    Set doc = ActiveDocument
    If doc.ReadOnly Then MsgBox "The active document is read-only.", vbExclamation, "Yellow-Highlighted Redline": Exit Sub
    If doc.ProtectionType <> wdNoProtection Then MsgBox "The active document is protected. Remove protection before running this macro.", vbExclamation, "Yellow-Highlighted Redline": Exit Sub

    totalRevisions = CountAllStoryRevisions(doc)
    If totalRevisions = 0 Then MsgBox "The active document contains no tracked changes.", vbInformation, "Yellow-Highlighted Redline": Exit Sub

    originalScreenUpdating = Application.ScreenUpdating
    originalTrackRevisions = doc.TrackRevisions
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.StatusBar = "Preparing yellow-highlighted redline..."
    doc.TrackRevisions = False

    EnsureDeletedTextStyle doc, DELETED_TEXT_STYLE_NAME
    startTime = Timer

    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            Set nextStoryRange = currentStoryRange.NextStoryRange
            ProcessStoryRevisions currentStoryRange, totalRevisions, completedRevisions, insertionCount, deletionCount, movedToCount, movedFromCount, cellInsertionCount, cellDeletionCount, otherAcceptedCount, startTime, PROGRESS_INTERVAL, DELETED_TEXT_STYLE_NAME, HIGHLIGHT_RESTORED_DELETIONS
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    processingSucceeded = True

CleanExit:
    On Error Resume Next
    If settingsCaptured Then
        doc.TrackRevisions = originalTrackRevisions
        Application.ScreenUpdating = originalScreenUpdating
    End If
    Application.StatusBar = False
    On Error GoTo 0

    If processingSucceeded Then
        elapsedSeconds = GetElapsedSeconds(startTime)
        MsgBox "Yellow-highlighted redline creation complete." & vbCrLf & vbCrLf & "Total revisions processed: " & completedRevisions & " of " & totalRevisions & vbCrLf & "Elapsed time: " & FormatDuration(elapsedSeconds), vbInformation, "Yellow-Highlighted Redline Complete"
    ElseIf errorNumber <> 0 Then
        MsgBox "The macro stopped because of an error." & vbCrLf & "Error " & errorNumber & ": " & errorDescription, vbCritical, "Yellow-Highlighted Redline Error"
    End If
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    processingSucceeded = False
    Resume CleanExit
End Sub

Private Sub ProcessStoryRevisions(ByVal storyRange As Range, ByVal totalRevisions As Long, ByRef completedRevisions As Long, ByRef insertionCount As Long, ByRef deletionCount As Long, ByRef movedToCount As Long, ByRef movedFromCount As Long, ByRef cellInsertionCount As Long, ByRef cellDeletionCount As Long, ByRef otherAcceptedCount As Long, ByVal startTime As Single, ByVal progressInterval As Long, ByVal deletedTextStyleName As String, ByVal highlightRestoredDeletions As Boolean)
    Dim rev As Revision, rngInserted As Range
    Dim revisionType As Long, elapsedSeconds As Double, estimatedRemainingSeconds As Double, percentComplete As Double

    For Each rev In storyRange.Revisions
        revisionType = rev.Type
        Select Case revisionType
            Case wdRevisionInsert
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.HighlightColorIndex = wdYellow
                insertionCount = insertionCount + 1
            Case wdRevisionDelete
                MarkRestoredDeletion rev, deletedTextStyleName, highlightRestoredDeletions
                deletionCount = deletionCount + 1
            Case wdRevisionMovedTo
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.HighlightColorIndex = wdYellow
                movedToCount = movedToCount + 1
            Case wdRevisionMovedFrom
                MarkRestoredDeletion rev, deletedTextStyleName, highlightRestoredDeletions
                movedFromCount = movedFromCount + 1
            Case wdRevisionCellInsertion
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.HighlightColorIndex = wdYellow
                cellInsertionCount = cellInsertionCount + 1
            Case wdRevisionCellDeletion
                MarkRestoredDeletion rev, deletedTextStyleName, highlightRestoredDeletions
                cellDeletionCount = cellDeletionCount + 1
            Case Else
                rev.Accept
                otherAcceptedCount = otherAcceptedCount + 1
        End Select

        completedRevisions = completedRevisions + 1
        If completedRevisions Mod progressInterval = 0 Or completedRevisions = totalRevisions Then
            elapsedSeconds = GetElapsedSeconds(startTime)
            If completedRevisions > 0 Then estimatedRemainingSeconds = (elapsedSeconds / completedRevisions) * (totalRevisions - completedRevisions) Else estimatedRemainingSeconds = 0
            percentComplete = (completedRevisions / totalRevisions) * 100
            Application.StatusBar = "Yellow redline: " & Format(percentComplete, "0.0") & "%"
            DoEvents
        End If
    Next rev
End Sub

Private Sub EnsureDeletedTextStyle(ByVal doc As Document, ByVal styleName As String)
    Dim deletedTextStyle As Style
    On Error Resume Next
    Set deletedTextStyle = doc.Styles(styleName)
    On Error GoTo 0
    If deletedTextStyle Is Nothing Then Set deletedTextStyle = doc.Styles.Add(Name:=styleName, Type:=wdStyleTypeCharacter)
    With deletedTextStyle.Font
        .Color = wdColorRed
        .StrikeThrough = True
    End With
End Sub

Private Sub MarkRestoredDeletion(ByVal rev As Revision, ByVal styleName As String, ByVal highlightText As Boolean)
    Dim rngDeleted As Range
    Set rngDeleted = rev.Range.Duplicate
    rngDeleted.Style = styleName
    If highlightText Then rngDeleted.HighlightColorIndex = wdYellow
    rev.Reject
End Sub

' ==============================================================================
' 2. STANDARD REDLINE
' ==============================================================================
Public Sub MW_StandardRedline()
    Const PROGRESS_INTERVAL As Long = 25
    Dim doc As Document
    Dim storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim originalScreenUpdating As Boolean, originalTrackRevisions As Boolean, settingsCaptured As Boolean
    Dim totalRevisions As Long, completedRevisions As Long
    Dim insertionCount As Long, deletionCount As Long, movedToCount As Long, movedFromCount As Long
    Dim cellInsertionCount As Long, cellDeletionCount As Long, otherAcceptedCount As Long
    Dim startTime As Single, elapsedSeconds As Double
    Dim processingSucceeded As Boolean, errorNumber As Long, errorDescription As String

    On Error GoTo CleanFail
    Set doc = ActiveDocument

    If doc.ReadOnly Then MsgBox "The active document is read-only.", vbExclamation, "Standard Redline": Exit Sub
    If doc.ProtectionType <> wdNoProtection Then MsgBox "The active document is protected. Remove protection before running this macro.", vbExclamation, "Standard Redline": Exit Sub

    totalRevisions = CountAllStoryRevisions(doc)
    If totalRevisions = 0 Then MsgBox "The active document contains no tracked changes.", vbInformation, "Standard Redline": Exit Sub

    originalScreenUpdating = Application.ScreenUpdating
    originalTrackRevisions = doc.TrackRevisions
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.StatusBar = "Preparing standard redline..."
    doc.TrackRevisions = False
    startTime = Timer

    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            Set nextStoryRange = currentStoryRange.NextStoryRange
            ProcessStandardRedlineStory currentStoryRange, totalRevisions, completedRevisions, insertionCount, deletionCount, movedToCount, movedFromCount, cellInsertionCount, cellDeletionCount, otherAcceptedCount, startTime, PROGRESS_INTERVAL
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    processingSucceeded = True

CleanExit:
    On Error Resume Next
    If settingsCaptured Then
        doc.TrackRevisions = originalTrackRevisions
        Application.ScreenUpdating = originalScreenUpdating
    End If
    Application.StatusBar = False
    On Error GoTo 0

    If processingSucceeded Then
        elapsedSeconds = GetElapsedSeconds(startTime)
        MsgBox "Standard redline creation complete." & vbCrLf & vbCrLf & "Total revisions processed: " & completedRevisions & " of " & totalRevisions & vbCrLf & "Elapsed time: " & FormatDuration(elapsedSeconds), vbInformation, "Standard Redline Complete"
    ElseIf errorNumber <> 0 Then
        MsgBox "The macro stopped because of an error." & vbCrLf & "Error " & errorNumber & ": " & errorDescription, vbCritical, "Standard Redline Error"
    End If
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    processingSucceeded = False
    Resume CleanExit
End Sub

Private Sub ProcessStandardRedlineStory(ByVal storyRange As Range, ByVal totalRevisions As Long, ByRef completedRevisions As Long, ByRef insertionCount As Long, ByRef deletionCount As Long, ByRef movedToCount As Long, ByRef movedFromCount As Long, ByRef cellInsertionCount As Long, ByRef cellDeletionCount As Long, ByRef otherAcceptedCount As Long, ByVal startTime As Single, ByVal progressInterval As Long)
    Dim rev As Revision, rngInserted As Range, rngDeleted As Range
    Dim revisionType As Long, elapsedSeconds As Double, estimatedRemainingSeconds As Double, percentComplete As Double

    For Each rev In storyRange.Revisions
        revisionType = rev.Type
        Select Case revisionType
            Case wdRevisionInsert
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.Font.Color = wdColorRed
                insertionCount = insertionCount + 1
            Case wdRevisionDelete
                Set rngDeleted = rev.Range.Duplicate
                With rngDeleted.Font
                    .Color = wdColorRed
                    .StrikeThrough = True
                End With
                rev.Reject
                deletionCount = deletionCount + 1
            Case wdRevisionMovedTo
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.Font.Color = wdColorRed
                movedToCount = movedToCount + 1
            Case wdRevisionMovedFrom
                Set rngDeleted = rev.Range.Duplicate
                With rngDeleted.Font
                    .Color = wdColorRed
                    .StrikeThrough = True
                End With
                rev.Reject
                movedFromCount = movedFromCount + 1
            Case wdRevisionCellInsertion
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.Font.Color = wdColorRed
                cellInsertionCount = cellInsertionCount + 1
            Case wdRevisionCellDeletion
                Set rngDeleted = rev.Range.Duplicate
                With rngDeleted.Font
                    .Color = wdColorRed
                    .StrikeThrough = True
                End With
                rev.Reject
                cellDeletionCount = cellDeletionCount + 1
            Case Else
                rev.Accept
                otherAcceptedCount = otherAcceptedCount + 1
        End Select

        completedRevisions = completedRevisions + 1
        If completedRevisions Mod progressInterval = 0 Or completedRevisions = totalRevisions Then
            elapsedSeconds = GetElapsedSeconds(startTime)
            percentComplete = (completedRevisions / totalRevisions) * 100
            Application.StatusBar = "Standard redline: " & Format(percentComplete, "0.0") & "%"
            DoEvents
        End If
    Next rev
End Sub

' ==============================================================================
' 3. SHARED HELPERS
' ==============================================================================
Private Function CountAllStoryRevisions(ByVal doc As Document) As Long
    Dim storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim total As Long
    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            total = total + currentStoryRange.Revisions.Count
            Set nextStoryRange = currentStoryRange.NextStoryRange
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    CountAllStoryRevisions = total
End Function

Private Function GetElapsedSeconds(ByVal startTime As Single) As Double
    Dim elapsedSeconds As Double
    elapsedSeconds = Timer - startTime
    If elapsedSeconds < 0 Then elapsedSeconds = elapsedSeconds + 86400
    GetElapsedSeconds = elapsedSeconds
End Function

Private Function FormatDuration(ByVal totalSeconds As Double) As String
    Dim roundedSeconds As Long, hours As Long, minutes As Long, seconds As Long
    roundedSeconds = CLng(totalSeconds)
    hours = roundedSeconds \ 3600
    minutes = (roundedSeconds Mod 3600) \ 60
    seconds = roundedSeconds Mod 60
    If hours > 0 Then FormatDuration = hours & "h " & minutes & "m" Else If minutes > 0 Then FormatDuration = minutes & "m " & seconds & "s" Else FormatDuration = seconds & "s"
End Function