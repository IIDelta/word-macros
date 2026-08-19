Attribute VB_Name = "modStandardRedline"

Option Explicit


Public Sub MW_StandardRedline()

    Const PROGRESS_INTERVAL As Long = 25

    Dim doc As Document

    Dim storyRange As Range
    Dim currentStoryRange As Range
    Dim nextStoryRange As Range

    Dim originalScreenUpdating As Boolean
    Dim originalTrackRevisions As Boolean
    Dim settingsCaptured As Boolean

    Dim totalRevisions As Long
    Dim completedRevisions As Long

    Dim insertionCount As Long
    Dim deletionCount As Long
    Dim movedToCount As Long
    Dim movedFromCount As Long
    Dim cellInsertionCount As Long
    Dim cellDeletionCount As Long
    Dim otherAcceptedCount As Long

    Dim startTime As Single
    Dim elapsedSeconds As Double

    Dim processingSucceeded As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo CleanFail

    Set doc = ActiveDocument

    If doc.ReadOnly Then

        MsgBox _
            "The active document is read-only.", _
            vbExclamation, _
            "Standard Redline"

        Exit Sub

    End If

    If doc.ProtectionType <> wdNoProtection Then

        MsgBox _
            "The active document is protected. Remove protection before running this macro.", _
            vbExclamation, _
            "Standard Redline"

        Exit Sub

    End If

    totalRevisions = CountAllStoryRevisions(doc)

    If totalRevisions = 0 Then

        MsgBox _
            "The active document contains no tracked changes.", _
            vbInformation, _
            "Standard Redline"

        Exit Sub

    End If

    originalScreenUpdating = Application.ScreenUpdating
    originalTrackRevisions = doc.TrackRevisions
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.StatusBar = "Preparing standard redline..."

    ' Prevent the macro's red font and strike-through formatting
    ' from creating new tracked formatting revisions.
    doc.TrackRevisions = False

    startTime = Timer

    ' Process main text plus accessible Word story ranges:
    ' headers, footers, footnotes, endnotes, text boxes, etc.
    For Each storyRange In doc.StoryRanges

        Set currentStoryRange = storyRange

        Do While Not currentStoryRange Is Nothing

            ' Preserve next linked story before revision changes are applied.
            Set nextStoryRange = currentStoryRange.NextStoryRange

            ProcessStandardRedlineStory _
                currentStoryRange, _
                totalRevisions, _
                completedRevisions, _
                insertionCount, _
                deletionCount, _
                movedToCount, _
                movedFromCount, _
                cellInsertionCount, _
                cellDeletionCount, _
                otherAcceptedCount, _
                startTime, _
                PROGRESS_INTERVAL

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

        elapsedSeconds = GetStandardRedlineElapsedSeconds(startTime)

        MsgBox _
            "Standard redline creation complete." & vbCrLf & vbCrLf & _
            "Insertions accepted and formatted red: " & insertionCount & vbCrLf & _
            "Deletions restored and formatted red strikethrough: " & deletionCount & vbCrLf & _
            "Moved-to revisions treated as insertions: " & movedToCount & vbCrLf & _
            "Moved-from revisions treated as deletions: " & movedFromCount & vbCrLf & _
            "Table-cell insertions treated as insertions: " & cellInsertionCount & vbCrLf & _
            "Table-cell deletions treated as deletions: " & cellDeletionCount & vbCrLf & _
            "Other revisions accepted: " & otherAcceptedCount & vbCrLf & vbCrLf & _
            "Total revisions processed: " & completedRevisions & _
            " of " & totalRevisions & vbCrLf & _
            "Elapsed time: " & FormatStandardRedlineDuration(elapsedSeconds), _
            vbInformation, _
            "Standard Redline Complete"

    ElseIf errorNumber <> 0 Then

        MsgBox _
            "The macro stopped because of an error." & vbCrLf & vbCrLf & _
            "Error " & errorNumber & ": " & errorDescription & vbCrLf & vbCrLf & _
            "Some revisions may already have been processed. Review the " & _
            "document before saving.", _
            vbCritical, _
            "Standard Redline Error"

    End If

    Exit Sub

CleanFail:

    errorNumber = Err.Number
    errorDescription = Err.Description
    processingSucceeded = False

    Resume CleanExit

End Sub


Private Sub ProcessStandardRedlineStory( _
    ByVal storyRange As Range, _
    ByVal totalRevisions As Long, _
    ByRef completedRevisions As Long, _
    ByRef insertionCount As Long, _
    ByRef deletionCount As Long, _
    ByRef movedToCount As Long, _
    ByRef movedFromCount As Long, _
    ByRef cellInsertionCount As Long, _
    ByRef cellDeletionCount As Long, _
    ByRef otherAcceptedCount As Long, _
    ByVal startTime As Single, _
    ByVal progressInterval As Long)

    Dim rev As Revision

    Dim rngInserted As Range
    Dim rngDeleted As Range

    Dim revisionType As Long

    Dim elapsedSeconds As Double
    Dim estimatedRemainingSeconds As Double
    Dim percentComplete As Double

    ' Uses the confirmed Word-native sequence:
    '
    ' Insertions:
    '   Duplicate range -> Accept -> Format duplicate range.
    '
    ' Deletions:
    '   Format live deletion range -> Reject -> Word restores formatted text.
    For Each rev In storyRange.Revisions

        revisionType = rev.Type

        Select Case revisionType

            ' -----------------------------------------------------------
            ' INSERTED TEXT
            ' -----------------------------------------------------------
            Case wdRevisionInsert

                Set rngInserted = rev.Range.Duplicate

                rev.Accept

                rngInserted.Font.Color = wdColorRed

                insertionCount = insertionCount + 1

            ' -----------------------------------------------------------
            ' DELETED TEXT
            ' -----------------------------------------------------------
            Case wdRevisionDelete

                Set rngDeleted = rev.Range.Duplicate

                With rngDeleted.Font
                    .Color = wdColorRed
                    .StrikeThrough = True
                End With

                rev.Reject

                deletionCount = deletionCount + 1

            ' -----------------------------------------------------------
            ' MOVED TEXT — DESTINATION
            ' Treat as an insertion.
            ' -----------------------------------------------------------
            Case wdRevisionMovedTo

                Set rngInserted = rev.Range.Duplicate

                rev.Accept

                rngInserted.Font.Color = wdColorRed

                movedToCount = movedToCount + 1

            ' -----------------------------------------------------------
            ' MOVED TEXT — ORIGINAL LOCATION
            ' Treat as a deletion.
            ' -----------------------------------------------------------
            Case wdRevisionMovedFrom

                Set rngDeleted = rev.Range.Duplicate

                With rngDeleted.Font
                    .Color = wdColorRed
                    .StrikeThrough = True
                End With

                rev.Reject

                movedFromCount = movedFromCount + 1

            ' -----------------------------------------------------------
            ' TABLE-CELL INSERTION
            ' -----------------------------------------------------------
            Case wdRevisionCellInsertion

                Set rngInserted = rev.Range.Duplicate

                rev.Accept

                rngInserted.Font.Color = wdColorRed

                cellInsertionCount = cellInsertionCount + 1

            ' -----------------------------------------------------------
            ' TABLE-CELL DELETION
            ' -----------------------------------------------------------
            Case wdRevisionCellDeletion

                Set rngDeleted = rev.Range.Duplicate

                With rngDeleted.Font
                    .Color = wdColorRed
                    .StrikeThrough = True
                End With

                rev.Reject

                cellDeletionCount = cellDeletionCount + 1

            ' -----------------------------------------------------------
            ' ALL OTHER REVISION TYPES
            '
            ' Accept formatting, paragraph, table-property, section,
            ' style, conflict, and other non-text revisions so the output
            ' no longer contains Track Changes history.
            ' -----------------------------------------------------------
            Case Else

                rev.Accept

                otherAcceptedCount = otherAcceptedCount + 1

        End Select

        completedRevisions = completedRevisions + 1

        If completedRevisions Mod progressInterval = 0 _
            Or completedRevisions = totalRevisions Then

            elapsedSeconds = GetStandardRedlineElapsedSeconds(startTime)

            If completedRevisions > 0 Then

                estimatedRemainingSeconds = _
                    (elapsedSeconds / completedRevisions) * _
                    (totalRevisions - completedRevisions)

            Else

                estimatedRemainingSeconds = 0

            End If

            percentComplete = _
                (completedRevisions / totalRevisions) * 100

            Application.StatusBar = _
                "Standard redline: " & _
                Format(percentComplete, "0.0") & "% (" & _
                completedRevisions & " of " & totalRevisions & ")" & _
                " | Elapsed: " & _
                FormatStandardRedlineDuration(elapsedSeconds) & _
                " | ETA: " & _
                FormatStandardRedlineDuration(estimatedRemainingSeconds)

            DoEvents

        End If

    Next rev

End Sub


Private Function CountAllStoryRevisions( _
    ByVal doc As Document) As Long

    Dim storyRange As Range
    Dim currentStoryRange As Range
    Dim nextStoryRange As Range

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


Private Function GetStandardRedlineElapsedSeconds( _
    ByVal startTime As Single) As Double

    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startTime

    ' Handle a run that crosses midnight.
    If elapsedSeconds < 0 Then

        elapsedSeconds = elapsedSeconds + 86400

    End If

    GetStandardRedlineElapsedSeconds = elapsedSeconds

End Function


Private Function FormatStandardRedlineDuration( _
    ByVal totalSeconds As Double) As String

    Dim roundedSeconds As Long
    Dim hours As Long
    Dim minutes As Long
    Dim seconds As Long

    roundedSeconds = CLng(totalSeconds)

    hours = roundedSeconds \ 3600
    minutes = (roundedSeconds Mod 3600) \ 60
    seconds = roundedSeconds Mod 60

    If hours > 0 Then

        FormatStandardRedlineDuration = _
            hours & "h " & _
            minutes & "m"

    ElseIf minutes > 0 Then

        FormatStandardRedlineDuration = _
            minutes & "m " & _
            seconds & "s"

    Else

        FormatStandardRedlineDuration = _
            seconds & "s"

    End If

End Function
