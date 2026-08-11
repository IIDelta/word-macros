Option Explicit

Public Sub MW_YellowHighlightRedline()

    Const PROGRESS_INTERVAL As Long = 25
    Const DELETED_TEXT_STYLE_NAME As String = _
        "DC PleaseReview Deleted Text"
    Const HIGHLIGHT_RESTORED_DELETIONS As Boolean = True

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
            "Yellow-Highlighted Redline"
        Exit Sub
    End If

    If doc.ProtectionType <> wdNoProtection Then
        MsgBox _
            "The active document is protected. Remove protection before running this macro.", _
            vbExclamation, _
            "Yellow-Highlighted Redline"
        Exit Sub
    End If

    totalRevisions = CountAllStoryRevisions(doc)

    If totalRevisions = 0 Then
        MsgBox _
            "The active document contains no tracked changes.", _
            vbInformation, _
            "Yellow-Highlighted Redline"
        Exit Sub
    End If

    originalScreenUpdating = Application.ScreenUpdating
    originalTrackRevisions = doc.TrackRevisions
    settingsCaptured = True

    Application.ScreenUpdating = False
    Application.StatusBar = "Preparing yellow-highlighted redline..."

    ' Do not create new tracked formatting revisions while applying
    ' highlights and the custom deleted-text character style.
    doc.TrackRevisions = False

    ' Create or standardize the character style used to identify
    ' restored deletions for later Find/Replace removal.
    EnsureDeletedTextStyle doc, DELETED_TEXT_STYLE_NAME

    startTime = Timer

    ' Process the main body plus accessible Word story ranges:
    ' headers, footers, footnotes, endnotes, comments, text boxes, etc.
    For Each storyRange In doc.StoryRanges

        Set currentStoryRange = storyRange

        Do While Not currentStoryRange Is Nothing

            ' Preserve the linked-story reference before processing revisions.
            Set nextStoryRange = currentStoryRange.NextStoryRange

            ProcessStoryRevisions _
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
                PROGRESS_INTERVAL, _
                DELETED_TEXT_STYLE_NAME, _
                HIGHLIGHT_RESTORED_DELETIONS

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

        MsgBox _
            "Yellow-highlighted redline creation complete." & vbCrLf & vbCrLf & _
            "Insertions accepted and highlighted: " & insertionCount & vbCrLf & _
            "Deletions restored, highlighted, and styled: " & deletionCount & vbCrLf & _
            "Moved-to revisions treated as insertions: " & movedToCount & vbCrLf & _
            "Moved-from revisions treated as deletions: " & movedFromCount & vbCrLf & _
            "Table-cell insertions treated as insertions: " & cellInsertionCount & vbCrLf & _
            "Table-cell deletions treated as deletions: " & cellDeletionCount & vbCrLf & _
            "Other revisions accepted: " & otherAcceptedCount & vbCrLf & vbCrLf & _
            "Deleted-text style: " & DELETED_TEXT_STYLE_NAME & vbCrLf & _
            "Total revisions processed: " & completedRevisions & _
            " of " & totalRevisions & vbCrLf & _
            "Elapsed time: " & FormatDuration(elapsedSeconds), _
            vbInformation, _
            "Yellow-Highlighted Redline Complete"

    ElseIf errorNumber <> 0 Then

        MsgBox _
            "The macro stopped because of an error." & vbCrLf & vbCrLf & _
            "Error " & errorNumber & ": " & errorDescription & vbCrLf & vbCrLf & _
            "Some revisions may already have been processed. Review the " & _
            "document before saving.", _
            vbCritical, _
            "Yellow-Highlighted Redline Error"

    End If

    Exit Sub

CleanFail:

    errorNumber = Err.Number
    errorDescription = Err.Description
    processingSucceeded = False

    Resume CleanExit

End Sub


Private Sub ProcessStoryRevisions( _
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
    ByVal progressInterval As Long, _
    ByVal deletedTextStyleName As String, _
    ByVal highlightRestoredDeletions As Boolean)

    Dim rev As Revision
    Dim rngInserted As Range

    Dim revisionType As Long

    Dim elapsedSeconds As Double
    Dim estimatedRemainingSeconds As Double
    Dim percentComplete As Double

    ' This uses the confirmed Word-native sequence:
    '
    ' Insertion:
    '   Duplicate range -> Accept -> Highlight duplicate range.
    '
    ' Deletion:
    '   Duplicate range -> Apply custom style/highlight -> Reject.
    '
    ' Rejecting restores the deleted text while retaining the formatting
    ' applied to the duplicated deletion range.

    For Each rev In storyRange.Revisions

        revisionType = rev.Type

        Select Case revisionType

            ' -----------------------------------------------------------
            ' INSERTED TEXT
            ' -----------------------------------------------------------
            Case wdRevisionInsert

                Set rngInserted = rev.Range.Duplicate

                rev.Accept

                rngInserted.HighlightColorIndex = wdYellow

                insertionCount = insertionCount + 1

            ' -----------------------------------------------------------
            ' DELETED TEXT
            ' -----------------------------------------------------------
            Case wdRevisionDelete

                MarkRestoredDeletion _
                    rev, _
                    deletedTextStyleName, _
                    highlightRestoredDeletions

                deletionCount = deletionCount + 1

            ' -----------------------------------------------------------
            ' MOVED TEXT — DESTINATION
            ' Treat as an insertion.
            ' -----------------------------------------------------------
            Case wdRevisionMovedTo

                Set rngInserted = rev.Range.Duplicate

                rev.Accept

                rngInserted.HighlightColorIndex = wdYellow

                movedToCount = movedToCount + 1

            ' -----------------------------------------------------------
            ' MOVED TEXT — ORIGINAL LOCATION
            ' Treat as a deletion.
            ' -----------------------------------------------------------
            Case wdRevisionMovedFrom

                MarkRestoredDeletion _
                    rev, _
                    deletedTextStyleName, _
                    highlightRestoredDeletions

                movedFromCount = movedFromCount + 1

            ' -----------------------------------------------------------
            ' TABLE-CELL INSERTION
            ' -----------------------------------------------------------
            Case wdRevisionCellInsertion

                Set rngInserted = rev.Range.Duplicate

                rev.Accept

                rngInserted.HighlightColorIndex = wdYellow

                cellInsertionCount = cellInsertionCount + 1

            ' -----------------------------------------------------------
            ' TABLE-CELL DELETION
            ' -----------------------------------------------------------
            Case wdRevisionCellDeletion

                MarkRestoredDeletion _
                    rev, _
                    deletedTextStyleName, _
                    highlightRestoredDeletions

                cellDeletionCount = cellDeletionCount + 1

            ' -----------------------------------------------------------
            ' OTHER REVISION TYPES
            '
            ' Formatting, paragraph, table-property, section, style,
            ' conflict, and other non-text revisions are accepted.
            ' -----------------------------------------------------------
            Case Else

                rev.Accept

                otherAcceptedCount = otherAcceptedCount + 1

        End Select

        completedRevisions = completedRevisions + 1

        If completedRevisions Mod progressInterval = 0 _
            Or completedRevisions = totalRevisions Then

            elapsedSeconds = GetElapsedSeconds(startTime)

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
                "Yellow redline: " & _
                Format(percentComplete, "0.0") & "% (" & _
                completedRevisions & " of " & totalRevisions & ")" & _
                " | Elapsed: " & FormatDuration(elapsedSeconds) & _
                " | ETA: " & FormatDuration(estimatedRemainingSeconds)

            DoEvents

        End If

    Next rev

End Sub


Private Sub EnsureDeletedTextStyle( _
    ByVal doc As Document, _
    ByVal styleName As String)

    Dim deletedTextStyle As Style

    On Error Resume Next
    Set deletedTextStyle = doc.Styles(styleName)
    On Error GoTo 0

    If deletedTextStyle Is Nothing Then

        Set deletedTextStyle = doc.Styles.Add( _
            Name:=styleName, _
            Type:=wdStyleTypeCharacter)

    End If

    ' This matches the intended PleaseReview deletion appearance.
    With deletedTextStyle.Font
        .Color = wdColorRed
        .StrikeThrough = True
    End With

End Sub


Private Sub MarkRestoredDeletion( _
    ByVal rev As Revision, _
    ByVal styleName As String, _
    ByVal highlightText As Boolean)

    Dim rngDeleted As Range

    Set rngDeleted = rev.Range.Duplicate

    ' Apply the reusable character style that identifies text for
    ' later Find/Replace removal.
    rngDeleted.Style = styleName

    ' Apply yellow highlighting directly so it remains visually obvious.
    If highlightText Then
        rngDeleted.HighlightColorIndex = wdYellow
    End If

    ' Rejecting a deletion restores the original text.
    rev.Reject

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


Private Function GetElapsedSeconds( _
    ByVal startTime As Single) As Double

    Dim elapsedSeconds As Double

    elapsedSeconds = Timer - startTime

    ' Handle processing that crosses midnight.
    If elapsedSeconds < 0 Then
        elapsedSeconds = elapsedSeconds + 86400
    End If

    GetElapsedSeconds = elapsedSeconds

End Function


Private Function FormatDuration( _
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

        FormatDuration = _
            hours & "h " & _
            minutes & "m"

    ElseIf minutes > 0 Then

        FormatDuration = _
            minutes & "m " & _
            seconds & "s"

    Else

        FormatDuration = _
            seconds & "s"

    End If

End Function
