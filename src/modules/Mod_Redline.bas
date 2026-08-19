Attribute VB_Name = "Mod_Redlines"
Option Explicit

' ==============================================================================
' 1. YELLOW HIGHLIGHT REDLINE
' ==============================================================================
Public Sub MW_YellowHighlightRedline()
    Const PROGRESS_INTERVAL As Long = 50
    Const DELETED_TEXT_STYLE_NAME As String = "DC PleaseReview Deleted Text"
    Const HIGHLIGHT_RESTORED_DELETIONS As Boolean = True

    Dim doc As Document
    Dim storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim totalRevisions As Long, completedRevisions As Long
    Dim insertionCount As Long, deletionCount As Long
    Dim startTime As Single
    Dim processingSucceeded As Boolean

    On Error GoTo CleanFail

    Set doc = ActiveDocument
    If doc.ReadOnly Then MsgBox "The active document is read-only.", vbExclamation, "Yellow-Highlighted Redline": Exit Sub
    If doc.ProtectionType <> wdNoProtection Then MsgBox "The active document is protected. Remove protection before running this macro.", vbExclamation, "Yellow-Highlighted Redline": Exit Sub

    totalRevisions = CountAllStoryRevisions(doc)
    If totalRevisions = 0 Then MsgBox "The active document contains no tracked changes.", vbInformation, "Yellow-Highlighted Redline": Exit Sub

    Mod_Utilities.StartOptimization
    EnsureDeletedTextStyle doc, DELETED_TEXT_STYLE_NAME
    startTime = Timer

    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            Set nextStoryRange = currentStoryRange.NextStoryRange
            ProcessStoryRevisions currentStoryRange, totalRevisions, completedRevisions, insertionCount, deletionCount, startTime, PROGRESS_INTERVAL, DELETED_TEXT_STYLE_NAME, HIGHLIGHT_RESTORED_DELETIONS, "Yellow Redline"
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    processingSucceeded = True

CleanExit:
    Mod_Utilities.EndOptimization
    
    If processingSucceeded Then
        MsgBox "Yellow-highlighted redline creation complete." & vbCrLf & vbCrLf & "Total revisions processed: " & completedRevisions & " of " & totalRevisions & vbCrLf & "Elapsed time: " & Mod_Utilities.FormatDuration(Mod_Utilities.GetElapsedSeconds(startTime)), vbInformation, "Yellow-Highlighted Redline Complete"
    Else
        MsgBox "The macro stopped because of an error.", vbCritical, "Error"
    End If
    Exit Sub

CleanFail:
    processingSucceeded = False
    Resume CleanExit
End Sub

Private Sub ProcessStoryRevisions(ByVal storyRange As Range, ByVal totalRevisions As Long, ByRef completedRevisions As Long, ByRef insertionCount As Long, ByRef deletionCount As Long, ByVal startTime As Single, ByVal progressInterval As Long, ByVal deletedTextStyleName As String, ByVal highlightRestoredDeletions As Boolean, ByVal actionName As String)
    Dim rev As Revision, rngInserted As Range
    Dim revisionType As Long

    For Each rev In storyRange.Revisions
        revisionType = rev.Type
        Select Case revisionType
            Case wdRevisionInsert, wdRevisionMovedTo, wdRevisionCellInsertion
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.HighlightColorIndex = wdYellow
                insertionCount = insertionCount + 1
            Case wdRevisionDelete, wdRevisionMovedFrom, wdRevisionCellDeletion
                MarkRestoredDeletion rev, deletedTextStyleName, highlightRestoredDeletions
                deletionCount = deletionCount + 1
            Case Else
                rev.Accept
        End Select

        completedRevisions = completedRevisions + 1
        Mod_Utilities.UpdateProgress completedRevisions, totalRevisions, startTime, actionName, progressInterval
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
    Const PROGRESS_INTERVAL As Long = 50
    Const DELETED_TEXT_STYLE_NAME As String = "DC PleaseReview Deleted Text"
    Dim doc As Document
    Dim storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim totalRevisions As Long, completedRevisions As Long
    Dim insertionCount As Long, deletionCount As Long
    Dim startTime As Single
    Dim processingSucceeded As Boolean

    On Error GoTo CleanFail
    Set doc = ActiveDocument

    If doc.ReadOnly Then MsgBox "The active document is read-only.", vbExclamation, "Standard Redline": Exit Sub
    If doc.ProtectionType <> wdNoProtection Then MsgBox "The active document is protected. Remove protection before running this macro.", vbExclamation, "Standard Redline": Exit Sub

    totalRevisions = CountAllStoryRevisions(doc)
    If totalRevisions = 0 Then MsgBox "The active document contains no tracked changes.", vbInformation, "Standard Redline": Exit Sub

    Mod_Utilities.StartOptimization
    EnsureDeletedTextStyle doc, DELETED_TEXT_STYLE_NAME
    startTime = Timer

    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            Set nextStoryRange = currentStoryRange.NextStoryRange
            ProcessStandardRedlineStory currentStoryRange, totalRevisions, completedRevisions, insertionCount, deletionCount, startTime, PROGRESS_INTERVAL, DELETED_TEXT_STYLE_NAME
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    processingSucceeded = True

CleanExit:
    Mod_Utilities.EndOptimization
    
    If processingSucceeded Then
        MsgBox "Standard redline creation complete." & vbCrLf & vbCrLf & "Total revisions processed: " & completedRevisions & " of " & totalRevisions & vbCrLf & "Elapsed time: " & Mod_Utilities.FormatDuration(Mod_Utilities.GetElapsedSeconds(startTime)), vbInformation, "Standard Redline Complete"
    Else
        MsgBox "The macro stopped because of an error.", vbCritical, "Error"
    End If
    Exit Sub

CleanFail:
    processingSucceeded = False
    Resume CleanExit
End Sub

Private Sub ProcessStandardRedlineStory(ByVal storyRange As Range, ByVal totalRevisions As Long, ByRef completedRevisions As Long, ByRef insertionCount As Long, ByRef deletionCount As Long, ByVal startTime As Single, ByVal progressInterval As Long, ByVal deletedTextStyleName As String)
    Dim rev As Revision, rngInserted As Range, rngDeleted As Range
    Dim revisionType As Long

    For Each rev In storyRange.Revisions
        revisionType = rev.Type
        Select Case revisionType
            Case wdRevisionInsert, wdRevisionMovedTo, wdRevisionCellInsertion
                Set rngInserted = rev.Range.Duplicate
                rev.Accept
                rngInserted.Font.Color = wdColorRed
                insertionCount = insertionCount + 1
            Case wdRevisionDelete, wdRevisionMovedFrom, wdRevisionCellDeletion
                MarkRestoredDeletion rev, deletedTextStyleName, False
                deletionCount = deletionCount + 1
            Case Else
                rev.Accept
        End Select

        completedRevisions = completedRevisions + 1
        Mod_Utilities.UpdateProgress completedRevisions, totalRevisions, startTime, "Standard Redline", progressInterval
    Next rev
End Sub

' ==============================================================================
' 2.5 STANDARD REDLINE (SELECTION ONLY)
' ==============================================================================
Public Sub MW_StandardRedlineSelection()
    Const PROGRESS_INTERVAL As Long = 50
    Const DELETED_TEXT_STYLE_NAME As String = "DC PleaseReview Deleted Text"
    Dim doc As Document
    Dim selectionRange As Range
    Dim totalRevisions As Long, completedRevisions As Long
    Dim insertionCount As Long, deletionCount As Long
    Dim startTime As Single
    Dim processingSucceeded As Boolean

    On Error GoTo CleanFail
    Set doc = ActiveDocument

    If doc.ReadOnly Then MsgBox "The active document is read-only.", vbExclamation, "Standard Redline Selection": Exit Sub
    If doc.ProtectionType <> wdNoProtection Then MsgBox "The active document is protected. Remove protection before running this macro.", vbExclamation, "Standard Redline Selection": Exit Sub
    
    If Selection.Type = wdSelectionIP Then MsgBox "Please select some text first.", vbInformation, "Standard Redline Selection": Exit Sub

    Set selectionRange = Selection.Range
    totalRevisions = selectionRange.Revisions.Count
    If totalRevisions = 0 Then MsgBox "The selected text contains no tracked changes.", vbInformation, "Standard Redline Selection": Exit Sub

    Mod_Utilities.StartOptimization
    EnsureDeletedTextStyle doc, DELETED_TEXT_STYLE_NAME
    startTime = Timer

    ProcessStandardRedlineStory selectionRange, totalRevisions, completedRevisions, insertionCount, deletionCount, startTime, PROGRESS_INTERVAL, DELETED_TEXT_STYLE_NAME
    
    processingSucceeded = True

CleanExit:
    Mod_Utilities.EndOptimization
    
    If processingSucceeded Then
        MsgBox "Standard redline selection complete." & vbCrLf & vbCrLf & "Total revisions processed: " & completedRevisions & " of " & totalRevisions & vbCrLf & "Elapsed time: " & Mod_Utilities.FormatDuration(Mod_Utilities.GetElapsedSeconds(startTime)), vbInformation, "Standard Redline Complete"
    Else
        MsgBox "The macro stopped because of an error.", vbCritical, "Error"
    End If
    Exit Sub

CleanFail:
    processingSucceeded = False
    Resume CleanExit
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