Attribute VB_Name = "Mod_Cleanup"
Option Explicit

' ==============================================================================
' 1. REMOVE MULTIPLE SPACES
' ==============================================================================
Public Sub MW_RemoveMultipleSpaces()
    Dim storyRange As Range, searchRange As Range
    Application.ScreenUpdating = False

    For Each storyRange In ActiveDocument.StoryRanges
        Set searchRange = storyRange.Duplicate
        With searchRange.Find
            .ClearFormatting
            .Text = " {2,}"
            .Replacement.Text = " "
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = True
        End With

        Do While searchRange.Find.Execute
            If searchRange.Fields.Count = 0 Then searchRange.Text = " "
            searchRange.Collapse wdCollapseEnd
        Loop
    Next storyRange

    Application.ScreenUpdating = True
    MsgBox "Multiple spaces removed (fields preserved).", vbInformation
End Sub

' ==============================================================================
' 2. UPDATE ALL FIELDS
' ==============================================================================
Public Sub MW_UpdateFields()
    Dim doc As Document, storyRange As Range
    Dim toc As TableOfContents, tof As TableOfFigures, toa As TableOfAuthorities
    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    For Each storyRange In doc.StoryRanges
        Do
            storyRange.Fields.Update
            Set storyRange = storyRange.NextStoryRange
        Loop Until storyRange Is Nothing
    Next storyRange

    For Each toc In doc.TablesOfContents: toc.Update: Next toc
    For Each tof In doc.TablesOfFigures: tof.Update: Next tof
    For Each toa In doc.TablesOfAuthorities: toa.Update: Next toa

    Application.ScreenUpdating = True
    MsgBox "All document fields have been fully updated.", vbInformation
End Sub

' ==============================================================================
' 3. DELETE HIDDEN TEXT
' ==============================================================================
Public Sub MW_DeleteHiddenText()
    Dim doc As Document, storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim originalScreenUpdating As Boolean, originalTrackRevisions As Boolean, originalShowHidden As Boolean, settingsCaptured As Boolean
    Dim deletedRangeCount As Long, processingSucceeded As Boolean, errorNumber As Long, errorDescription As String
    Dim userResponse As VbMsgBoxResult

    On Error GoTo CleanFail
    Set doc = ActiveDocument

    If doc.ReadOnly Then MsgBox "The active document is read-only.", vbExclamation, "Delete Hidden Text": Exit Sub
    If doc.ProtectionType <> wdNoProtection Then MsgBox "The active document is protected. Remove protection before running this macro.", vbExclamation, "Delete Hidden Text": Exit Sub

    userResponse = MsgBox("This macro permanently deletes ordinary text formatted as Hidden." & vbCrLf & vbCrLf & "Run it only on a saved copy or backup. Continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Delete Hidden Text")
    If userResponse <> vbYes Then Exit Sub

    originalScreenUpdating = Application.ScreenUpdating
    originalTrackRevisions = doc.TrackRevisions
    If Not ActiveWindow Is Nothing Then
        originalShowHidden = ActiveWindow.View.ShowHiddenText
        ActiveWindow.View.ShowHiddenText = True
    End If
    settingsCaptured = True
    Application.ScreenUpdating = False
    Application.StatusBar = "Deleting hidden text..."
    doc.TrackRevisions = False

    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            Set nextStoryRange = currentStoryRange.NextStoryRange
            Application.StatusBar = "Deleting hidden text in Story Type " & currentStoryRange.StoryType & "..."
            deletedRangeCount = deletedRangeCount + DeleteHiddenTextFromStory(currentStoryRange)
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    processingSucceeded = True

CleanExit:
    On Error Resume Next
    If settingsCaptured Then
        doc.TrackRevisions = originalTrackRevisions
        Application.ScreenUpdating = originalScreenUpdating
        If Not ActiveWindow Is Nothing Then ActiveWindow.View.ShowHiddenText = originalShowHidden
    End If
    Application.StatusBar = False
    On Error GoTo 0

    If processingSucceeded Then
        MsgBox "Hidden text deletion complete." & vbCrLf & vbCrLf & "Document layers processed: " & deletedRangeCount & vbCrLf & vbCrLf & "Review the document before saving.", vbInformation, "Delete Hidden Text Complete"
    ElseIf errorNumber <> 0 Then
        MsgBox "The macro stopped because of an error." & vbCrLf & "Error " & errorNumber & ": " & errorDescription, vbCritical, "Delete Hidden Text Error"
    End If
    Exit Sub

CleanFail:
    errorNumber = Err.Number
    errorDescription = Err.Description
    processingSucceeded = False
    Resume CleanExit
End Sub

Private Function DeleteHiddenTextFromStory(ByVal storyRange As Range) As Long
    With storyRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ""
        .Replacement.Text = ""
        .Format = True
        .Font.Hidden = True
        .Forward = True
        .Wrap = wdFindContinue
        storyRange.TextRetrievalMode.IncludeHiddenText = True
        .Execute Replace:=wdReplaceAll
    End With
    DeleteHiddenTextFromStory = 1
End Function