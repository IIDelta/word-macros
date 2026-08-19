Attribute VB_Name = "Mod_Cleanup"
Option Explicit

' ==============================================================================
' 1. REMOVE MULTIPLE SPACES
' ==============================================================================
Public Sub MW_RemoveMultipleSpaces()
    Dim storyRange As Range
    Dim startTime As Single
    Dim processingSucceeded As Boolean
    
    On Error GoTo CleanFail
    
    If ActiveDocument.ReadOnly Or ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "Document is read-only or protected.", vbExclamation, "Remove Multiple Spaces"
        Exit Sub
    End If
    
    Mod_Utilities.StartOptimization
    startTime = Timer
    
    For Each storyRange In ActiveDocument.StoryRanges
        Dim currentStory As Range
        Set currentStory = storyRange
        Do While Not currentStory Is Nothing
            With currentStory.Find
                .ClearFormatting
                .Text = " {2,}"
                .Replacement.Text = " "
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchWildcards = True
                .Execute Replace:=wdReplaceAll
            End With
            Set currentStory = currentStory.NextStoryRange
        Loop
    Next storyRange
    
    processingSucceeded = True

CleanExit:
    Mod_Utilities.EndOptimization
    If processingSucceeded Then
        MsgBox "Multiple spaces removed." & vbCrLf & "Elapsed time: " & Mod_Utilities.FormatDuration(Mod_Utilities.GetElapsedSeconds(startTime)), vbInformation, "Complete"
    Else
        MsgBox "An error occurred.", vbCritical, "Error"
    End If
    Exit Sub
CleanFail:
    processingSucceeded = False
    Resume CleanExit
End Sub

' ==============================================================================
' 2. UPDATE ALL FIELDS
' ==============================================================================
Public Sub MW_UpdateFields()
    Dim doc As Document, storyRange As Range
    Dim toc As TableOfContents, tof As TableOfFigures, toa As TableOfAuthorities
    Set doc = ActiveDocument

    Mod_Utilities.StartOptimization

    For Each storyRange In doc.StoryRanges
        Do
            On Error Resume Next
            storyRange.Fields.Update
            On Error GoTo 0
            Set storyRange = storyRange.NextStoryRange
        Loop Until storyRange Is Nothing
    Next storyRange

    For Each toc In doc.TablesOfContents: toc.Update: Next toc
    For Each tof In doc.TablesOfFigures: tof.Update: Next tof
    For Each toa In doc.TablesOfAuthorities: toa.Update: Next toa

    Mod_Utilities.EndOptimization
    MsgBox "All document fields have been fully updated.", vbInformation
End Sub

' ==============================================================================
' 3. DELETE HIDDEN TEXT
' ==============================================================================
Public Sub MW_DeleteHiddenText()
    Dim doc As Document, storyRange As Range, currentStoryRange As Range, nextStoryRange As Range
    Dim deletedRangeCount As Long, processingSucceeded As Boolean
    Dim userResponse As VbMsgBoxResult
    Dim startTime As Single

    On Error GoTo CleanFail
    Set doc = ActiveDocument

    If doc.ReadOnly Or doc.ProtectionType <> wdNoProtection Then
        MsgBox "Document is read-only or protected.", vbExclamation, "Delete Hidden Text"
        Exit Sub
    End If

    userResponse = MsgBox("This macro permanently deletes ordinary text formatted as Hidden." & vbCrLf & vbCrLf & "Run it only on a saved copy or backup. Continue?", vbYesNo + vbExclamation + vbDefaultButton2, "Delete Hidden Text")
    If userResponse <> vbYes Then Exit Sub

    Mod_Utilities.StartOptimization
    startTime = Timer

    For Each storyRange In doc.StoryRanges
        Set currentStoryRange = storyRange
        Do While Not currentStoryRange Is Nothing
            Set nextStoryRange = currentStoryRange.NextStoryRange
            Application.StatusBar = "Deleting hidden text in Story Type " & currentStoryRange.StoryType & "..."
            DoEvents
            deletedRangeCount = deletedRangeCount + DeleteHiddenTextFromStory(currentStoryRange)
            Set currentStoryRange = nextStoryRange
        Loop
    Next storyRange
    processingSucceeded = True

CleanExit:
    Mod_Utilities.EndOptimization
    If processingSucceeded Then
        MsgBox "Hidden text deletion complete." & vbCrLf & vbCrLf & "Document layers processed: " & deletedRangeCount & vbCrLf & "Elapsed time: " & Mod_Utilities.FormatDuration(Mod_Utilities.GetElapsedSeconds(startTime)), vbInformation, "Delete Hidden Text Complete"
    Else
        MsgBox "An error occurred.", vbCritical, "Error"
    End If
    Exit Sub

CleanFail:
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