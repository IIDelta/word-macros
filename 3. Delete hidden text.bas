Attribute VB_Name = "modDeleteHiddenText" 
 
Option Explicit 
 
 
Public Sub MW_DeleteHiddenText() 
 
    Dim doc As Document 
 
    Dim storyRange As Range 
    Dim currentStoryRange As Range 
    Dim nextStoryRange As Range 
 
    Dim originalScreenUpdating As Boolean 
    Dim originalTrackRevisions As Boolean
    Dim originalShowHidden As Boolean
    Dim settingsCaptured As Boolean 
 
    Dim deletedRangeCount As Long 
 
    Dim processingSucceeded As Boolean 
    Dim errorNumber As Long 
    Dim errorDescription As String 
 
    Dim userResponse As VbMsgBoxResult 
 
    On Error GoTo CleanFail 
 
    Set doc = ActiveDocument 
 
    If doc.ReadOnly Then 
 
        MsgBox _ 
            "The active document is read-only.", _ 
            vbExclamation, _ 
            "Delete Hidden Text" 
 
        Exit Sub 
 
    End If 
 
    If doc.ProtectionType <> wdNoProtection Then 
 
        MsgBox _ 
            "The active document is protected. Remove protection before running this macro.", _ 
            vbExclamation, _ 
            "Delete Hidden Text" 
 
        Exit Sub 
 
    End If 
 
    userResponse = MsgBox( _ 
        "This macro permanently deletes ordinary text formatted as Hidden." & vbCrLf & vbCrLf & _ 
        "Run it only on a saved copy or backup. Continue?", _ 
        vbYesNo + vbExclamation + vbDefaultButton2, _ 
        "Delete Hidden Text") 
 
    If userResponse <> vbYes Then 
        Exit Sub 
    End If 
 
    ' Capture original settings
    originalScreenUpdating = Application.ScreenUpdating 
    originalTrackRevisions = doc.TrackRevisions 
    
    ' Crucial Fix: Force hidden text to be visible so the Find object doesn't ignore it
    If Not ActiveWindow Is Nothing Then
        originalShowHidden = ActiveWindow.View.ShowHiddenText
        ActiveWindow.View.ShowHiddenText = True
    End If
    
    settingsCaptured = True 
 
    Application.ScreenUpdating = False 
    Application.StatusBar = "Deleting hidden text..." 
 
    ' Prevent the deletion operation from becoming new tracked deletions. 
    doc.TrackRevisions = False 
 
    For Each storyRange In doc.StoryRanges 
 
        Set currentStoryRange = storyRange 
 
        Do While Not currentStoryRange Is Nothing 
 
            Set nextStoryRange = currentStoryRange.NextStoryRange 
 
            Application.StatusBar = _ 
                "Deleting hidden text in Story Type " & _ 
                currentStoryRange.StoryType & "..." 
 
            deletedRangeCount = deletedRangeCount + _ 
                DeleteHiddenTextFromStory(currentStoryRange) 
 
            Set currentStoryRange = nextStoryRange 
 
        Loop 
 
    Next storyRange 
 
    processingSucceeded = True 
 
CleanExit: 
 
    On Error Resume Next 
 
    ' Restore all original settings to leave the user's Word environment exactly as we found it
    If settingsCaptured Then 
 
        doc.TrackRevisions = originalTrackRevisions 
        Application.ScreenUpdating = originalScreenUpdating 
        
        If Not ActiveWindow Is Nothing Then
            ActiveWindow.View.ShowHiddenText = originalShowHidden
        End If
 
    End If 
 
    Application.StatusBar = False 
 
    On Error GoTo 0 
 
    If processingSucceeded Then 
 
        MsgBox _ 
            "Hidden text deletion complete." & vbCrLf & vbCrLf & _ 
            "Document layers processed: " & deletedRangeCount & vbCrLf & vbCrLf & _ 
            "Review the document before saving.", _ 
            vbInformation, _ 
            "Delete Hidden Text Complete" 
 
    ElseIf errorNumber <> 0 Then 
 
        MsgBox _ 
            "The macro stopped because of an error." & vbCrLf & vbCrLf & _ 
            "Error " & errorNumber & ": " & errorDescription & vbCrLf & vbCrLf & _ 
            "Some text may already have been deleted. Review the document before saving.", _ 
            vbCritical, _ 
            "Delete Hidden Text Error" 
 
    End If 
 
    Exit Sub 
 
CleanFail: 
 
    errorNumber = Err.Number 
    errorDescription = Err.Description 
    processingSucceeded = False 
 
    Resume CleanExit 
 
End Sub 
 
 
Private Function DeleteHiddenTextFromStory(ByVal storyRange As Range) As Long 
    
    ' Crucial Fix: Swapped the fragile Do While loop for the native ReplaceAll engine
    With storyRange.Find 
        .ClearFormatting 
        .Replacement.ClearFormatting 
 
        .Text = "" 
        .Replacement.Text = "" 
        .Format = True 
        .Font.Hidden = True 
        
        .Forward = True 
        .Wrap = wdFindContinue 
        
        ' Double-tap to ensure Word sees the text regardless of UI settings
        storyRange.TextRetrievalMode.IncludeHiddenText = True
 
        .Execute Replace:=wdReplaceAll
    End With 
 
    ' wdReplaceAll handles the shifting ranges perfectly, but doesn't return a precise count. 
    ' We return 1 to signify this StoryRange loop executed successfully.
    DeleteHiddenTextFromStory = 1 
 
End Function