Attribute VB_Name = "Mod_Review"
Option Explicit

' ==============================================================================
' REOPEN ALL COMMENTS
' ==============================================================================
Public Sub MW_ReopenAllComments()
    Dim c As Comment
    Dim totalComments As Long
    Dim completedComments As Long
    Dim startTime As Single
    Dim processingSucceeded As Boolean
    
    On Error GoTo CleanFail
    
    If ActiveDocument.ReadOnly Or ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "Document is read-only or protected.", vbExclamation, "Reopen Comments"
        Exit Sub
    End If
    
    totalComments = ActiveDocument.Comments.Count
    If totalComments = 0 Then
        MsgBox "No comments found in this document.", vbInformation, "Reopen Comments"
        Exit Sub
    End If
    
    Mod_Utilities.StartOptimization
    startTime = Timer
    
    For Each c In ActiveDocument.Comments
        ' Check if the comment is resolved
        If c.Done = True Then
            c.Done = False
        End If
        
        completedComments = completedComments + 1
        Mod_Utilities.UpdateProgress completedComments, totalComments, startTime, "Reopening Comments", 25
    Next c
    
    processingSucceeded = True

CleanExit:
    Mod_Utilities.EndOptimization
    If processingSucceeded Then
        MsgBox "All resolved comments have been reopened." & vbCrLf & "Elapsed time: " & Mod_Utilities.FormatDuration(Mod_Utilities.GetElapsedSeconds(startTime)), vbInformation, "Complete"
    Else
        MsgBox "An error occurred.", vbCritical, "Error"
    End If
    Exit Sub
CleanFail:
    processingSucceeded = False
    Resume CleanExit
End Sub
' ==============================================================================
' ACCEPT CHANGES IN SELECTION (ALT + SHIFT + A)
' ==============================================================================
Public Sub MW_AcceptChangesInSelection()
    On Error GoTo CleanFail
    
    If ActiveDocument.ReadOnly Or ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "Document is read-only or protected.", vbExclamation, "Accept Changes"
        Exit Sub
    End If
    
    If Selection.Type = wdSelectionIP Then
        MsgBox "Please select some text containing tracked changes first.", vbInformation, "Accept Changes"
        Exit Sub
    End If
    
    Dim revCount As Long
    revCount = Selection.Range.Revisions.Count
    
    If revCount = 0 Then
        MsgBox "No tracked changes found in the selected text.", vbInformation, "Accept Changes"
        Exit Sub
    End If
    
    Selection.Range.Revisions.AcceptAll
    ' Optional: MsgBox revCount & " change(s) accepted.", vbInformation, "Accept Changes"

    Exit Sub
CleanFail:
    MsgBox "An error occurred while accepting changes: " & Err.Description, vbCritical, "Error"
End Sub
