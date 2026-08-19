Attribute VB_Name = "Mod_Review"
Option Explicit

' ==============================================================================
' REOPEN ALL COMMENTS
' ==============================================================================
Public Sub MW_ReopenAllComments()
    Dim c As Comment
    
    For Each c In ActiveDocument.Comments
        ' Check if the comment is resolved
        If c.Done = True Then
            c.Done = False
        End If
    Next c
    
    MsgBox "All resolved comments have been reopened.", vbInformation
End Sub