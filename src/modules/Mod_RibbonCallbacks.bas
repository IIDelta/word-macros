Attribute VB_Name = "Mod_RibbonCallbacks"
Option Explicit

' --- Comparison & Redlines ---
Public Sub OnAction_YellowRedline(control As IRibbonControl)
    Call Mod_Redlines.MW_YellowHighlightRedline
End Sub

Public Sub OnAction_StandardRedline(control As IRibbonControl)
    Call Mod_Redlines.MW_StandardRedline
End Sub

' --- Document Cleanup ---
Public Sub OnAction_RemoveSpaces(control As IRibbonControl)
    Call Mod_Cleanup.MW_RemoveMultipleSpaces
End Sub

Public Sub OnAction_DeleteHidden(control As IRibbonControl)
    Call Mod_Cleanup.MW_DeleteHiddenText
End Sub

Public Sub OnAction_UpdateFields(control As IRibbonControl)
    Call Mod_Cleanup.MW_UpdateFields
End Sub

' --- Review Tools ---
Public Sub OnAction_ReopenComments(control As IRibbonControl)
    Call Mod_Review.MW_ReopenAllComments
End Sub