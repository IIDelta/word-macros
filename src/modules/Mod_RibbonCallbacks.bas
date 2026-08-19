Attribute VB_Name = "Mod_RibbonCallbacks"
Option Explicit

' --- Comparison & Redlines ---
Public Sub OnAction_YellowRedline(control As Object)
    Call Mod_Redlines.MW_YellowHighlightRedline
End Sub

Public Sub OnAction_StandardRedline(control As Object)
    Call Mod_Redlines.MW_StandardRedline
End Sub

' --- Document Cleanup ---
Public Sub OnAction_RemoveSpaces(control As Object)
    Call Mod_Cleanup.MW_RemoveMultipleSpaces
End Sub

Public Sub OnAction_DeleteHidden(control As Object)
    Call Mod_Cleanup.MW_DeleteHiddenText
End Sub

Public Sub OnAction_UpdateFields(control As Object)
    Call Mod_Cleanup.MW_UpdateFields
End Sub

' --- Review Tools ---
Public Sub OnAction_ReopenComments(control As Object)
    Call Mod_Review.MW_ReopenAllComments
End Sub
' --- COMMANDBAR FALLBACKS ---
' CommandBar buttons cannot pass the IRibbonControl parameter.
' These subroutines take no arguments and call the underlying logic.

Public Sub OnAction_YellowRedline_Fallback()
    Call OnAction_YellowRedline(Nothing)
End Sub

Public Sub OnAction_StandardRedline_Fallback()
    Call OnAction_StandardRedline(Nothing)
End Sub

Public Sub OnAction_RemoveSpaces_Fallback()
    Call OnAction_RemoveSpaces(Nothing)
End Sub

Public Sub OnAction_DeleteHidden_Fallback()
    Call OnAction_DeleteHidden(Nothing)
End Sub

Public Sub OnAction_UpdateFields_Fallback()
    Call OnAction_UpdateFields(Nothing)
End Sub

Public Sub OnAction_ReopenComments_Fallback()
    Call OnAction_ReopenComments(Nothing)
End Sub
End Sub
