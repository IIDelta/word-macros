Attribute VB_Name = "Mod_AutoExec"
Option Explicit

Public Sub AutoExec()
    ' Runs when the Add-in is loaded at Word startup
    CreateMWToolbar
End Sub

Public Sub AutoOpen()
    ' Runs when an existing document is opened
    CreateMWToolbar
End Sub

Public Sub AutoNew()
    ' Runs when a new blank document is created
    CreateMWToolbar
End Sub

Public Sub CreateMWToolbar()
    Dim cb As Object ' CommandBar
    Dim btn As Object ' CommandBarButton
    Dim existingBar As Object
    
    ' Check if the toolbar already exists (created by a previous event)
    On Error Resume Next
    Set existingBar = Application.CommandBars("MW Tools")
    On Error GoTo 0
    
    If Not existingBar Is Nothing Then
        ' It already exists, just ensure it's visible and exit to prevent duplicates/flashing
        existingBar.Visible = True
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' Explicitly use NormalTemplate. 
    ' If we use ThisDocument (the add-in), it's often loaded strictly as Read-Only, 
    ' which causes CommandBar modifications to fail silently.
    CustomizationContext = NormalTemplate
    
    ' Create the new toolbar as a Temporary bar (lives only for this Word session)
    Set cb = Application.CommandBars.Add(Name:="MW Tools", Position:=1, Temporary:=True)
    
    ' Button 1: Yellow Redline
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Yellow Highlight Redline"
    btn.Style = 3 ' msoButtonIconAndCaption
    btn.FaceId = 328 ' Highlight icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_YellowRedline_Fallback"
    
    ' Button 2: Standard Redline
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Standard Redline (Whole Doc)"
    btn.Style = 3
    btn.FaceId = 111 ' Document compare icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_StandardRedline_Fallback"
    
    ' Button 2.5: Standard Redline Selection
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Standard Redline (Selection)"
    btn.Style = 3
    btn.FaceId = 112 ' Alternate icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_StandardRedlineSelection_Fallback"
    
    ' Button 3: Remove Spaces
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Remove Multiple Spaces"
    btn.Style = 3
    btn.FaceId = 239 ' Find/Replace icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_RemoveSpaces_Fallback"
    
    ' Button 4: Delete Hidden
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Delete Hidden Text"
    btn.Style = 3
    btn.FaceId = 44 ' Show/Hide icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_DeleteHidden_Fallback"
    
    ' Button 5: Update Fields
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Update All Fields"
    btn.Style = 3
    btn.FaceId = 1007 ' Update icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_UpdateFields_Fallback"
    
    ' Button 6: Reopen Comments
    Set btn = cb.Controls.Add(Type:=1)
    btn.Caption = "Reopen All Comments"
    btn.Style = 3
    btn.FaceId = 1589 ' Comments icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_ReopenComments_Fallback"
    
    ' Make it visible
    cb.Visible = True
    Exit Sub

ErrorHandler:
    ' If it fails (e.g., Zotero locked the CommandBars collection), we ignore it here.
    ' AutoOpen/AutoNew will try again when a document is actually opened!
End Sub

Public Sub AutoExit()
    ' Clean up when Word closes
    On Error Resume Next
    Application.CommandBars("MW Tools").Delete
    On Error GoTo 0
End Sub
