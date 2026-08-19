Attribute VB_Name = "Mod_AutoExec"
Public Sub AutoExec()
    ' This macro runs automatically when the global template loads.
    ' It creates a native Word CommandBar (toolbar) that appears on the Add-Ins tab.
    
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    
    ' 1. Delete the existing toolbar if it exists to prevent duplicates
    On Error Resume Next
    Application.CommandBars("MW Tools").Delete
    On Error GoTo 0
    
    ' 2. Create the new toolbar
    Set cb = Application.CommandBars.Add(Name:="MW Tools", Position:=msoBarTop, Temporary:=True)
    
    ' 3. Add buttons
    
    ' Button 1: Yellow Redline
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Yellow Highlight Redline"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 328 ' Highlight icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_YellowRedline_Fallback"
    
    ' Button 2: Standard Redline
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Standard Redline (Whole Doc)"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 111 ' Document compare icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_StandardRedline_Fallback"
    
    ' Button 2.5: Standard Redline Selection
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Standard Redline (Selection)"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 112 ' Alternate icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_StandardRedlineSelection_Fallback"
    
    ' Button 3: Remove Spaces
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Remove Multiple Spaces"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 239 ' Find/Replace icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_RemoveSpaces_Fallback"
    
    ' Button 4: Delete Hidden
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Delete Hidden Text"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 44 ' Show/Hide icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_DeleteHidden_Fallback"
    
    ' Button 5: Update Fields
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Update All Fields"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 1007 ' Update icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_UpdateFields_Fallback"
    
    ' Button 6: Reopen Comments
    Set btn = cb.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Reopen All Comments"
    btn.Style = msoButtonIconAndCaption
    btn.FaceId = 1589 ' Comments icon
    btn.OnAction = "Mod_RibbonCallbacks.OnAction_ReopenComments_Fallback"
    
    ' 4. Make it visible
    cb.Visible = True
End Sub

Public Sub AutoExit()
    ' Clean up when Word closes
    On Error Resume Next
    Application.CommandBars("MW Tools").Delete
    On Error GoTo 0
End Sub
