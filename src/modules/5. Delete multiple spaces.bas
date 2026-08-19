Attribute VB_Name = "modRemoveMultipleSpaces"

Sub MW_RemoveMultipleSpaces()
    Dim storyRange As Range
    Dim searchRange As Range

    Application.ScreenUpdating = False

    ' Loop through all story ranges
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
            ' Only replace if NOT inside a field
            If searchRange.Fields.Count = 0 Then
                searchRange.Text = " "
            End If
            searchRange.Collapse wdCollapseEnd
        Loop
    Next storyRange

    Application.ScreenUpdating = True

    MsgBox "Multiple spaces removed (fields preserved).", vbInformation
End Sub
