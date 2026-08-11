Attribute VB_Name = "modUpdateFields"

Sub MW_UpdateFields()
    Dim doc As Document
    Dim storyRange As Range
    Dim toc As TableOfContents
    Dim tof As TableOfFigures
    Dim toa As TableOfAuthorities

    Set doc = ActiveDocument

    Application.ScreenUpdating = False

    ' Update all fields in all story ranges
    For Each storyRange In doc.StoryRanges
        Do
            storyRange.Fields.Update
            Set storyRange = storyRange.NextStoryRange
        Loop Until storyRange Is Nothing
    Next storyRange

    ' Explicitly update TOCs (titles + page numbers)
    For Each toc In doc.TablesOfContents
        toc.Update
    Next toc

    ' Update Tables of Figures (includes tables, figures, equations)
    For Each tof In doc.TablesOfFigures
        tof.Update
    Next tof

    ' Update Table of Authorities (if present)
    For Each toa In doc.TablesOfAuthorities
        toa.Update
    Next toa

    Application.ScreenUpdating = True

    MsgBox "All document fields have been fully updated.", vbInformation
End Sub
