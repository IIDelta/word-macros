Attribute VB_Name = "Mod_CSR_Table"
Option Explicit

' ==============================================================================
' TAKEDA CLINICAL STUDY REPORT (CSR) TABLE AUTOMATION SUITE (PRODUCTION v3.5)
' ==============================================================================
' Integrated into Medical Writing Tools Suite
' Compliant with:
'   - Takeda Formatting Rulebook Version 3.0 (07 Nov 2024) § 1, § 6
'   - Takeda Style Guide Version 6.0 (28 Mar 2025) § 4, § 9.1
'   - ICH E3 Clinical Study Reporting Guidelines
' ==============================================================================

' Central Configuration Constants
Private Const CSR_FONT_NAME As String = "Times New Roman"
Private Const CSR_FONT_SIZE_BODY As Single = 10#
Private Const CSR_FONT_SIZE_HEADER As Single = 10#
Private Const CSR_FONT_SIZE_TITLE As Single = 12#
Private Const CSR_FONT_SIZE_FOOTNOTE As Single = 10#

Private Const CSR_SPACE_AFTER_CELL As Single = 3#
Private Const CSR_SPACE_AFTER_FOOTNOTE As Single = 2#
Private Const CSR_SPACE_BEFORE_TITLE As Single = 12#
Private Const CSR_SPACE_AFTER_TITLE As Single = 4#

Private Const CSR_CELL_MARGIN_LEFT_INCHES As Single = 0.08
Private Const CSR_CELL_MARGIN_RIGHT_INCHES As Single = 0.08
Private Const CSR_CELL_MARGIN_TOP_INCHES As Single = 0.03
Private Const CSR_CELL_MARGIN_BOTTOM_INCHES As Single = 0.03

' Rulebook Section 1: Top 1.5", Left/Right/Bottom 1.0"
Private Const CSR_MARGIN_TOP_INCHES As Single = 1.5
Private Const CSR_MARGIN_BOTTOM_INCHES As Single = 1.0
Private Const CSR_MARGIN_LEFT_INCHES As Single = 1.0
Private Const CSR_MARGIN_RIGHT_INCHES As Single = 1.0
Private Const CSR_HEADER_DISTANCE_INCHES As Single = 1.0
Private Const CSR_FOOTER_DISTANCE_INCHES As Single = 0.8

Private Const CSR_PORTRAIT_MAX_WIDTH_PTS As Single = 468#
Private Const CSR_AUTO_LANDSCAPE_COL_THRESHOLD As Long = 5

' Indentation Step Constants (Points)
Private Const CSR_INDENT_LEVEL_1_PTS As Single = 12# ' ~0.17" (Level 1: 2 spaces or \li200)
Private Const CSR_INDENT_LEVEL_2_PTS As Single = 24# ' ~0.33" (Level 2: 4 spaces or \li400)
Private Const CSR_INDENT_LEVEL_3_PTS As Single = 36# ' ~0.50" (Level 3: 6 spaces or \li540)
Private Const CSR_INDENT_LEVEL_4_PTS As Single = 48# ' ~0.67" (Level 4: 8+ spaces)

' ==============================================================================
' PUBLIC ENTRY POINTS
' ==============================================================================

Public Sub CSR_Insert_RTF_Table()
    ImportAndFormatCSRTable
End Sub

Public Sub CSR_Format_Selected_Table()
    FormatCurrentSelectionAsCSRTable
End Sub

' ==============================================================================
' CORE IMPORT & MULTI-PAGE SAS STITCHING PIPELINE
' ==============================================================================

Public Sub ImportAndFormatCSRTable()
    Dim fd As Object
    Dim rtfPath As String
    Dim targetDoc As Document
    Dim stagingDoc As Document
    Dim targetRange As Range
    Dim finalTable As Table
    
    Dim requiresLandscape As Boolean
    Dim colCount As Long
    Dim totalWidthPts As Single
    
    Dim rawTableTitle As String
    Dim cleanTitleText As String
    Dim sourceNumber As String
    Dim distinctFootnotes As Collection
    Dim bylinesColl As Collection
    Dim populationName As String
    
    Dim tIdx As Long
    Dim srcTbl As Table
    Dim rIdx As Long
    Dim cIdx As Long
    Dim cellText As String
    Dim cleanCellText As String
    Dim cIndent As Single
    
    Dim numHeaderRows As Long
    Dim headerValues() As String
    Dim bodyRowsColl As Collection
    Dim rowData() As String
    Dim rowIndents() As Single
    Dim isSpacerRow As Boolean
    
    Dim p As Paragraph
    Dim pText As String
    Dim fnItem As Variant
    Dim combinedFootnotes As String
    
    On Error GoTo ErrorHandler
    
    If Documents.Count = 0 Then
        MsgBox "Please open or focus a CSR document first.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    Set targetDoc = ActiveDocument
    
    ' 1. Select RTF Output file
    Set fd = Application.FileDialog(1) ' 1 = msoFileDialogFilePicker
    With fd
        .Title = "Select SAS/R Clinical Table or Listing (RTF)"
        .Filters.Clear
        .Filters.Add "Rich Text Format (*.rtf)", "*.rtf"
        .Filters.Add "All Files (*.*)", "*.*"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        rtfPath = .SelectedItems(1)
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = 0 ' wdAlertsNone
    
    ' 2. Open RTF in hidden staging document
    Set stagingDoc = Documents.Open(FileName:=rtfPath, ConfirmConversions:=False, ReadOnly:=True, Visible:=False)
    
    ' 3. Parse Metadata (Title, Source Number, Bylines, Footnotes)
    Set distinctFootnotes = New Collection
    Set bylinesColl = New Collection
    rawTableTitle = ""
    sourceNumber = ""
    populationName = ""
    
    For Each p In stagingDoc.Paragraphs
        Dim rawPText As String
        rawPText = p.Range.Text
        pText = Trim(CleanParagraphString(rawPText))
        If Len(pText) > 0 Then
            If (InStr(1, pText, "Table ", vbTextCompare) = 1 Or InStr(1, pText, "Listing ", vbTextCompare) = 1) And Len(rawTableTitle) = 0 Then
                rawTableTitle = pText
                sourceNumber = ExtractSourceNumber(pText)
                cleanTitleText = ExtractCleanTitleWithPopulation(rawPText, populationName)
            ElseIf InStr(1, pText, "Treatment Group:", vbTextCompare) > 0 Or InStr(1, pText, "Population:", vbTextCompare) > 0 Then
                If Not CollectionContains(bylinesColl, pText) Then
                    bylinesColl.Add pText, pText
                End If
            ElseIf IsFootnoteParagraph(pText) Then
                If Not CollectionContains(distinctFootnotes, pText) Then
                    distinctFootnotes.Add pText, pText
                End If
            End If
        End If
    Next p
    
    If Len(cleanTitleText) = 0 Then
        If Len(rawTableTitle) > 0 Then
            cleanTitleText = rawTableTitle
        Else
            cleanTitleText = "Clinical Summary Table"
        End If
    End If
    
    ' Append byline to title if present (enclosing population in parentheses)
    If bylinesColl.Count > 0 Then
        Dim bLine As Variant
        For Each bLine In bylinesColl
            Dim bStr As String
            bStr = CStr(bLine)
            If InStr(1, bStr, "Population:", vbTextCompare) > 0 Then
                If Len(populationName) = 0 Then
                    Dim byPop As String
                    byPop = Trim(Mid(bStr, InStr(1, bStr, "Population:", vbTextCompare) + 11))
                    Do While Left(byPop, 1) = "(" And Right(byPop, 1) = ")"
                        byPop = Trim(Mid(byPop, 2, Len(byPop) - 2))
                    Loop
                    If Len(byPop) > 0 Then
                        populationName = "(" & byPop & ")"
                        cleanTitleText = cleanTitleText & " " & populationName
                    End If
                End If
            Else
                cleanTitleText = cleanTitleText & " (" & bStr & ")"
            End If
        Next bLine
    End If
    
    ' Parse Tables across pages (Multi-Page SAS ODS Stitching)
    Set bodyRowsColl = New Collection
    colCount = 0
    numHeaderRows = 0
    
    For tIdx = 1 To stagingDoc.Tables.Count
        Set srcTbl = stagingDoc.Tables(tIdx)
        
        ' Filter out 2-cell SAS Header/Footer Banners
        If Not IsBannerTable(srcTbl) Then
            ' First clinical table defines structure
            If colCount = 0 Then
                colCount = srcTbl.Columns.Count
                CalculateTableDimensions srcTbl, colCount, totalWidthPts
                
                ' Determine header rows (trhdr)
                numHeaderRows = 0
                For rIdx = 1 To srcTbl.Rows.Count
                    If srcTbl.Rows(rIdx).HeadingFormat Then
                        numHeaderRows = numHeaderRows + 1
                    Else
                        Exit For
                    End If
                Next rIdx
                If numHeaderRows = 0 Then numHeaderRows = 1
                
                ReDim headerValues(1 To numHeaderRows, 1 To colCount)
                For rIdx = 1 To numHeaderRows
                    For cIdx = 1 To colCount
                        On Error Resume Next
                        cellText = srcTbl.Cell(rIdx, cIdx).Range.Text
                        headerValues(rIdx, cIdx) = CleanHeaderCellString(cellText)
                        On Error GoTo ErrorHandler
                    Next cIdx
                Next rIdx
            End If
            
            ' Extract body rows (skipping 1-cell spacer rows)
            For rIdx = (numHeaderRows + 1) To srcTbl.Rows.Count
                isSpacerRow = False
                If srcTbl.Rows(rIdx).Cells.Count = 1 Then
                    isSpacerRow = True
                Else
                    Dim allBlank As Boolean: allBlank = True
                    For cIdx = 1 To srcTbl.Rows(rIdx).Cells.Count
                        cleanCellText = CleanCellString(srcTbl.Rows(rIdx).Cells(cIdx).Range.Text)
                        If Len(cleanCellText) > 0 And cleanCellText <> "~" Then
                            allBlank = False
                            Exit For
                        End If
                    Next cIdx
                    If allBlank Then isSpacerRow = True
                End If
                
                If Not isSpacerRow Then
                    ReDim rowData(1 To colCount)
                    ReDim rowIndents(1 To colCount)
                    
                    For cIdx = 1 To colCount
                        On Error Resume Next
                        cellText = srcTbl.Rows(rIdx).Cells(cIdx).Range.Text
                        Dim rawIndentPts As Single
                        rawIndentPts = srcTbl.Rows(rIdx).Cells(cIdx).Range.ParagraphFormat.LeftIndent
                        
                        ' Calculate true hierarchy indentation points (for column 1 stubs)
                        If cIdx = 1 Then
                            cIndent = DetermineIndentPoints(rawIndentPts, cellText)
                            rowData(1) = FormatRowStubSentenceCase(CleanCellString(cellText))
                        Else
                            cIndent = 0#
                            rowData(cIdx) = CleanCellString(cellText)
                        End If
                        
                        rowIndents(cIdx) = cIndent
                        On Error GoTo ErrorHandler
                    Next cIdx
                    
                    bodyRowsColl.Add Array(rowData, rowIndents)
                End If
            Next rIdx
        End If
    Next tIdx
    
    ' Close staging document
    stagingDoc.Close SaveChanges:=0
    Set stagingDoc = Nothing
    
    If colCount = 0 Or bodyRowsColl.Count = 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = -1
        MsgBox "No valid clinical data table rows could be parsed from the selected file.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    
    ' 4. Determine orientation (Rulebook § 6 / Section 1)
    If colCount >= CSR_AUTO_LANDSCAPE_COL_THRESHOLD Or totalWidthPts > CSR_PORTRAIT_MAX_WIDTH_PTS Then
        requiresLandscape = True
    Else
        requiresLandscape = False
    End If
    
    ' 5. Prepare Target Insertion Point
    targetDoc.Activate
    Set targetRange = Selection.Range
    
    ' Open Landscape Section if required
    If requiresLandscape Then
        InsertLandscapeSection targetRange
    End If
    
    ' 6. Construct the Unified Takeda CSR Table
    Dim totalTargetRows As Long
    totalTargetRows = 1 + numHeaderRows + bodyRowsColl.Count + 1 ' Title + Headers + Body + Footnote Row
    
    Set finalTable = targetDoc.Tables.Add(Range:=targetRange, NumRows:=totalTargetRows, NumColumns:=colCount)
    
    ' Populate Row 1: Title with Proper Cross-Referenceable Field Codes
    If colCount > 1 Then
        finalTable.Rows(1).Cells.Merge
    End If
    InsertCrossReferencedTitle finalTable.Cell(1, 1).Range, cleanTitleText
    
    ' Populate Header Rows
    For rIdx = 1 To numHeaderRows
        For cIdx = 1 To colCount
            finalTable.Cell(1 + rIdx, cIdx).Range.Text = headerValues(rIdx, cIdx)
        Next cIdx
        ApplyHeaderRowStyle finalTable.Rows(1 + rIdx)
    Next rIdx
    
    ' Populate Body Rows (Text only in this pass)
    Dim targetRowOffset As Long: targetRowOffset = 1 + numHeaderRows
    Dim rItem As Variant
    Dim rVals As Variant
    Dim rInds As Variant
    
    For Each rItem In bodyRowsColl
        targetRowOffset = targetRowOffset + 1
        rVals = rItem(0)
        
        For cIdx = 1 To colCount
            finalTable.Cell(targetRowOffset, cIdx).Range.Text = rVals(cIdx)
        Next cIdx
    Next rItem
    
    ' Populate Final Row: Footnotes (Merged across all columns)
    Dim footnoteRowIdx As Long: footnoteRowIdx = totalTargetRows
    If colCount > 1 Then
        finalTable.Rows(footnoteRowIdx).Cells.Merge
    End If
    
    ' Build combined footnote text starting with clean Source: line
    combinedFootnotes = "Source: " & sourceNumber & "." & vbCrLf
    For Each fnItem In distinctFootnotes
        combinedFootnotes = combinedFootnotes & CStr(fnItem) & vbCrLf
    Next fnItem
    
    ' Strip trailing carriage returns or line feeds to prevent empty trailing paragraph
    Do While Right(combinedFootnotes, 1) = vbCr Or Right(combinedFootnotes, 1) = vbLf Or Right(combinedFootnotes, 1) = " "
        combinedFootnotes = Left(combinedFootnotes, Len(combinedFootnotes) - 1)
    Loop
    
    finalTable.Cell(footnoteRowIdx, 1).Range.Text = combinedFootnotes
    ApplyFootnoteStyle finalTable.Cell(footnoteRowIdx, 1).Range
    
    ' 7. Apply Full Takeda CSR Table Rules (Borders, Padding, AutoFit, Base Row Styles)
    FormatCSRTableStructure finalTable, numHeaderRows:=numHeaderRows
    
    ' 7b. CRITICAL: Re-apply Hierarchy Indentations & Styles to Column 1
    targetRowOffset = 1 + numHeaderRows
    For Each rItem In bodyRowsColl
        targetRowOffset = targetRowOffset + 1
        rInds = rItem(1)
        Dim indVal As Single: indVal = CSng(rInds(1))
        If indVal > 0 Then
            If indVal <= CSR_INDENT_LEVEL_1_PTS Then
                ' First level indent: Apply Table:Indent style
                On Error Resume Next
                finalTable.Cell(targetRowOffset, 1).Range.Style = "Table:Indent"
                If Err.Number <> 0 Then
                    Err.Clear
                    finalTable.Cell(targetRowOffset, 1).Range.ParagraphFormat.LeftIndent = CSR_INDENT_LEVEL_1_PTS
                End If
                On Error GoTo 0
            Else
                ' Higher indent levels (Level 2+): Apply Table:Indent style with explicit LeftIndent
                On Error Resume Next
                finalTable.Cell(targetRowOffset, 1).Range.Style = "Table:Indent"
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
                finalTable.Cell(targetRowOffset, 1).Range.ParagraphFormat.LeftIndent = indVal
            End If
        End If
    Next rItem
    
    ' 8. Close Landscape section if one was opened
    If requiresLandscape Then
        Dim afterRange As Range
        Set afterRange = finalTable.Range
        afterRange.Collapse 0 ' wdCollapseEnd
        CloseLandscapeSection afterRange
    End If
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = -1
    
    MsgBox "Table imported and Takeda CSR formatting successfully applied!" & vbCrLf & _
           "Orientation: " & IIf(requiresLandscape, "Landscape", "Portrait") & vbCrLf & _
           "Columns: " & colCount & " | Data Rows: " & bodyRowsColl.Count, vbInformation, "Takeda CSR Tools"
    Exit Sub

ErrorHandler:
    If Not stagingDoc Is Nothing Then
        On Error Resume Next
        stagingDoc.Close SaveChanges:=0
    End If
    Application.ScreenUpdating = True
    Application.DisplayAlerts = -1
    MsgBox "An error occurred during RTF import: " & Err.Description, vbCritical, "Takeda CSR Tools Error"
End Sub

Public Sub FormatCurrentSelectionAsCSRTable()
    Dim tbl As Table
    If Selection.Tables.Count = 0 Then
        MsgBox "Please place your cursor inside a table first.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    
    Set tbl = Selection.Tables(1)
    FormatCSRTableStructure tbl, numHeaderRows:=1
    MsgBox "Takeda CSR formatting applied to selected table.", vbInformation, "Takeda CSR Tools"
End Sub

' ==============================================================================
' DYNAMIC CROSS-REFERENCEABLE CAPTION INSERTER
' ==============================================================================

Private Sub InsertCrossReferencedTitle(ByRef rngCell As Range, ByVal strCleanTitle As String)
    Dim doc As Document
    Dim rTarget As Range
    Dim fld1 As Field
    Dim fld2 As Field
    
    On Error GoTo FallbackPlain
    
    Set doc = rngCell.Document
    
    ' 1. Clear cell text completely and apply Caption Table style
    rngCell.Text = ""
    ApplyCaptionStyle rngCell
    
    ' 2. Place exact text with clean alphanumeric placeholder tokens:
    '    Layout: Table __SEC__.__SEQ__<Tab><strCleanTitle>
    rngCell.Text = "Table __SEC__.__SEQ__" & vbTab & strCleanTitle
    
    ' 3. Replace __SEC__ with { STYLEREF 1 \s }
    Set rTarget = rngCell.Duplicate
    With rTarget.Find
        .ClearFormatting
        .Text = "__SEC__"
        .Forward = True
        .Wrap = 0 ' wdFindStop
        .MatchCase = True
        .MatchWholeWord = False
        If .Execute Then
            Set fld1 = doc.Fields.Add(Range:=rTarget, Type:=-1, Text:="STYLEREF 1 \s", PreserveFormatting:=False)
        End If
    End With
    
    ' 4. Replace __SEQ__ with { SEQ Table \* alphabetic \s 1 }
    Set rTarget = rngCell.Duplicate
    With rTarget.Find
        .ClearFormatting
        .Text = "__SEQ__"
        .Forward = True
        .Wrap = 0 ' wdFindStop
        .MatchCase = True
        .MatchWholeWord = False
        If .Execute Then
            Set fld2 = doc.Fields.Add(Range:=rTarget, Type:=-1, Text:="SEQ Table \* alphabetic \s 1", PreserveFormatting:=False)
        End If
    End With
    
    Exit Sub

FallbackPlain:
    rngCell.Text = "Table " & vbTab & strCleanTitle
    ApplyCaptionStyle rngCell
End Sub

' ==============================================================================
' TABULATURE & HIERARCHY INDENTATION ENGINE
' ==============================================================================

Private Function DetermineIndentPoints(ByVal srcIndentPts As Single, ByVal strCellText As String) As Single
    Dim leadSpaces As Long
    leadSpaces = CountLeadingSpaces(strCellText)
    
    ' Priority 1: Monospace leading spaces (Standard SAS ASCII output)
    If leadSpaces > 0 Then
        If leadSpaces >= 8 Then
            DetermineIndentPoints = CSR_INDENT_LEVEL_4_PTS
        ElseIf leadSpaces >= 6 Then
            DetermineIndentPoints = CSR_INDENT_LEVEL_3_PTS
        ElseIf leadSpaces >= 4 Then
            DetermineIndentPoints = CSR_INDENT_LEVEL_2_PTS
        ElseIf leadSpaces >= 2 Then
            DetermineIndentPoints = CSR_INDENT_LEVEL_1_PTS
        Else
            DetermineIndentPoints = 6#
        End If
    ' Priority 2: RTF \li Twip-based left indentations
    ElseIf srcIndentPts > 0 Then
        If srcIndentPts >= 25# Then        ' \li540 = 27 pt
            DetermineIndentPoints = CSR_INDENT_LEVEL_3_PTS
        ElseIf srcIndentPts >= 15# Then    ' \li360 / \li400 = 18-20 pt
            DetermineIndentPoints = CSR_INDENT_LEVEL_2_PTS
        ElseIf srcIndentPts >= 5# Then     ' \li180 / \li200 = 9-10 pt
            DetermineIndentPoints = CSR_INDENT_LEVEL_1_PTS
        Else
            DetermineIndentPoints = srcIndentPts
        End If
    Else
        DetermineIndentPoints = 0#
    End If
End Function

Private Function CountLeadingSpaces(ByVal s As String) As Long
    Dim i As Long
    Dim c As String
    Dim count As Long
    count = 0
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If c = " " Or c = Chr(160) Then
            count = count + 1
        ElseIf c = vbTab Then
            count = count + 4
        Else
            Exit For
        End If
    Next i
    CountLeadingSpaces = count
End Function

' ==============================================================================
' SECTION & ORIENTATION HELPERS (Rulebook § 1 Compliant)
' ==============================================================================

Public Sub CalculateTableDimensions(ByVal tbl As Table, ByRef outCols As Long, ByRef outWidthPts As Single)
    Dim c As Column
    outCols = tbl.Columns.Count
    outWidthPts = 0#
    On Error Resume Next
    For Each c In tbl.Columns
        outWidthPts = outWidthPts + c.Width
    Next c
    On Error GoTo 0
End Sub

Public Sub InsertLandscapeSection(ByRef targetRange As Range)
    Dim newSec As Section
    targetRange.Collapse 0 ' wdCollapseEnd
    targetRange.InsertBreak 3 ' wdSectionBreakNextPage
    
    Set newSec = targetRange.Document.Sections(targetRange.Sections(1).Index)
    With newSec
        .PageSetup.Orientation = 1 ' wdOrientLandscape
        ApplyTakedaPageMargins .PageSetup
        UnlinkHeadersAndFooters newSec
    End With
    Set targetRange = newSec.Range
    targetRange.Collapse 1 ' wdCollapseStart
End Sub

Public Sub CloseLandscapeSection(ByRef afterRange As Range)
    Dim returnSec As Section
    afterRange.Collapse 0 ' wdCollapseEnd
    afterRange.InsertBreak 3 ' wdSectionBreakNextPage
    
    Set returnSec = afterRange.Document.Sections(afterRange.Sections(1).Index)
    With returnSec
        .PageSetup.Orientation = 0 ' wdOrientPortrait
        ApplyTakedaPageMargins .PageSetup
        UnlinkHeadersAndFooters returnSec
    End With
End Sub

Private Sub ApplyTakedaPageMargins(ByRef ps As PageSetup)
    With ps
        .TopMargin = InchesToPoints(CSR_MARGIN_TOP_INCHES)
        .BottomMargin = InchesToPoints(CSR_MARGIN_BOTTOM_INCHES)
        .LeftMargin = InchesToPoints(CSR_MARGIN_LEFT_INCHES)
        .RightMargin = InchesToPoints(CSR_MARGIN_RIGHT_INCHES)
        .HeaderDistance = InchesToPoints(CSR_HEADER_DISTANCE_INCHES)
        .FooterDistance = InchesToPoints(CSR_FOOTER_DISTANCE_INCHES)
    End With
End Sub

Private Sub UnlinkHeadersAndFooters(ByRef sec As Section)
    Dim hf As HeaderFooter
    For Each hf In sec.Headers
        hf.LinkToPrevious = False
    Next hf
    For Each hf In sec.Footers
        hf.LinkToPrevious = False
    Next hf
End Sub

' ==============================================================================
' TABLE FORMATTING & TAKEDA 3-LINE BORDER RULES
' ==============================================================================

Public Sub FormatCSRTableStructure(ByVal tbl As Table, ByVal numHeaderRows As Long)
    Dim totalRows As Long
    Dim rIdx As Long
    Dim footnoteRowIdx As Long
    
    totalRows = tbl.Rows.Count
    If totalRows < (2 + numHeaderRows) Then Exit Sub
    footnoteRowIdx = totalRows
    
    ' Global Table Geometry
    With tbl
        .Rows.Alignment = 0 ' wdAlignRowLeft
        .Rows.LeftIndent = 0#
        .Rows.AllowBreakAcrossPages = False
        
        .LeftPadding = InchesToPoints(CSR_CELL_MARGIN_LEFT_INCHES)
        .RightPadding = InchesToPoints(CSR_CELL_MARGIN_RIGHT_INCHES)
        .TopPadding = InchesToPoints(CSR_CELL_MARGIN_TOP_INCHES)
        .BottomPadding = InchesToPoints(CSR_CELL_MARGIN_BOTTOM_INCHES)
        
        .AllowAutoFit = True
        .AutoFitBehavior 2 ' wdAutoFitWindow
    End With
    
    ' Strip All Default Borders
    tbl.Borders.Enable = False
    tbl.Borders(-1).LineStyle = 0 ' wdBorderTop = wdLineStyleNone
    tbl.Borders(-3).LineStyle = 0 ' wdBorderBottom
    tbl.Borders(-2).LineStyle = 0 ' wdBorderLeft
    tbl.Borders(-4).LineStyle = 0 ' wdBorderRight
    tbl.Borders(-5).LineStyle = 0 ' wdBorderHorizontal
    tbl.Borders(-6).LineStyle = 0 ' wdBorderVertical
    
    ' Set Repeat Header Rows on Title and Header Rows
    tbl.Rows(1).HeadingFormat = True
    For rIdx = 1 To numHeaderRows
        tbl.Rows(1 + rIdx).HeadingFormat = True
    Next rIdx
    
    ' Apply Base Body Styling to all data rows
    For rIdx = (2 + numHeaderRows) To (totalRows - 1)
        ApplyBodyRowStyle tbl.Rows(rIdx)
        
        ' Format Row Stub Column 1 in Sentence Case if applicable (e.g. for pre-existing tables)
        Dim stubCell As Cell
        Set stubCell = tbl.Cell(rIdx, 1)
        Dim stubText As String
        stubText = stubCell.Range.Text
        If Len(stubText) > 2 Then
            ' Word cell text ends with Chr(13) & Chr(7)
            stubText = Left(stubText, Len(stubText) - 2)
            Dim formattedStub As String
            formattedStub = FormatRowStubSentenceCase(stubText)
            If formattedStub <> stubText Then
                stubCell.Range.Text = formattedStub
            End If
        End If
        
        ' Check if Column 1 has existing indent to apply Table:Indent style
        If stubCell.Range.ParagraphFormat.LeftIndent > 0 Then
            If stubCell.Range.ParagraphFormat.LeftIndent <= CSR_INDENT_LEVEL_1_PTS Then
                On Error Resume Next
                stubCell.Range.Style = "Table:Indent"
                If Err.Number <> 0 Then Err.Clear
                On Error GoTo 0
            End If
        End If
    Next rIdx
    
    ' -------------------------------------------------------------------------
    ' Takeda 3-Line Border Rules (wdLineWidth050pt = 4):
    ' -------------------------------------------------------------------------
    ' 1. Below Title Row
    With tbl.Rows(1).Borders(-3) ' wdBorderBottom
        .LineStyle = 1           ' wdLineStyleSingle
        .LineWidth = 4           ' wdLineWidth050pt
        .Color = 0               ' wdColorBlack
    End With
    
    ' 2. Below Header Row(s)
    With tbl.Rows(1 + numHeaderRows).Borders(-3)
        .LineStyle = 1
        .LineWidth = 4
        .Color = 0
    End With
    
    ' 3. Above Footnote Row
    With tbl.Rows(footnoteRowIdx).Borders(-1) ' wdBorderTop
        .LineStyle = 1
        .LineWidth = 4
        .Color = 0
    End With
    
    ' -------------------------------------------------------------------------
    ' Automated Regulatory Refinements:
    ' -------------------------------------------------------------------------
    ' 1. Regulatory symbols (>= -> ≥, <= -> ≤) and equals spacing
    ReplaceSymbolsInTable tbl
    RemoveEqualsSpacesInTable tbl
    
    ' 2. Regulatory terminology standardization: replace patient(s) -> participant(s)
    ReplacePatientInTable tbl
    
    ' 3. Convert numeric/letter footnote callouts ([1]..[9], [a]..[z]) to Takeda regulatory superscript letters (a..z)
    ConvertFootnoteCalloutsToSuperscriptLetters tbl
    
    ' 4. Regulatory citation styling: color Table/Listing/Figure mentions in footnotes as Blue (#0000FF)
    ColorCitationsInFootnotes tbl
End Sub

Private Sub ReplacePatientInTable(ByVal tbl As Table)
    Dim pairs As Variant
    Dim i As Long
    Dim srcWord As String
    Dim dstWord As String
    Dim r As Range
    
    pairs = Array( _
        "Patients'", "Participants'", _
        "patients'", "participants'", _
        "Patient's", "Participant's", _
        "patient's", "participant's", _
        "PATIENTS", "PARTICIPANTS", _
        "PATIENT", "PARTICIPANT", _
        "Patients", "Participants", _
        "patients", "participants", _
        "Patient", "Participant", _
        "patient", "participant" _
    )
    
    For i = LBound(pairs) To UBound(pairs) Step 2
        srcWord = CStr(pairs(i))
        dstWord = CStr(pairs(i + 1))
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = srcWord
            .Replacement.Text = dstWord
            .MatchCase = True
            .MatchWholeWord = True
            .Forward = True
            .Wrap = 0 ' wdFindStop
            .Execute Replace:=2 ' wdReplaceAll
        End With
    Next i
End Sub

Private Sub ConvertFootnoteCalloutsToSuperscriptLetters(ByVal tbl As Table)
    Dim i As Long
    Dim numStr As String
    Dim letStr As String
    Dim r As Range
    
    ' 1. Convert numeric bracket callouts with leading space " [1]".." [9]" -> "a".."i" (superscript)
    For i = 1 To 9
        numStr = " [" & CStr(i) & "]"
        letStr = Chr(96 + i)
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = numStr
            .Replacement.Text = letStr
            .Replacement.Font.Superscript = True
            .Forward = True
            .Wrap = 0
            .MatchCase = True
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
    Next i
    
    ' 2. Convert remaining numeric bracket callouts "[1]".."[9]" -> "a".."i" (superscript)
    For i = 1 To 9
        numStr = "[" & CStr(i) & "]"
        letStr = Chr(96 + i)
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = numStr
            .Replacement.Text = letStr
            .Replacement.Font.Superscript = True
            .Forward = True
            .Wrap = 0
            .MatchCase = True
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
    Next i
    
    ' 3. Convert letter bracket callouts with leading space " [a]".." [z]" -> "a".."z" (superscript)
    For i = 1 To 26
        numStr = " [" & Chr(96 + i) & "]"
        letStr = Chr(96 + i)
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = numStr
            .Replacement.Text = letStr
            .Replacement.Font.Superscript = True
            .Forward = True
            .Wrap = 0
            .MatchCase = True
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
    Next i
    
    ' 4. Convert remaining letter bracket callouts "[a]".."[z]" -> "a".."z" (superscript)
    For i = 1 To 26
        numStr = "[" & Chr(96 + i) & "]"
        letStr = Chr(96 + i)
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = numStr
            .Replacement.Text = letStr
            .Replacement.Font.Superscript = True
            .Forward = True
            .Wrap = 0
            .MatchCase = True
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
    Next i
End Sub

Private Sub ColorCitationsInFootnotes(ByVal tbl As Table)
    Dim footnoteRowIdx As Long
    Dim fnRange As Range
    Dim pType As Variant
    Dim rFind As Range
    Dim searchWord As String
    Dim nextChar As String
    
    footnoteRowIdx = tbl.Rows.Count
    If footnoteRowIdx < 2 Then Exit Sub
    
    For Each pType In Array("Table ", "Listing ", "Figure ")
        searchWord = CStr(pType)
        Set fnRange = tbl.Cell(footnoteRowIdx, 1).Range
        Set rFind = fnRange.Duplicate
        
        With rFind.Find
            .ClearFormatting
            .Text = searchWord
            .Forward = True
            .Wrap = 0 ' wdFindStop
            .MatchCase = True
            .MatchWholeWord = False
            
            Do While .Execute
                If rFind.Start >= fnRange.End Then Exit Do
                
                Do While rFind.End < fnRange.End
                    nextChar = fnRange.Document.Range(rFind.End, rFind.End + 1).Text
                    If (nextChar >= "0" And nextChar <= "9") Or _
                       (nextChar >= "a" And nextChar <= "z") Or _
                       (nextChar >= "A" And nextChar <= "Z") Or _
                       nextChar = "." Or nextChar = "_" Or nextChar = "-" Then
                        rFind.End = rFind.End + 1
                    Else
                        Exit Do
                    End If
                Loop
                
                Do While Right(rFind.Text, 1) = "." Or Right(rFind.Text, 1) = " "
                    rFind.End = rFind.End - 1
                Loop
                
                rFind.Font.Color = RGB(0, 0, 255)
                
                rFind.Collapse 0 ' wdCollapseEnd
                If rFind.Start >= fnRange.End Then Exit Do
            Loop
        End With
    Next pType
End Sub

Private Sub ApplyCaptionStyle(ByRef rng As Range)
    On Error Resume Next
    rng.Style = "Caption Table"
    If Err.Number <> 0 Then
        Err.Clear
        rng.Style = "Caption"
        If Err.Number <> 0 Then
            Err.Clear
            rng.Style = "Normal"
        End If
    End If
    On Error GoTo 0
    
    With rng
        .Font.Name = CSR_FONT_NAME
        .Font.Size = CSR_FONT_SIZE_TITLE
        .Font.Bold = True
        .ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        .ParagraphFormat.SpaceBefore = CSR_SPACE_BEFORE_TITLE
        .ParagraphFormat.SpaceAfter = CSR_SPACE_AFTER_TITLE
        .ParagraphFormat.KeepWithNext = True
    End With
End Sub

Private Sub ApplyHeaderRowStyle(ByVal r As Row)
    Dim c As Cell
    On Error Resume Next
    r.Range.Style = "Table:Header"
    If Err.Number <> 0 Then
        Err.Clear
        With r.Range
            .Font.Name = CSR_FONT_NAME
            .Font.Size = CSR_FONT_SIZE_HEADER
            .Font.Bold = True
            .ParagraphFormat.SpaceBefore = 0#
            .ParagraphFormat.SpaceAfter = CSR_SPACE_AFTER_CELL
        End With
    End If
    On Error GoTo 0
    
    r.Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
    For Each c In r.Cells
        c.VerticalAlignment = 3 ' wdCellAlignVerticalBottom per Rulebook
    Next c
End Sub

Private Sub ApplyBodyRowStyle(ByVal r As Row)
    Dim cIdx As Long
    Dim cellRng As Range
    
    ' 1. Stub column (Column 1) -> Style "Table:Text", Left-aligned
    If r.Cells.Count >= 1 Then
        Set cellRng = r.Cells(1).Range
        On Error Resume Next
        cellRng.Style = "Table:Text"
        If Err.Number <> 0 Then
            Err.Clear
            With cellRng
                .Font.Name = CSR_FONT_NAME
                .Font.Size = CSR_FONT_SIZE_BODY
                .Font.Bold = False
                .ParagraphFormat.SpaceBefore = 0#
                .ParagraphFormat.SpaceAfter = CSR_SPACE_AFTER_CELL
                .ParagraphFormat.WidowControl = True
            End With
        End If
        On Error GoTo 0
        cellRng.ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        r.Cells(1).VerticalAlignment = 1     ' wdCellAlignVerticalCenter
    End If
    
    ' 2. Data columns (Columns 2+) -> Style "Table:Numbers"
    ' Do not override paragraph alignment to centered; allow Table:Numbers to govern alignment
    ' and preserve decimal tab stops so writers can position decimals under headers via the Ruler.
    For cIdx = 2 To r.Cells.Count
        Set cellRng = r.Cells(cIdx).Range
        On Error Resume Next
        cellRng.Style = "Table:Numbers"
        If Err.Number <> 0 Then
            Err.Clear
            cellRng.Style = "Table:Text"
            If Err.Number <> 0 Then
                Err.Clear
                With cellRng
                    .Font.Name = CSR_FONT_NAME
                    .Font.Size = CSR_FONT_SIZE_BODY
                    .Font.Bold = False
                    .ParagraphFormat.SpaceBefore = 0#
                    .ParagraphFormat.SpaceAfter = CSR_SPACE_AFTER_CELL
                    .ParagraphFormat.WidowControl = True
                End With
            End If
        End If
        On Error GoTo 0
        r.Cells(cIdx).VerticalAlignment = 1   ' wdCellAlignVerticalCenter
    Next cIdx
End Sub

Private Sub ApplyFootnoteStyle(ByRef rng As Range)
    On Error Resume Next
    rng.Style = "Footnote Text"
    If Err.Number <> 0 Then
        Err.Clear
        rng.Style = "Table:Text"
    End If
    On Error GoTo 0
    
    With rng
        .Font.Name = CSR_FONT_NAME
        .Font.Size = CSR_FONT_SIZE_FOOTNOTE
        .Font.Bold = False
        .ParagraphFormat.Alignment = 0 ' wdAlignParagraphLeft
        .ParagraphFormat.SpaceBefore = 3#
        .ParagraphFormat.SpaceAfter = CSR_SPACE_AFTER_FOOTNOTE
    End With
End Sub

' ==============================================================================
' METADATA, STRING & PARSING UTILITIES
' ==============================================================================

Private Function IsBannerTable(ByVal tbl As Table) As Boolean
    Dim txt As String
    If tbl.Rows.Count = 1 And tbl.Columns.Count = 2 Then
        txt = tbl.Range.Text
        If InStr(1, txt, "Page ", vbTextCompare) > 0 Or _
           InStr(1, txt, "Protocol:", vbTextCompare) > 0 Or _
           InStr(1, txt, "Program:", vbTextCompare) > 0 Or _
           InStr(1, txt, "Executed:", vbTextCompare) > 0 Then
            IsBannerTable = True
            Exit Function
        End If
    End If
    IsBannerTable = False
End Function

Private Function ExtractSourceNumber(ByVal titlePara As String) As String
    Dim parts() As String
    Dim cleanP As String
    cleanP = CleanParagraphString(titlePara)
    parts = Split(cleanP, " ")
    If UBound(parts) >= 1 Then
        ExtractSourceNumber = parts(0) & " " & parts(1)
    Else
        ExtractSourceNumber = "Table"
    End If
End Function

Private Function ExtractCleanTitle(ByVal titlePara As String) As String
    Dim parts() As String
    Dim cleanP As String
    Dim i As Long
    Dim res As String
    
    cleanP = CleanParagraphString(titlePara)
    parts = Split(cleanP, " ")
    If UBound(parts) >= 2 Then
        res = ""
        For i = 2 To UBound(parts)
            res = res & parts(i) & " "
        Next i
        ExtractCleanTitle = Trim(res)
    Else
        ExtractCleanTitle = cleanP
    End If
End Function

Private Function IsFootnoteParagraph(ByVal pText As String) As Boolean
    Dim t As String
    t = Trim(pText)
    If Len(t) = 0 Then
        IsFootnoteParagraph = False
    ElseIf InStr(1, t, "Note:", vbTextCompare) = 1 Or _
           InStr(1, t, "[1]", vbTextCompare) = 1 Or _
           InStr(1, t, "[2]", vbTextCompare) = 1 Or _
           InStr(1, t, "[3]", vbTextCompare) = 1 Or _
           InStr(1, t, "[4]", vbTextCompare) = 1 Or _
           InStr(1, t, "*", vbTextCompare) = 1 Or _
           InStr(1, t, "Source:", vbTextCompare) = 1 Then
        IsFootnoteParagraph = True
    Else
        IsFootnoteParagraph = False
    End If
End Function

Private Function CleanParagraphString(ByVal s As String) As String
    Dim res As String
    res = Replace(s, vbCr, " ")
    res = Replace(res, vbLf, " ")
    res = Replace(res, vbTab, " ")
    res = Replace(res, Chr(7), "")
    res = Replace(res, Chr(11), " ")
    res = Replace(res, Chr(160), " ")
    res = Replace(res, "\line", " ")
    res = Replace(res, ">=", ChrW(8805))
    res = Replace(res, "=>", ChrW(8805))
    res = Replace(res, "<=", ChrW(8804))
    res = Replace(res, "=<", ChrW(8804))
    Do While InStr(res, " =") > 0: res = Replace(res, " =", "="): Loop
    Do While InStr(res, "= ") > 0: res = Replace(res, "= ", "="): Loop
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    CleanParagraphString = Trim(res)
End Function

Private Function CleanHeaderCellString(ByVal s As String) As String
    Dim res As String
    res = Replace(s, Chr(13), "")
    res = Replace(res, Chr(7), "")
    res = Replace(res, Chr(160), " ")
    res = Replace(res, ">=", ChrW(8805))
    res = Replace(res, "=>", ChrW(8805))
    res = Replace(res, "<=", ChrW(8804))
    res = Replace(res, "=<", ChrW(8804))
    Do While InStr(res, " =") > 0: res = Replace(res, " =", "="): Loop
    Do While InStr(res, "= ") > 0: res = Replace(res, "= ", "="): Loop
    
    ' Strip whitespace directly bordering soft line breaks Chr(11)
    Do While InStr(res, " " & Chr(11)) > 0
        res = Replace(res, " " & Chr(11), Chr(11))
    Loop
    Do While InStr(res, Chr(11) & " ") > 0
        res = Replace(res, Chr(11) & " ", Chr(11))
    Loop
    
    ' Collapse runs of 3 or more spaces into a soft line break Chr(11) (stacks SAS labels)
    Do While InStr(res, "   ") > 0
        res = Replace(res, "   ", Chr(11))
    Loop
    
    ' Collapse double spaces to single space
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    
    CleanHeaderCellString = Trim(res)
End Function

Private Function CleanCellString(ByVal s As String) As String
    Dim res As String
    res = Replace(s, Chr(13), "")
    res = Replace(res, Chr(7), "")
    res = Replace(res, Chr(11), " ")
    res = Replace(res, Chr(160), " ")
    res = Replace(res, ">=", ChrW(8805))
    res = Replace(res, "=>", ChrW(8805))
    res = Replace(res, "<=", ChrW(8804))
    res = Replace(res, "=<", ChrW(8804))
    Do While InStr(res, " =") > 0: res = Replace(res, " =", "="): Loop
    Do While InStr(res, "= ") > 0: res = Replace(res, "= ", "="): Loop
    Do While InStr(res, ChrW(8805) & " ") > 0: res = Replace(res, ChrW(8805) & " ", ChrW(8805)): Loop
    Do While InStr(res, ChrW(8804) & " ") > 0: res = Replace(res, ChrW(8804) & " ", ChrW(8804)): Loop
    Do While InStr(res, "> ") > 0: res = Replace(res, "> ", ">"): Loop
    Do While InStr(res, "< ") > 0: res = Replace(res, "< ", "<"): Loop
    CleanCellString = Trim(res)
End Function

Private Function CollectionContains(ByVal coll As Collection, ByVal keyVal As String) As Boolean
    On Error Resume Next
    Dim v As Variant
    v = coll(keyVal)
    CollectionContains = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' ==============================================================================
' REGULATORY TYPOGRAPHY & SYMBOL NORMALIZATION HELPERS
' ==============================================================================

Private Sub ReplaceSymbolsInTable(ByVal tbl As Table)
    Dim r As Range
    
    ' 1. Greater than or equal to: >= and => -> ≥ (ChrW(8805))
    Set r = tbl.Range
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = ">="
        .Replacement.Text = ChrW(8805)
        .Forward = True
        .Wrap = 0 ' wdFindStop
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=2 ' wdReplaceAll
    End With
    
    Set r = tbl.Range
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "=>"
        .Replacement.Text = ChrW(8805)
        .Forward = True
        .Wrap = 0
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=2
    End With
    
    ' 2. Less than or equal to: <= and =< -> ≤ (ChrW(8804))
    Set r = tbl.Range
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "<="
        .Replacement.Text = ChrW(8804)
        .Forward = True
        .Wrap = 0
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=2
    End With
    
    Set r = tbl.Range
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "=<"
        .Replacement.Text = ChrW(8804)
        .Forward = True
        .Wrap = 0
        .MatchCase = False
        .MatchWholeWord = False
        .Execute Replace:=2
    End With

    ' 3. Remove space after relational operators: >=, <=, >, <, ≥, ≤
    Dim pass As Long
    For pass = 1 To 3
        ' Space after ≥
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ChrW(8805) & " "
            .Replacement.Text = ChrW(8805)
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ChrW(8805) & ChrW(160)
            .Replacement.Text = ChrW(8805)
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        
        ' Space after ≤
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ChrW(8804) & " "
            .Replacement.Text = ChrW(8804)
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ChrW(8804) & ChrW(160)
            .Replacement.Text = ChrW(8804)
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        
        ' Space after >
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "> "
            .Replacement.Text = ">"
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ">" & ChrW(160)
            .Replacement.Text = ">"
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        
        ' Space after <
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "< "
            .Replacement.Text = "<"
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "<" & ChrW(160)
            .Replacement.Text = "<"
            .Forward = True
            .Wrap = 0
            .Execute Replace:=2
        End With
    Next pass
End Sub

Private Sub RemoveEqualsSpacesInTable(ByVal tbl As Table)
    Dim r As Range
    Dim pass As Long
    
    For pass = 1 To 5
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = " ="
            .Replacement.Text = "="
            .Forward = True
            .Wrap = 0
            .MatchCase = False
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "= "
            .Replacement.Text = "="
            .Forward = True
            .Wrap = 0
            .MatchCase = False
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = ChrW(160) & "="
            .Replacement.Text = "="
            .Forward = True
            .Wrap = 0
            .MatchCase = False
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
        
        Set r = tbl.Range
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "=" & ChrW(160)
            .Replacement.Text = "="
            .Forward = True
            .Wrap = 0
            .MatchCase = False
            .MatchWholeWord = False
            .Execute Replace:=2
        End With
    Next pass
End Sub

Private Function ExtractCleanTitleWithPopulation(ByVal rawParaText As String, ByRef outPop As String) As String
    Dim s As String
    Dim lines() As String
    Dim i As Long
    Dim validLines() As String
    Dim validCount As Long
    Dim mainTitle As String
    Dim popCandidate As String
    
    outPop = ""
    s = rawParaText
    s = Replace(s, vbCrLf, Chr(11))
    s = Replace(s, vbLf, Chr(11))
    s = Replace(s, vbCr, Chr(11))
    s = Replace(s, "\line", Chr(11))
    
    lines = Split(s, Chr(11))
    validCount = 0
    ReDim validLines(0 To UBound(lines))
    
    For i = 0 To UBound(lines)
        Dim cl As String
        cl = Trim(CleanParagraphString(lines(i)))
        If Len(cl) > 0 Then
            validLines(validCount) = cl
            validCount = validCount + 1
        End If
    Next i
    
    If validCount >= 3 Then
        mainTitle = ""
        For i = 1 To validCount - 2
            If Len(mainTitle) > 0 Then
                mainTitle = mainTitle & " " & validLines(i)
            Else
                mainTitle = validLines(i)
            End If
        Next i
        popCandidate = validLines(validCount - 1)
    ElseIf validCount = 2 Then
        mainTitle = validLines(1)
        popCandidate = ""
    ElseIf validCount = 1 Then
        mainTitle = ExtractCleanTitle(validLines(0))
        popCandidate = ""
    Else
        mainTitle = "Clinical Summary Table"
        popCandidate = ""
    End If
    
    ' Clean population candidate if found
    If Len(popCandidate) > 0 Then
        If InStr(1, popCandidate, "Population:", vbTextCompare) = 1 Then
            popCandidate = Trim(Mid(popCandidate, 12))
        End If
        Do While Left(popCandidate, 1) = "(" And Right(popCandidate, 1) = ")"
            popCandidate = Trim(Mid(popCandidate, 2, Len(popCandidate) - 2))
        Loop
        If Len(popCandidate) > 0 Then
            outPop = "(" & popCandidate & ")"
        End If
    End If
    
    ' If no population from lines, check if mainTitle ends with known population pattern
    If Len(outPop) = 0 Then
        ExtractPopulationFromTitleEnd mainTitle, mainTitle, outPop
    End If
    
    If Len(outPop) > 0 Then
        If InStr(mainTitle, outPop) = 0 Then
            ExtractCleanTitleWithPopulation = mainTitle & " " & outPop
        Else
            ExtractCleanTitleWithPopulation = mainTitle
        End If
    Else
        ExtractCleanTitleWithPopulation = mainTitle
    End If
End Function

Private Sub ExtractPopulationFromTitleEnd(ByVal fullTitle As String, ByRef outTitle As String, ByRef outPop As String)
    Dim pats As Variant
    Dim pat As Variant
    Dim pLen As Long
    Dim tTrim As String
    
    outTitle = fullTitle
    outPop = ""
    tTrim = Trim(fullTitle)
    
    If Right(tTrim, 1) = ")" Then Exit Sub
    
    pats = Array( _
        "Safety Population", _
        "Safety Analysis Set", _
        "Full Analysis Set", _
        "Intent-to-Treat Population", _
        "Per-Protocol Population", _
        "Per-Protocol Set", _
        "Enrolled Participants", _
        "Enrolled Patients", _
        "Efficacy Population", _
        "PK Population", _
        "Pharmacokinetic Population", _
        "Evaluable Population" _
    )
    
    For Each pat In pats
        pLen = Len(CStr(pat))
        If Right(tTrim, pLen) = CStr(pat) Then
            outPop = "(" & CStr(pat) & ")"
            outTitle = Trim(Left(tTrim, Len(tTrim) - pLen))
            Exit Sub
        End If
    Next pat
    
    ' Generic check: ends with " Population" or " Analysis Set"
    If Right(tTrim, 11) = " Population" Then
        Dim spIdx As Long
        spIdx = InStrRev(Left(tTrim, Len(tTrim) - 11), " ")
        If spIdx > 0 Then
            outPop = "(" & Trim(Mid(tTrim, spIdx + 1)) & ")"
            outTitle = Trim(Left(tTrim, spIdx))
        Else
            outPop = "(" & tTrim & ")"
            outTitle = ""
        End If
    ElseIf Right(tTrim, 13) = " Analysis Set" Then
        spIdx = InStrRev(Left(tTrim, Len(tTrim) - 13), " ")
        If spIdx > 0 Then
            outPop = "(" & Trim(Mid(tTrim, spIdx + 1)) & ")"
            outTitle = Trim(Left(tTrim, spIdx))
        Else
            outPop = "(" & tTrim & ")"
            outTitle = ""
        End If
    End If
End Sub

Public Function FormatRowStubSentenceCase(ByVal strStub As String) As String
    If Len(Trim(strStub)) = 0 Then
        FormatRowStubSentenceCase = strStub
        Exit Function
    End If
    
    Dim s As String
    s = strStub
    
    s = Replace(s, ">=", ChrW(8805))
    s = Replace(s, "=>", ChrW(8805))
    s = Replace(s, "<=", ChrW(8804))
    s = Replace(s, "=<", ChrW(8804))
    Do While InStr(s, " =") > 0: s = Replace(s, " =", "="): Loop
    Do While InStr(s, "= ") > 0: s = Replace(s, "= ", "="): Loop
    Do While InStr(s, ChrW(8805) & " ") > 0: s = Replace(s, ChrW(8805) & " ", ChrW(8805)): Loop
    Do While InStr(s, ChrW(8805) & ChrW(160)) > 0: s = Replace(s, ChrW(8805) & ChrW(160), ChrW(8805)): Loop
    Do While InStr(s, ChrW(8804) & " ") > 0: s = Replace(s, ChrW(8804) & " ", ChrW(8804)): Loop
    Do While InStr(s, ChrW(8804) & ChrW(160)) > 0: s = Replace(s, ChrW(8804) & ChrW(160), ChrW(8804)): Loop
    Do While InStr(s, "> ") > 0: s = Replace(s, "> ", ">"): Loop
    Do While InStr(s, ">" & ChrW(160)) > 0: s = Replace(s, ">" & ChrW(160), ">"): Loop
    Do While InStr(s, "< ") > 0: s = Replace(s, "< ", "<"): Loop
    Do While InStr(s, "<" & ChrW(160)) > 0: s = Replace(s, "<" & ChrW(160), "<"): Loop
    
    Dim i As Long, ch As String
    Dim currToken As String
    Dim res As String
    Dim firstWordDone As Boolean
    
    currToken = ""
    res = ""
    firstWordDone = False
    
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If IsDelimiterChar(ch) Then
            If Len(currToken) > 0 Then
                res = res & FormatWordSentenceCase(currToken, Not firstWordDone)
                If HasAlphaChar(currToken) Then firstWordDone = True
                currToken = ""
            End If
            If ch >= "0" And ch <= "9" Then firstWordDone = True
            res = res & ch
        Else
            currToken = currToken & ch
        End If
    Next i
    
    If Len(currToken) > 0 Then
        res = res & FormatWordSentenceCase(currToken, Not firstWordDone)
    End If
    
    FormatRowStubSentenceCase = res
End Function

Private Function IsDelimiterChar(ByVal ch As String) As Boolean
    Select Case ch
        Case " ", vbTab, vbCr, vbLf, Chr(11), Chr(160), "(", ")", "[", "]", ",", ";", ":", "+", "=", "<", ">", ChrW(8805), ChrW(8804)
            IsDelimiterChar = True
        Case Else
            IsDelimiterChar = False
    End Select
End Function

Private Function HasAlphaChar(ByVal s As String) As Boolean
    Dim i As Long, c As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            HasAlphaChar = True
            Exit Function
        End If
    Next i
    HasAlphaChar = False
End Function

Private Function FormatWordSentenceCase(ByVal w As String, ByVal isFirst As Boolean) As String
    If Not HasAlphaChar(w) Then
        FormatWordSentenceCase = w
        Exit Function
    End If
    
    Dim u As String
    u = GetStandardUnit(w)
    If Len(u) > 0 Then
        FormatWordSentenceCase = u
        Exit Function
    End If
    
    If IsRomanNumeral(w) Then
        FormatWordSentenceCase = UCase(w)
        Exit Function
    End If
    
    If Len(w) = 1 Then
        If w = "n" Then
            FormatWordSentenceCase = "n"
        ElseIf w >= "A" And w <= "Z" Then
            FormatWordSentenceCase = w
        ElseIf isFirst Then
            FormatWordSentenceCase = UCase(w)
        Else
            FormatWordSentenceCase = w
        End If
        Exit Function
    End If
    
    Dim baseW As String, suffix As String
    baseW = w
    suffix = ""
    If Right(w, 2) = "'s" Or Right(w, 2) = ChrW(8217) & "s" Then
        baseW = Left(w, Len(w) - 2)
        suffix = Right(w, 2)
    End If
    
    Dim pnMatch As String
    pnMatch = GetProperNoun(baseW)
    If Len(pnMatch) > 0 Then
        FormatWordSentenceCase = pnMatch & suffix
        Exit Function
    End If
    
    Dim numLetters As Long, numUpper As Long, i As Long, ch As String
    numLetters = 0: numUpper = 0
    For i = 1 To Len(baseW)
        ch = Mid(baseW, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
            numLetters = numLetters + 1
            If ch >= "A" And ch <= "Z" Then numUpper = numUpper + 1
        End If
    Next i
    
    If numLetters >= 2 Then
        If numUpper = numLetters Then
            FormatWordSentenceCase = w
            Exit Function
        End If
        If Mid(baseW, 1, 1) >= "a" And Mid(baseW, 1, 1) <= "z" Then
            If numUpper > 0 Then
                FormatWordSentenceCase = w
                Exit Function
            End If
        End If
    End If
    
    For i = 1 To Len(baseW)
        ch = Mid(baseW, i, 1)
        If ch >= "0" And ch <= "9" Then
            FormatWordSentenceCase = w
            Exit Function
        End If
    Next i
    
    If InStr(w, "-") > 0 Then
        Dim parts() As String
        Dim newParts() As String
        Dim pIdx As Long, pStr As String, pClean As String, pSuffix As String, pPn As String
        parts = Split(w, "-")
        ReDim newParts(0 To UBound(parts))
        
        For pIdx = 0 To UBound(parts)
            pStr = parts(pIdx)
            If pIdx = 0 And isFirst Then
                newParts(pIdx) = FormatWordSentenceCase(pStr, True)
            Else
                If Len(pStr) = 1 And (pStr >= "A" And pStr <= "Z") Then
                    newParts(pIdx) = pStr
                ElseIf IsAllUpper(pStr) And Len(pStr) >= 2 Then
                    newParts(pIdx) = pStr
                ElseIf HasDigits(pStr) Then
                    newParts(pIdx) = pStr
                Else
                    pClean = pStr
                    pSuffix = ""
                    If Right(pStr, 2) = "'s" Or Right(pStr, 2) = ChrW(8217) & "s" Then
                        pClean = Left(pStr, Len(pStr) - 2)
                        pSuffix = Right(pStr, 2)
                    End If
                    pPn = GetProperNoun(pClean)
                    If Len(pPn) > 0 Then
                        newParts(pIdx) = pPn & pSuffix
                    Else
                        newParts(pIdx) = LCase(pStr)
                    End If
                End If
            End If
        Next pIdx
        FormatWordSentenceCase = Join(newParts, "-")
        Exit Function
    End If
    
    If isFirst Then
        FormatWordSentenceCase = UCase(Left(w, 1)) & LCase(Mid(w, 2))
    Else
        FormatWordSentenceCase = LCase(w)
    End If
End Function

Private Function IsAllUpper(ByVal s As String) As Boolean
    Dim i As Long, ch As String, count As Long
    count = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then
            count = count + 1
        ElseIf ch >= "a" And ch <= "z" Then
            IsAllUpper = False
            Exit Function
        End If
    Next i
    IsAllUpper = (count > 0)
End Function

Private Function HasDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            HasDigits = True
            Exit Function
        End If
    Next i
    HasDigits = False
End Function

Private Function GetStandardUnit(ByVal w As String) As String
    Select Case LCase(w)
        Case "u/l": GetStandardUnit = "U/L"
        Case "g/dl": GetStandardUnit = "g/dL"
        Case "mg/kg": GetStandardUnit = "mg/kg"
        Case "mg/dl": GetStandardUnit = "mg/dL"
        Case "ml": GetStandardUnit = "mL"
        Case "mcg": GetStandardUnit = "mcg"
        Case "iu": GetStandardUnit = "IU"
        Case "mmhg": GetStandardUnit = "mmHg"
        Case "mmol/l": GetStandardUnit = "mmol/L"
        Case "ng/ml": GetStandardUnit = "ng/mL"
        Case "pg/ml": GetStandardUnit = "pg/mL"
        Case Else: GetStandardUnit = ""
    End Select
End Function

Private Function IsRomanNumeral(ByVal w As String) As Boolean
    Select Case UCase(w)
        Case "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"
            IsRomanNumeral = True
        Case Else
            IsRomanNumeral = False
    End Select
End Function

Private Function GetProperNoun(ByVal w As String) As String
    Select Case LCase(w)
        Case "fisher": GetProperNoun = "Fisher"
        Case "pearson": GetProperNoun = "Pearson"
        Case "spearman": GetProperNoun = "Spearman"
        Case "kaplan": GetProperNoun = "Kaplan"
        Case "meier": GetProperNoun = "Meier"
        Case "wilcoxon": GetProperNoun = "Wilcoxon"
        Case "cochran": GetProperNoun = "Cochran"
        Case "mantel": GetProperNoun = "Mantel"
        Case "haenszel": GetProperNoun = "Haenszel"
        Case "hodges": GetProperNoun = "Hodges"
        Case "lehmann": GetProperNoun = "Lehmann"
        Case "bonferroni": GetProperNoun = "Bonferroni"
        Case "dunnett": GetProperNoun = "Dunnett"
        Case "tukey": GetProperNoun = "Tukey"
        Case "takeda": GetProperNoun = "Takeda"
        Case "keros": GetProperNoun = "Keros"
        Case "meddra": GetProperNoun = "MedDRA"
        Case "who": GetProperNoun = "WHO"
        Case "ctcae": GetProperNoun = "CTCAE"
        Case "ecog": GetProperNoun = "ECOG"
        Case "karnofsky": GetProperNoun = "Karnofsky"
        Case "alzheimer": GetProperNoun = "Alzheimer"
        Case "hodgkin": GetProperNoun = "Hodgkin"
        Case "parkinson": GetProperNoun = "Parkinson"
        Case "crohn": GetProperNoun = "Crohn"
        Case "raynaud": GetProperNoun = "Raynaud"
        Case "grave": GetProperNoun = "Grave"
        Case "bell": GetProperNoun = "Bell"
        Case "down": GetProperNoun = "Down"
        Case "cushing": GetProperNoun = "Cushing"
        Case "addison": GetProperNoun = "Addison"
        Case "sjogren": GetProperNoun = "Sjogren"
        Case "ebv": GetProperNoun = "EBV"
        Case "cmv": GetProperNoun = "CMV"
        Case "jan": GetProperNoun = "Jan"
        Case "feb": GetProperNoun = "Feb"
        Case "mar": GetProperNoun = "Mar"
        Case "apr": GetProperNoun = "Apr"
        Case "may": GetProperNoun = "May"
        Case "jun": GetProperNoun = "Jun"
        Case "jul": GetProperNoun = "Jul"
        Case "aug": GetProperNoun = "Aug"
        Case "sep": GetProperNoun = "Sep"
        Case "oct": GetProperNoun = "Oct"
        Case "nov": GetProperNoun = "Nov"
        Case "dec": GetProperNoun = "Dec"
        Case Else: GetProperNoun = ""
    End Select
End Function

