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

' Session-persistent last used RTF folder for writer convenience
Private mLastRTFFolder As String

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
    Dim undoRec As Object
    Dim undoStarted As Boolean
    
    Dim rawTableTitle As String
    Dim defaultCleanTitle As String
    Dim defaultSourceNumber As String
    Dim defaultPopulation As String
    
    Dim p As Paragraph
    Dim pText As String
    Dim rawPText As String
    Dim bylinesColl As Collection
    
    Dim parsedTablesColl As Collection
    Dim activeTable As Object
    Dim hfTitleBlock As String
    Dim hfFootnotes As Collection
    Dim tIdx As Long
    Dim totalSrcTables As Long
    Dim srcTbl As Table
    
    undoStarted = False
    On Error GoTo ErrorHandler
    
    If Documents.Count = 0 Then
        MsgBox "Please open or focus a CSR document first.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    Set targetDoc = ActiveDocument
    
    ' 1. Select RTF Output file (remembers last folder used in this Word session)
    Set fd = Application.FileDialog(1) ' 1 = msoFileDialogFilePicker
    With fd
        .Title = "Select SAS/R Clinical Table or Listing (RTF)"
        .Filters.Clear
        .Filters.Add "Rich Text Format (*.rtf)", "*.rtf"
        .Filters.Add "Word Documents (*.doc; *.docx)", "*.doc; *.docx"
        .Filters.Add "All Files (*.*)", "*.*"
        .AllowMultiSelect = False
        If Len(mLastRTFFolder) > 0 Then
            .InitialFileName = mLastRTFFolder
        End If
        If .Show <> -1 Then Exit Sub
        rtfPath = .SelectedItems(1)
    End With
    mLastRTFFolder = ExtractFolderPath(rtfPath)
    
    ' 1b. Pre-Flight Check before Word opens the RTF:
    '     Reads the raw RTF header in <0.03s to detect massive files (e.g. lab shift / by-visit tables)
    '     so the writer can see the table title, page count, and estimated scan time and cancel immediately.
    If Not CheckRTFPreFlightBeforeOpen(rtfPath) Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = 0 ' wdAlertsNone
    Application.StatusBar = "Takeda CSR Tools: Opening " & ExtractFileName(rtfPath) & "..."
    
    ' 2. Open RTF in hidden staging document
    Set stagingDoc = Documents.Open(FileName:=rtfPath, ConfirmConversions:=False, ReadOnly:=True, Visible:=False)
    
    ' 3. Parse Metadata (Title, Source Number, Bylines)
    Set bylinesColl = New Collection
    rawTableTitle = ""
    defaultSourceNumber = ""
    defaultPopulation = ""
    
    For Each p In stagingDoc.Paragraphs
        rawPText = p.Range.Text
        pText = Trim(CleanParagraphString(rawPText))
        If Len(pText) > 0 Then
            If (InStr(1, pText, "Table ", vbTextCompare) = 1 Or InStr(1, pText, "Listing ", vbTextCompare) = 1) And Len(rawTableTitle) = 0 Then
                rawTableTitle = pText
                defaultSourceNumber = ExtractSourceNumber(pText)
                defaultCleanTitle = ExtractCleanTitleWithPopulation(rawPText, defaultPopulation)
            ElseIf InStr(1, pText, "Treatment Group:", vbTextCompare) > 0 Or InStr(1, pText, "Population:", vbTextCompare) > 0 Then
                If Not CollectionContains(bylinesColl, pText) Then
                    bylinesColl.Add pText
                End If
            End If
        End If
    Next p
    
    ' 3b. Many SAS ODS RTF outputs place titles/footnotes in the page header/footer
    '     (BODYTITLE off) rather than in the document body. Scan those stories too.
    Set hfFootnotes = New Collection
    hfTitleBlock = ""
    ScanHeaderFooterMetadata stagingDoc, hfTitleBlock, hfFootnotes
    If Len(rawTableTitle) = 0 And Len(hfTitleBlock) > 0 Then
        rawTableTitle = CleanParagraphString(hfTitleBlock)
        defaultSourceNumber = ExtractSourceNumber(hfTitleBlock)
        defaultCleanTitle = ExtractCleanTitleWithPopulation(hfTitleBlock, defaultPopulation)
    End If
    
    If Len(defaultCleanTitle) = 0 Then
        If Len(rawTableTitle) > 0 Then
            defaultCleanTitle = rawTableTitle
        Else
            defaultCleanTitle = "Clinical Summary Table"
        End If
    End If
    
    ' Append byline to title if present (enclosing population in parentheses)
    If bylinesColl.Count > 0 Then
        Dim bLine As Variant
        For Each bLine In bylinesColl
            Dim bStr As String
            bStr = CStr(bLine)
            If InStr(1, bStr, "Population:", vbTextCompare) > 0 Then
                If Len(defaultPopulation) = 0 Then
                    Dim byPop As String
                    byPop = Trim(Mid(bStr, InStr(1, bStr, "Population:", vbTextCompare) + 11))
                    Do While Left(byPop, 1) = "(" And Right(byPop, 1) = ")"
                        byPop = Trim(Mid(byPop, 2, Len(byPop) - 2))
                    Loop
                    If Len(byPop) > 0 Then
                        defaultPopulation = "(" & byPop & ")"
                    End If
                End If
            Else
                defaultCleanTitle = defaultCleanTitle & " (" & bStr & ")"
            End If
        Next bLine
    End If
    
    ' 4. Parse Tables across pages (detecting distinct tables vs multi-page continuation)
    Set parsedTablesColl = New Collection
    Set activeTable = Nothing
    
    Dim secSubgroups() As String
    PreScanSectionSubgroups stagingDoc, secSubgroups
    
    totalSrcTables = stagingDoc.Tables.Count
    For tIdx = 1 To totalSrcTables
        If tIdx = 1 Or (tIdx Mod 5 = 0) Or tIdx = totalSrcTables Then
            Application.StatusBar = "Takeda CSR Tools: Scanning RTF page " & tIdx & " of " & totalSrcTables & " (" & CLng((tIdx * 100) / totalSrcTables) & "%)..."
            If tIdx Mod 10 = 0 Then DoEvents
        End If
        Set srcTbl = stagingDoc.Tables(tIdx)
        
        If Not IsBannerTable(srcTbl) Then
            Dim tableTitleInRow1 As Boolean
            Dim thisSubtitle As String
            Dim thisSourceNumber As String
            Dim thisBaseTitle As String
            Dim thisPop As String
            Dim headerStartRow As Long
            Dim thisNumHeaderRows As Long
            Dim thisColCount As Long
            Dim r As Long, c As Long
            
            tableTitleInRow1 = False
            thisSubtitle = ""
            thisSourceNumber = defaultSourceNumber
            thisBaseTitle = defaultCleanTitle
            thisPop = defaultPopulation
            
            ' Check preceding paragraph for table title
            Dim precTitle As String
            precTitle = GetPrecedingTableTitle(srcTbl)
            If Len(precTitle) > 0 Then
                thisSourceNumber = ExtractSourceNumber(precTitle)
                thisBaseTitle = ExtractCleanTitleWithPopulation(precTitle, thisPop)
            End If
            
            ' Check if Row 1 has table title / subtitle
            If srcTbl.Rows.Count >= 2 Then
                Dim firstCellText As String
                firstCellText = CleanParagraphString(srcTbl.Rows(1).Cells(1).Range.Text)
                If InStr(1, firstCellText, "Table ", vbTextCompare) = 1 Or _
                   InStr(1, firstCellText, "Listing ", vbTextCompare) = 1 Then
                    If srcTbl.Rows(2).HeadingFormat Then
                        tableTitleInRow1 = True
                        thisSourceNumber = ExtractSourceNumber(firstCellText)
                        thisBaseTitle = ExtractCleanTitleWithPopulation(firstCellText, thisPop)
                        If srcTbl.Rows(1).Cells.Count >= 2 Then
                            thisSubtitle = CleanParagraphString(srcTbl.Rows(1).Cells(2).Range.Text)
                        End If
                    End If
                ElseIf srcTbl.Rows(1).Cells.Count = 2 And Len(firstCellText) = 0 Then
                    ' Cell 1 is blank, Cell 2 might be subtitle (spanner BY-group)
                    Dim cell2Text As String
                    cell2Text = CleanParagraphString(srcTbl.Rows(1).Cells(2).Range.Text)
                    If Len(cell2Text) > 0 And srcTbl.Rows(2).HeadingFormat Then
                        tableTitleInRow1 = True
                        thisSubtitle = cell2Text
                    End If
                End If
            End If
            
            If tableTitleInRow1 Then
                headerStartRow = 2
            Else
                headerStartRow = 1
            End If
            
            ' Check section header for BY-group subtitle (e.g. Treatment: Placebo (N=258))
            If Len(thisSubtitle) = 0 Then
                On Error Resume Next
                Dim tblSecIdx As Long: tblSecIdx = srcTbl.Range.Sections(1).Index
                If tblSecIdx >= LBound(secSubgroups) And tblSecIdx <= UBound(secSubgroups) Then
                    thisSubtitle = secSubgroups(tblSecIdx)
                End If
                On Error GoTo ErrorHandler
            End If
            
            ' Count header rows
            thisNumHeaderRows = 0
            For r = headerStartRow To srcTbl.Rows.Count
                If srcTbl.Rows(r).HeadingFormat Then
                    thisNumHeaderRows = thisNumHeaderRows + 1
                Else
                    Exit For
                End If
            Next r
            If thisNumHeaderRows = 0 Then thisNumHeaderRows = 1
            
            ' Determine column count from the last header row or table columns
            Dim lastHdrRowIdx As Long
            lastHdrRowIdx = headerStartRow + thisNumHeaderRows - 1
            thisColCount = srcTbl.Rows(lastHdrRowIdx).Cells.Count
            If thisColCount < srcTbl.Columns.Count Then
                thisColCount = srcTbl.Columns.Count
            End If
            
            ' Check continuation vs new distinct table
            Dim isContinuation As Boolean: isContinuation = False
            If Not activeTable Is Nothing Then
                If activeTable("ColCount") = thisColCount And activeTable("Subtitle") = thisSubtitle Then
                    isContinuation = True
                End If
            End If
            
            If Not isContinuation Then
                If Not activeTable Is Nothing Then
                    parsedTablesColl.Add activeTable
                    Set activeTable = Nothing
                End If
                
                ' Initialize new activeTable dictionary
                Set activeTable = CreateObject("Scripting.Dictionary")
                activeTable("TableIndex") = parsedTablesColl.Count + 1
                activeTable("SourceNumber") = thisSourceNumber
                activeTable("Title") = thisBaseTitle
                activeTable("Subtitle") = thisSubtitle
                activeTable("Population") = thisPop
                
                Dim fullT As String: fullT = Trim(thisBaseTitle)
                If Len(thisPop) > 0 Then
                    If Right(fullT, Len(thisPop)) = thisPop Then
                        fullT = Trim(Left(fullT, Len(fullT) - Len(thisPop)))
                    End If
                End If
                If Len(thisSubtitle) > 0 Then fullT = fullT & " - " & thisSubtitle
                If Len(thisPop) > 0 Then fullT = fullT & " " & thisPop
                activeTable("FullTitle") = fullT
                
                activeTable("ColCount") = thisColCount
                activeTable("NumHeaderRows") = thisNumHeaderRows
                Set activeTable("BodyRows") = New Collection
                Set activeTable("Footnotes") = New Collection
                
                Dim outWidthPts As Single
                CalculateTableDimensions srcTbl, thisColCount, outWidthPts
                activeTable("TotalWidthPts") = outWidthPts
                
                ' Extract column widths from last header row
                Dim colWidths() As Single
                ReDim colWidths(1 To thisColCount)
                Dim cWTotal As Single: cWTotal = 0#
                For c = 1 To thisColCount
                    If c <= srcTbl.Rows(lastHdrRowIdx).Cells.Count Then
                        colWidths(c) = srcTbl.Rows(lastHdrRowIdx).Cells(c).Width
                    Else
                        colWidths(c) = 50#
                    End If
                    cWTotal = cWTotal + colWidths(c)
                Next c
                
                ' Extract header values (mapping merged spanners)
                Dim hVals() As String
                ReDim hVals(1 To thisNumHeaderRows, 1 To thisColCount)
                For r = 1 To thisNumHeaderRows
                    Dim actualRIdx As Long: actualRIdx = headerStartRow + r - 1
                    Dim rCellsCount As Long: rCellsCount = srcTbl.Rows(actualRIdx).Cells.Count
                    
                    If rCellsCount = thisColCount Then
                        For c = 1 To thisColCount
                            Dim cTxt As String: cTxt = ""
                            On Error Resume Next
                            cTxt = srcTbl.Rows(actualRIdx).Cells(c).Range.Text
                            On Error GoTo ErrorHandler
                            hVals(r, c) = CleanHeaderCellString(cTxt)
                        Next c
                    Else
                        ' Row has merged cells (spanners)
                        Dim cumW As Single: cumW = 0#
                        Dim cMap As Long: cMap = 1
                        Dim kCell As Long
                        For kCell = 1 To rCellsCount
                            Dim cellW As Single
                            cellW = srcTbl.Rows(actualRIdx).Cells(kCell).Width
                            Dim sTxt As String
                            sTxt = CleanHeaderCellString(srcTbl.Rows(actualRIdx).Cells(kCell).Range.Text)
                            Dim spanEnd As Long
                            spanEnd = FindColumnSpanEnd(colWidths, thisColCount, cumW + cellW)
                            If spanEnd < cMap Then spanEnd = cMap
                            Dim fillC As Long
                            For fillC = cMap To spanEnd
                                If fillC <= thisColCount Then hVals(r, fillC) = sTxt
                            Next fillC
                            cumW = cumW + cellW
                            cMap = spanEnd + 1
                        Next kCell
                    End If
                Next r
                activeTable("HeaderValues") = hVals
            End If
            
            ' Extract body rows and footnote rows from srcTbl
            Dim bodyStartRow As Long
            bodyStartRow = headerStartRow + thisNumHeaderRows
            Dim rIdx As Long
            For rIdx = bodyStartRow To srcTbl.Rows.Count
                Dim isFootnoteRow As Boolean: isFootnoteRow = False
                Dim isSpacerRow As Boolean: isSpacerRow = False
                Dim rCells As Long: rCells = srcTbl.Rows(rIdx).Cells.Count
                
                If rCells = 1 Then
                    isSpacerRow = True
                    Dim singleText As String
                    singleText = srcTbl.Rows(rIdx).Cells(1).Range.Text
                    If IsFootnoteParagraph(singleText) Then
                        AddFootnotesFromText activeTable("Footnotes"), singleText
                    End If
                Else
                    Dim firstText As String
                    firstText = srcTbl.Rows(rIdx).Cells(1).Range.Text
                    If IsFootnoteParagraph(firstText) Then
                        isFootnoteRow = True
                        AddFootnotesFromText activeTable("Footnotes"), firstText
                    Else
                        Dim allBlank As Boolean: allBlank = True
                        For c = 1 To rCells
                            Dim clText As String
                            clText = CleanCellString(srcTbl.Rows(rIdx).Cells(c).Range.Text)
                            If Len(clText) > 0 And clText <> "~" Then
                                allBlank = False
                                Exit For
                            End If
                        Next c
                        If allBlank Then isSpacerRow = True
                    End If
                End If
                
                If Not isSpacerRow And Not isFootnoteRow Then
                    Dim rowData() As String
                    Dim rowIndents() As Single
                    ReDim rowData(1 To thisColCount)
                    ReDim rowIndents(1 To thisColCount)
                    
                    For c = 1 To thisColCount
                        Dim cellText As String: cellText = ""
                        Dim rawIndentPts As Single: rawIndentPts = 0#
                        
                        If c <= rCells Then
                            On Error Resume Next
                            cellText = srcTbl.Rows(rIdx).Cells(c).Range.Text
                            rawIndentPts = srcTbl.Rows(rIdx).Cells(c).Range.ParagraphFormat.LeftIndent
                            On Error GoTo ErrorHandler
                        End If
                        
                        If c = 1 Then
                            Dim cIndent As Single
                            cIndent = DetermineIndentPoints(rawIndentPts, cellText)
                            rowData(1) = FormatRowStubSentenceCase(CleanCellString(cellText))
                            rowIndents(1) = cIndent
                        Else
                            rowData(c) = CleanCellString(cellText)
                            rowIndents(c) = 0#
                        End If
                    Next c
                    
                    If Not IsContinuationStubRow(rowData, thisColCount) Then
                        activeTable("BodyRows").Add Array(rowData, rowIndents)
                    End If
                End If
            Next rIdx
        End If
    Next tIdx
    
    If Not activeTable Is Nothing Then
        parsedTablesColl.Add activeTable
        Set activeTable = Nothing
    End If
    
    ' Close staging doc
    stagingDoc.Close SaveChanges:=0
    Set stagingDoc = Nothing
    Application.StatusBar = False
    
    If parsedTablesColl.Count = 0 Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = -1
        MsgBox "No valid clinical data table rows could be parsed from the selected file.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    
    ' Normalize per-table indentation tiers and record original row counts before any threshold filtering
    Dim ptItem As Variant
    For Each ptItem In parsedTablesColl
        NormalizeTableIndentLevels ptItem
        ptItem("OriginalRowCount") = ptItem("BodyRows").Count
    Next ptItem
    
    ' 5. Prompt user if multiple distinct tables found (supports multi-select e.g. '5-8', 'Race', '2, 3, 4' or 'All')
    Dim tablesToInsert As Collection
    Set tablesToInsert = New Collection
    
    If parsedTablesColl.Count = 1 Then
        tablesToInsert.Add parsedTablesColl(1)
    Else
        Set tablesToInsert = PromptMultiTableSelection(parsedTablesColl)
        If tablesToInsert.Count = 0 Then
            Application.ScreenUpdating = True
            Application.DisplayAlerts = -1
            Exit Sub
        End If
    End If
    
    ' 5b. Check if any selected table is a long incidence/grade table (> 40 rows) and offer threshold filtering
    PromptAndFilterSelectedTables tablesToInsert
    
    ' 5c. Pre-Insertion Confirmation Step:
    '     Displays the exact table title(s) to be inserted, data row count, estimated CSR page length,
    '     orientation, and estimated insertion time, allowing the writer to Cancel before modifying the CSR.
    If Not ConfirmTableInsertionWithUser(rtfPath, tablesToInsert, defaultSourceNumber) Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = -1
        Application.StatusBar = False
        Exit Sub
    End If
    
    ' 6. Start a single custom Undo record so the writer can undo the entire insertion with one Ctrl+Z
    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "Insert Takeda CSR Table"
        undoStarted = True
    End If
    On Error GoTo ErrorHandler
    
    ' 7. Insert selected CSR tables into target document
    Set targetRange = Selection.Range
    Dim insTbl As Variant
    Dim lastTbl As Table
    Dim insertCount As Long: insertCount = 0
    Dim totalToInsert As Long: totalToInsert = tablesToInsert.Count
    
    For Each insTbl In tablesToInsert
        If insertCount > 0 Then
            ' Insert paragraph break between consecutive tables
            targetRange.Collapse 0 ' wdCollapseEnd
            targetRange.InsertParagraphAfter
            targetRange.Collapse 0 ' wdCollapseEnd
        End If
        
        Application.StatusBar = "Takeda CSR Tools: Inserting table " & (insertCount + 1) & " of " & totalToInsert & " (" & insTbl("BodyRows").Count & " rows)..."
        DoEvents
        
        Set lastTbl = InsertSingleCSRTable(targetDoc, targetRange, insTbl, defaultSourceNumber, hfFootnotes)
        insertCount = insertCount + 1
        
        Set targetRange = lastTbl.Range
        targetRange.Collapse 0 ' wdCollapseEnd
    Next insTbl
    
    If undoStarted And Not undoRec Is Nothing Then
        On Error Resume Next
        undoRec.EndCustomRecord
        undoStarted = False
        On Error GoTo ErrorHandler
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = -1
    
    If insertCount = 1 Then
        Dim sTbl As Object: Set sTbl = tablesToInsert(1)
        MsgBox "Table imported and Takeda CSR formatting successfully applied!" & vbCrLf & _
               "Title: " & sTbl("FullTitle") & vbCrLf & _
               "Columns: " & sTbl("ColCount") & " | Data Rows: " & sTbl("BodyRows").Count & vbCrLf & _
               "(Tip: Press Ctrl+Z once to undo the entire table insertion if needed.)", vbInformation, "Takeda CSR Tools"
    Else
        MsgBox "Successfully imported and formatted " & insertCount & " CSR tables!" & vbCrLf & _
               "(Tip: Press Ctrl+Z once to undo the entire table insertion if needed.)", vbInformation, "Takeda CSR Tools"
    End If
    Exit Sub
    
ErrorHandler:
    If Not stagingDoc Is Nothing Then
        On Error Resume Next
        stagingDoc.Close SaveChanges:=0
    End If
    If undoStarted And Not undoRec Is Nothing Then
        On Error Resume Next
        undoRec.EndCustomRecord
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = -1
    MsgBox "An error occurred during RTF import: " & Err.Description, vbCritical, "Takeda CSR Tools Error"
End Sub

Private Function InsertSingleCSRTable( _
    ByVal targetDoc As Document, _
    ByRef targetRange As Range, _
    ByVal tblData As Object, _
    ByVal defaultSourceNumber As String, _
    ByVal extraFootnotes As Collection _
) As Table

    Dim requiresLandscape As Boolean
    Dim colCount As Long: colCount = tblData("ColCount")
    Dim totalWidthPts As Single: totalWidthPts = CSng(tblData("TotalWidthPts"))
    Dim numHeaderRows As Long: numHeaderRows = tblData("NumHeaderRows")
    Dim bodyRowsColl As Collection: Set bodyRowsColl = tblData("BodyRows")
    Dim distinctFootnotes As Collection: Set distinctFootnotes = tblData("Footnotes")
    
    If colCount >= CSR_AUTO_LANDSCAPE_COL_THRESHOLD Or totalWidthPts > CSR_PORTRAIT_MAX_WIDTH_PTS Then
        requiresLandscape = True
    Else
        requiresLandscape = False
    End If
    
    If requiresLandscape Then
        InsertLandscapeSection targetRange
    End If
    
    Dim totalTargetRows As Long
    totalTargetRows = 1 + numHeaderRows + bodyRowsColl.Count + 1 ' Title + Headers + Body + Footnotes
    
    Dim finalTable As Table
    Set finalTable = targetDoc.Tables.Add(Range:=targetRange, NumRows:=totalTargetRows, NumColumns:=colCount)
    
    ' Populate Row 1: Title with Proper Cross-Referenceable Field Codes
    If colCount > 1 Then
        finalTable.Rows(1).Cells.Merge
    End If
    Dim srcNum As String: srcNum = tblData("SourceNumber")
    If Len(srcNum) = 0 Then srcNum = defaultSourceNumber
    InsertCrossReferencedTitle finalTable.Rows(1).Cells(1).Range, tblData("FullTitle")
    
    ' Populate Header Rows
    Dim hVals As Variant: hVals = tblData("HeaderValues")
    Dim rIdx As Long, cIdx As Long
    For rIdx = 1 To numHeaderRows
        For cIdx = 1 To colCount
            finalTable.Rows(1 + rIdx).Cells(cIdx).Range.Text = hVals(rIdx, cIdx)
        Next cIdx
        ApplyHeaderRowStyle finalTable.Rows(1 + rIdx)
        MergeAdjacentIdenticalHeaderCells finalTable.Rows(1 + rIdx), colCount
    Next rIdx
    
    ' Populate Body Rows
    Dim targetRowOffset As Long: targetRowOffset = 1 + numHeaderRows
    Dim bodyRowIdx As Long: bodyRowIdx = 0
    Dim totalBodyRows As Long: totalBodyRows = bodyRowsColl.Count
    Dim rItem As Variant, rVals As Variant
    For Each rItem In bodyRowsColl
        bodyRowIdx = bodyRowIdx + 1
        targetRowOffset = targetRowOffset + 1
        If bodyRowIdx Mod 25 = 0 Then
            Application.StatusBar = "Takeda CSR Tools: Populating row " & bodyRowIdx & " of " & totalBodyRows & "..."
            DoEvents
        End If
        rVals = rItem(0)
        For cIdx = 1 To colCount
            finalTable.Rows(targetRowOffset).Cells(cIdx).Range.Text = rVals(cIdx)
        Next cIdx
    Next rItem
    
    ' Populate Footnotes Row (Merged across all columns)
    Dim footnoteRowIdx As Long: footnoteRowIdx = totalTargetRows
    If colCount > 1 Then
        finalTable.Rows(footnoteRowIdx).Cells.Merge
    End If
    
    Dim combinedFootnotes As String
    combinedFootnotes = "Source: " & srcNum & "." & vbCr
    Dim fnItem As Variant
    For Each fnItem In distinctFootnotes
        combinedFootnotes = combinedFootnotes & CStr(fnItem) & vbCr
    Next fnItem
    If Not extraFootnotes Is Nothing Then
        For Each fnItem In extraFootnotes
            If Not CollectionContains(distinctFootnotes, CStr(fnItem)) Then
                combinedFootnotes = combinedFootnotes & CStr(fnItem) & vbCr
            End If
        Next fnItem
    End If
    Do While Right(combinedFootnotes, 1) = vbCr Or Right(combinedFootnotes, 1) = " "
        combinedFootnotes = Left(combinedFootnotes, Len(combinedFootnotes) - 1)
    Loop
    finalTable.Rows(footnoteRowIdx).Cells(1).Range.Text = combinedFootnotes
    ApplyFootnoteStyle finalTable.Rows(footnoteRowIdx).Cells(1).Range
    
    ' Apply Full Takeda CSR Table Rules (Borders, Padding, AutoFit, Base Row Styles)
    ' Pass skipStubReformat:=True since rowData(1) was already formatted in sentence case during RTF extraction
    FormatCSRTableStructure finalTable, numHeaderRows:=numHeaderRows, skipStubReformat:=True
    
    ' Re-apply Hierarchy Indentations & Styles to Column 1
    ' All indented rows receive "Table:Indent"; deeper levels (Level 2+) keep "Table:Indent" plus additional LeftIndent
    targetRowOffset = 1 + numHeaderRows
    Dim rInds As Variant, indVal As Single
    For Each rItem In bodyRowsColl
        targetRowOffset = targetRowOffset + 1
        rInds = rItem(1)
        indVal = CSng(rInds(1))
        If indVal > 0 Then
            On Error Resume Next
            finalTable.Rows(targetRowOffset).Cells(1).Range.Style = "Table:Indent"
            If Err.Number <> 0 Then
                Err.Clear
                finalTable.Rows(targetRowOffset).Cells(1).Range.Style = "Table:Text"
                If Err.Number <> 0 Then Err.Clear
            End If
            On Error GoTo 0
            finalTable.Rows(targetRowOffset).Cells(1).Range.ParagraphFormat.LeftIndent = indVal
        End If
    Next rItem
    
    ' Close landscape section if opened
    If requiresLandscape Then
        Dim afterRange As Range
        Set afterRange = finalTable.Range
        afterRange.Collapse 0 ' wdCollapseEnd
        CloseLandscapeSection afterRange
    End If
    
    Set InsertSingleCSRTable = finalTable
End Function

Private Sub MergeAdjacentIdenticalHeaderCells(ByVal rowObj As Row, ByVal colCount As Long)
    Dim cStart As Long
    Dim cEnd As Long
    Dim matchText As String
    Dim currText As String
    
    cStart = 1
    Do While cStart < rowObj.Cells.Count
        matchText = CleanCellString(rowObj.Cells(cStart).Range.Text)
        If Len(matchText) > 0 Then
            cEnd = cStart + 1
            Do While cEnd <= rowObj.Cells.Count
                currText = CleanCellString(rowObj.Cells(cEnd).Range.Text)
                If currText = matchText Then
                    cEnd = cEnd + 1
                Else
                    Exit Do
                End If
            Loop
            
            If cEnd - 1 > cStart Then
                On Error Resume Next
                rowObj.Cells(cStart).Merge MergeTo:=rowObj.Cells(cEnd - 1)
                rowObj.Cells(cStart).Range.Text = matchText
                rowObj.Cells(cStart).Range.ParagraphFormat.Alignment = 1 ' wdAlignParagraphCenter
                On Error GoTo 0
            End If
        End If
        cStart = cStart + 1
    Loop
End Sub

Private Function FindColumnSpanEnd(colWidths() As Single, ByVal colCount As Long, ByVal targetRightPts As Single) As Long
    Dim c As Long
    Dim runningW As Single: runningW = 0#
    For c = 1 To colCount
        runningW = runningW + colWidths(c)
        If runningW >= (targetRightPts - 2#) Then
            FindColumnSpanEnd = c
            Exit Function
        End If
    Next c
    FindColumnSpanEnd = colCount
End Function

Private Function GetPrecedingTableTitle(ByVal tbl As Table) As String
    Dim p As Paragraph
    Dim rng As Range
    On Error Resume Next
    Set rng = tbl.Range
    rng.Collapse 1 ' wdCollapseStart
    
    Dim i As Long
    For i = 1 To 5
        rng.MoveStart 4, -1 ' 4 = wdParagraph
        Set p = rng.Paragraphs(1)
        Dim t As String
        t = Trim(CleanParagraphString(p.Range.Text))
        If InStr(1, t, "Table ", vbTextCompare) = 1 Or InStr(1, t, "Listing ", vbTextCompare) = 1 Then
            GetPrecedingTableTitle = p.Range.Text
            Exit Function
        End If
    Next i
    On Error GoTo 0
    GetPrecedingTableTitle = ""
End Function

Public Sub FormatCurrentSelectionAsCSRTable()
    Dim tbl As Table
    Dim undoRec As Object
    Dim undoStarted As Boolean
    
    If Selection.Tables.Count = 0 Then
        MsgBox "Please place your cursor inside a table first.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    
    undoStarted = False
    On Error Resume Next
    Set undoRec = Application.UndoRecord
    If Not undoRec Is Nothing Then
        undoRec.StartCustomRecord "Format Takeda CSR Table"
        undoStarted = True
    End If
    On Error GoTo 0
    
    Set tbl = Selection.Tables(1)
    FormatCSRTableStructure tbl, numHeaderRows:=1, skipStubReformat:=False
    
    If undoStarted And Not undoRec Is Nothing Then
        On Error Resume Next
        undoRec.EndCustomRecord
        On Error GoTo 0
    End If
    
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
    ' Priority 2: RTF \li Twip-based left indentations (e.g. \li200=10pt, \li300=15pt, \li600=30pt)
    ElseIf srcIndentPts > 0 Then
        If srcIndentPts >= 32# Then        ' \li720+ = 36+ pt
            DetermineIndentPoints = CSR_INDENT_LEVEL_3_PTS
        ElseIf srcIndentPts >= 16.5 Then   ' \li360 / \li400 / \li600 = 18-30 pt
            DetermineIndentPoints = CSR_INDENT_LEVEL_2_PTS
        ElseIf srcIndentPts >= 4# Then     ' \li180 / \li200 / \li300 = 9-15 pt
            DetermineIndentPoints = CSR_INDENT_LEVEL_1_PTS
        Else
            DetermineIndentPoints = 0#
        End If
    Else
        DetermineIndentPoints = 0#
    End If
End Function

Private Sub NormalizeTableIndentLevels(ByVal tblDict As Object)
    ' Ensures that whatever indentation step size SAS used in a given table (e.g., 2 vs 3 vs 4 spaces,
    ' or \li200 vs \li300 vs \li360), the first positive indentation tier in the table is always mapped
    ' to Level 1 (12 pt), the second tier to Level 2 (24 pt), the third to Level 3 (36 pt), etc.
    Dim bodyRows As Collection
    Dim tiers() As Single
    Dim tierCount As Long
    Dim i As Long, j As Long
    Dim rItem As Variant
    Dim rData As Variant
    Dim rInds As Variant
    Dim val As Single
    Dim found As Boolean
    Dim tmp As Single
    
    On Error Resume Next
    Set bodyRows = tblDict("BodyRows")
    If bodyRows Is Nothing Then
        On Error GoTo 0
        Exit Sub
    End If
    If bodyRows.Count = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    tierCount = 0
    ReDim tiers(1 To 8)
    
    For i = 1 To bodyRows.Count
        rItem = bodyRows(i)
        rInds = rItem(1)
        val = CSng(rInds(1))
        If val > 0.5 Then
            found = False
            For j = 1 To tierCount
                If Abs(val - tiers(j)) < 1.5 Then
                    found = True
                    Exit For
                End If
            Next j
            If Not found And tierCount < 8 Then
                tierCount = tierCount + 1
                tiers(tierCount) = val
            End If
        End If
    Next i
    
    If tierCount = 0 Then Exit Sub
    
    ' Sort distinct positive tiers ascending
    For i = 1 To tierCount - 1
        For j = i + 1 To tierCount
            If tiers(j) < tiers(i) Then
                tmp = tiers(i)
                tiers(i) = tiers(j)
                tiers(j) = tmp
            End If
        Next j
    Next i
    
    Dim normRows As Collection
    Set normRows = New Collection
    
    For i = 1 To bodyRows.Count
        rItem = bodyRows(i)
        rData = rItem(0)
        rInds = rItem(1)
        val = CSng(rInds(1))
        If val <= 0.5 Then
            rInds(1) = 0#
        Else
            Dim matchedTier As Long: matchedTier = 1
            For j = 1 To tierCount
                If Abs(val - tiers(j)) < 1.5 Then
                    matchedTier = j
                    Exit For
                End If
            Next j
            If matchedTier = 1 Then
                rInds(1) = CSR_INDENT_LEVEL_1_PTS
            ElseIf matchedTier = 2 Then
                rInds(1) = CSR_INDENT_LEVEL_2_PTS
            ElseIf matchedTier = 3 Then
                rInds(1) = CSR_INDENT_LEVEL_3_PTS
            Else
                rInds(1) = CSR_INDENT_LEVEL_4_PTS
            End If
        End If
        normRows.Add Array(rData, rInds)
    Next i
    
    tblDict.Remove "BodyRows"
    Set tblDict("BodyRows") = normRows
End Sub

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
    On Error Resume Next
    For Each hf In sec.Headers
        hf.LinkToPrevious = False
    Next hf
    For Each hf In sec.Footers
        hf.LinkToPrevious = False
    Next hf
    Err.Clear
    On Error GoTo 0
End Sub

' ==============================================================================
' TABLE FORMATTING & TAKEDA 3-LINE BORDER RULES
' ==============================================================================

Public Sub FormatCSRTableStructure(ByVal tbl As Table, ByVal numHeaderRows As Long, Optional ByVal skipStubReformat As Boolean = False)
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
        
        If Not skipStubReformat Then
            ' Format Row Stub Column 1 in Sentence Case if applicable (e.g. for pre-existing tables)
            Dim stubCell As Cell
            Set stubCell = tbl.Rows(rIdx).Cells(1)
            Dim stubText As String
            stubText = stubCell.Range.Text
            If Len(stubText) >= 2 Then
                ' Word cell text ends with Chr(13) & Chr(7)
                stubText = Left(stubText, Len(stubText) - 2)
                Dim formattedStub As String
                formattedStub = FormatRowStubSentenceCase(stubText)
                If formattedStub <> stubText Then
                    stubCell.Range.Text = formattedStub
                End If
            End If
            
            ' Check if Column 1 has existing indent to apply Table:Indent style
            On Error Resume Next
            Dim existingIndent As Single
            existingIndent = stubCell.Range.ParagraphFormat.LeftIndent
            If existingIndent > 0.5 Then
                stubCell.Range.Style = "Table:Indent"
                If Err.Number <> 0 Then Err.Clear
                stubCell.Range.ParagraphFormat.LeftIndent = existingIndent
            End If
            On Error GoTo 0
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
    Dim nonFnRange As Range
    
    ' 1. In all non-footnote rows (Title, Header, and Body rows), ensure there is always a single space
    '    before any bracketed footnote callout (e.g. "Week 28[b]" or "Week 28 [b]" -> "Week 28 [b]"),
    '    so the resulting superscript always has a space before the letter ("Week 28 ᵇ").
    If tbl.Rows.Count >= 2 Then
        On Error Resume Next
        Set nonFnRange = tbl.Range.Document.Range(tbl.Rows(1).Range.Start, tbl.Rows(tbl.Rows.Count - 1).Range.End)
        On Error GoTo 0
        If Not nonFnRange Is Nothing Then
            For i = 1 To 9
                numStr = "[" & CStr(i) & "]"
                EnsureSpaceBeforeCalloutInRange nonFnRange, numStr
            Next i
            For i = 1 To 26
                numStr = "[" & Chr(96 + i) & "]"
                EnsureSpaceBeforeCalloutInRange nonFnRange, numStr
            Next i
        End If
    End If
    
    ' 2. Convert numeric bracket callouts "[1]".."[9]" -> "a".."i" (superscript, preserving any preceding space)
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
    
    ' 3. Convert letter bracket callouts "[a]".."[z]" -> "a".."z" (superscript, preserving any preceding space)
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

Private Sub EnsureSpaceBeforeCalloutInRange(ByVal targetRange As Range, ByVal bracketToken As String)
    Dim rSearch As Range
    Dim prevChar As String
    
    Set rSearch = targetRange.Duplicate
    With rSearch.Find
        .ClearFormatting
        .Text = bracketToken
        .Forward = True
        .Wrap = 0 ' wdFindStop
        .MatchCase = True
        .MatchWholeWord = False
        
        Do While .Execute
            If rSearch.Start >= targetRange.End Then Exit Do
            If rSearch.Start > targetRange.Start Then
                prevChar = targetRange.Document.Range(rSearch.Start - 1, rSearch.Start).Text
                If prevChar <> " " And prevChar <> Chr(160) And prevChar <> vbTab And _
                   prevChar <> vbCr And prevChar <> vbLf And prevChar <> Chr(11) And prevChar <> Chr(7) Then
                    rSearch.InsertBefore " "
                End If
            End If
            rSearch.Collapse 0 ' wdCollapseEnd
            If rSearch.Start >= targetRange.End Then Exit Do
        Loop
    End With
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
        Set fnRange = tbl.Rows(footnoteRowIdx).Cells(1).Range
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
    Do While InStr(res, ChrW(8805) & " ") > 0: res = Replace(res, ChrW(8805) & " ", ChrW(8805)): Loop
    Do While InStr(res, ChrW(8804) & " ") > 0: res = Replace(res, ChrW(8804) & " ", ChrW(8804)): Loop
    Do While InStr(res, "> ") > 0: res = Replace(res, "> ", ">"): Loop
    Do While InStr(res, "< ") > 0: res = Replace(res, "< ", "<"): Loop
    Do While InStr(res, "  ") > 0
        res = Replace(res, "  ", " ")
    Loop
    CleanParagraphString = Trim(res)
End Function

Private Function CleanHeaderCellString(ByVal s As String) As String
    Dim res As String
    res = Replace(s, Chr(7), "")
    
    Do While Right(res, 1) = Chr(13) Or Right(res, 1) = Chr(10)
        res = Left(res, Len(res) - 1)
    Loop
    
    res = Replace(res, vbCrLf, Chr(11))
    res = Replace(res, Chr(13), Chr(11))
    res = Replace(res, Chr(10), Chr(11))
    res = Replace(res, Chr(160), " ")
    res = Replace(res, ">=", ChrW(8805))
    res = Replace(res, "=>", ChrW(8805))
    res = Replace(res, "<=", ChrW(8804))
    res = Replace(res, "=<", ChrW(8804))
    Do While InStr(res, " =") > 0: res = Replace(res, " =", "="): Loop
    Do While InStr(res, "= ") > 0: res = Replace(res, "= ", "="): Loop
    
    Do While InStr(res, " " & Chr(11)) > 0: res = Replace(res, " " & Chr(11), Chr(11)): Loop
    Do While InStr(res, Chr(11) & " ") > 0: res = Replace(res, Chr(11) & " ", Chr(11)): Loop
    Do While InStr(res, Chr(11) & Chr(11)) > 0: res = Replace(res, Chr(11) & Chr(11), Chr(11)): Loop
    
    Do While InStr(res, "   ") > 0: res = Replace(res, "   ", Chr(11)): Loop
    Do While InStr(res, "  ") > 0: res = Replace(res, "  ", " "): Loop
    
    CleanHeaderCellString = Trim(res)
End Function

Private Sub ScanHeaderFooterMetadata(ByVal doc As Document, ByRef outTitleBlock As String, ByVal fnColl As Collection)
    Dim sec As Section
    Dim hfIdx As Long
    Dim hf As HeaderFooter
    
    On Error Resume Next
    For Each sec In doc.Sections
        For hfIdx = 1 To 3 ' primary, first page, even pages
            Set hf = Nothing
            Set hf = sec.Headers(hfIdx)
            If Not hf Is Nothing Then
                If hf.Exists And Len(outTitleBlock) = 0 Then
                    outTitleBlock = ExtractTitleBlockFromStory(hf.Range)
                End If
            End If
            
            Set hf = Nothing
            Set hf = sec.Footers(hfIdx)
            If Not hf Is Nothing Then
                If hf.Exists Then
                    AddFooterFootnotes hf.Range, fnColl
                End If
            End If
        Next hfIdx
    Next sec
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ExtractTitleBlockFromStory(ByVal rng As Range) As String
    Dim p As Paragraph
    Dim parts() As String
    Dim k As Long
    Dim ln As String
    Dim blk As String
    Dim started As Boolean
    
    blk = ""
    started = False
    For Each p In rng.Paragraphs
        parts = Split(Replace(Replace(p.Range.Text, Chr(7), ""), Chr(13), ""), Chr(11))
        For k = LBound(parts) To UBound(parts)
            ln = Trim(CleanParagraphString(parts(k)))
            If Len(ln) > 0 Then
                If Not started Then
                    If IsTableNumberLine(ln) Then
                        started = True
                        blk = ln
                    End If
                ElseIf Not IsFooterBoilerplateLine(ln) And Not IsSubgroupByline(ln) Then
                    blk = blk & Chr(11) & ln
                End If
            End If
        Next k
    Next p
    ExtractTitleBlockFromStory = blk
End Function

Private Function IsSubgroupByline(ByVal s As String) As Boolean
    Dim t As String
    Dim colonPos As Long
    t = LCase(Trim(s))
    IsSubgroupByline = False
    If Len(t) = 0 Then Exit Function
    
    ' Exclude header/footer boilerplate and population lines
    If IsFooterBoilerplateLine(t) Then Exit Function
    If IsTableNumberLine(s) Then Exit Function
    If Left(t, 11) = "population:" Or Left(t, 13) = "analysis set:" Then Exit Function
    If Left(t, 5) = "note:" Or Left(t, 7) = "source:" Then Exit Function
    
    ' Any SAS BY-line in the header formatted as "Category: Value" (e.g. Treatment:, Age:, Gender:, Race:, BMI:, etc.)
    colonPos = InStr(t, ":")
    If colonPos >= 2 And colonPos <= 45 And colonPos < Len(t) Then
        IsSubgroupByline = True
    End If
End Function

Private Sub PreScanSectionSubgroups(ByVal doc As Document, ByRef outSubgroups() As String)
    Dim secCount As Long
    Dim sIdx As Long
    Dim hfIdx As Long
    Dim hf As HeaderFooter
    Dim p As Paragraph
    Dim parts() As String
    Dim k As Long
    Dim ln As String
    Dim foundSub As String
    Dim lastSub As String
    
    On Error Resume Next
    secCount = doc.Sections.Count
    If secCount <= 0 Then
        ReDim outSubgroups(1 To 1)
        outSubgroups(1) = ""
        On Error GoTo 0
        Exit Sub
    End If
    
    ReDim outSubgroups(1 To secCount)
    lastSub = ""
    
    For sIdx = 1 To secCount
        foundSub = ""
        For hfIdx = 1 To 3
            Set hf = Nothing
            Set hf = doc.Sections(sIdx).Headers(hfIdx)
            If Not hf Is Nothing Then
                If hf.Exists Then
                    For Each p In hf.Range.Paragraphs
                        parts = Split(Replace(Replace(p.Range.Text, Chr(7), ""), Chr(13), ""), Chr(11))
                        For k = LBound(parts) To UBound(parts)
                            ln = Trim(CleanParagraphString(parts(k)))
                            If IsSubgroupByline(ln) Then
                                foundSub = ln
                                Exit For
                            End If
                        Next k
                        If Len(foundSub) > 0 Then Exit For
                    Next p
                End If
            End If
            If Len(foundSub) > 0 Then Exit For
        Next hfIdx
        
        If Len(foundSub) > 0 Then
            lastSub = foundSub
        End If
        outSubgroups(sIdx) = lastSub
    Next sIdx
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub AddFooterFootnotes(ByVal rng As Range, ByVal fnColl As Collection)
    Dim p As Paragraph
    Dim parts() As String
    Dim k As Long
    Dim ln As String
    
    For Each p In rng.Paragraphs
        parts = Split(Replace(Replace(p.Range.Text, Chr(7), ""), Chr(13), ""), Chr(11))
        For k = LBound(parts) To UBound(parts)
            ln = Trim(CleanParagraphString(parts(k)))
            If Len(ln) > 0 Then
                If Not IsFooterBoilerplateLine(ln) Then
                    If Not CollectionContains(fnColl, ln) Then
                        fnColl.Add ln
                    End If
                End If
            End If
        Next k
    Next p
End Sub

Private Function IsTableNumberLine(ByVal s As String) As Boolean
    Dim w As Variant
    Dim n As Long
    IsTableNumberLine = False
    For Each w In Array("Table ", "Listing ", "Figure ")
        n = Len(w)
        If Len(s) > n Then
            If StrComp(Left(s, n), CStr(w), vbTextCompare) = 0 Then
                If Mid(s, n + 1, 1) Like "#" Then
                    IsTableNumberLine = True
                    Exit Function
                End If
            End If
        End If
    Next w
End Function

Private Function IsFooterBoilerplateLine(ByVal s As String) As Boolean
    Dim t As String
    t = LCase(Trim(s))
    IsFooterBoilerplateLine = True
    If Len(t) = 0 Then Exit Function
    If InStr(t, "generated on") > 0 Then Exit Function
    If InStr(t, "program:") > 0 Then Exit Function
    If InStr(t, "executed:") > 0 Then Exit Function
    If InStr(t, "protocol:") > 0 Then Exit Function
    If InStr(t, "/programs/") > 0 Or InStr(t, "\programs\") > 0 Then Exit Function
    If InStr(t, ".sas") > 0 Then Exit Function
    If Left(t, 3) = "../" Or Left(t, 1) = "/" Then Exit Function
    If Left(t, 5) = "page " And InStr(t, " of ") > 0 Then Exit Function
    IsFooterBoilerplateLine = False
End Function

Private Function JoinTitleLines(ByVal prevLine As String, ByVal nextLine As String) As String
    ' Separate SAS TITLE lines are joined with ", " unless the previous line already ends
    ' in punctuation, ends with a connector word (manual wrap), or the next line starts lowercase.
    Dim lastCh As String
    Dim lastWord As String
    Dim sp As Long
    Dim firstCh As String
    
    lastCh = Right(prevLine, 1)
    sp = InStrRev(prevLine, " ")
    lastWord = LCase(Mid(prevLine, sp + 1))
    firstCh = Left(nextLine, 1)
    
    If InStr(1, ",;:-(/", lastCh) > 0 Then
        JoinTitleLines = prevLine & " " & nextLine
    ElseIf InStr(1, "|and|or|of|by|in|for|the|with|to|per|from|at|on|vs|versus|", "|" & lastWord & "|") > 0 Then
        JoinTitleLines = prevLine & " " & nextLine
    ElseIf firstCh Like "[a-z(]" Then
        JoinTitleLines = prevLine & " " & nextLine
    Else
        JoinTitleLines = prevLine & ", " & nextLine
    End If
End Function

Private Sub AddFootnotesFromText(ByVal fnColl As Collection, ByVal rawText As String)
    Dim norm As String
    norm = Replace(Replace(Replace(rawText, Chr(7), ""), vbCrLf, vbCr), Chr(11), vbCr)
    norm = Replace(norm, vbLf, vbCr)
    
    Dim lines() As String
    lines = Split(norm, vbCr)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim cleanL As String
        cleanL = CleanParagraphString(lines(i))
        If Len(cleanL) > 0 Then
            If IsFootnoteParagraph(cleanL) Then
                If Not CollectionContains(fnColl, cleanL) Then
                    fnColl.Add cleanL
                End If
            End If
        End If
    Next i
End Sub

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
    Do While InStr(res, "  ") > 0: res = Replace(res, "  ", " "): Loop
    CleanCellString = Trim(res)
End Function

Private Function CollectionContains(ByVal coll As Collection, ByVal keyVal As String) As Boolean
    Dim item As Variant
    CollectionContains = False
    If coll Is Nothing Then Exit Function
    For Each item In coll
        If StrComp(CStr(item), keyVal, vbTextCompare) = 0 Then
            CollectionContains = True
            Exit Function
        End If
    Next item
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
                mainTitle = JoinTitleLines(mainTitle, validLines(i))
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
                If IsNumberedDesignatorWord(currToken) And NextTokenStartsWithDigit(s, i) Then
                    res = res & UCase(Left(currToken, 1)) & LCase(Mid(currToken, 2))
                Else
                    res = res & FormatWordSentenceCase(currToken, Not firstWordDone)
                End If
                If HasAlphaChar(currToken) Or HasDigits(currToken) Then firstWordDone = True
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

Private Function IsNumberedDesignatorWord(ByVal w As String) As Boolean
    Select Case LCase(w)
        Case "week", "weeks", "day", "days", "cycle", "cycles", "month", "months", _
             "visit", "visits", "grade", "grades", "part", "parts", "cohort", "cohorts", _
             "phase", "phases", "stage", "stages", "group", "groups", "arm", "arms", "period", "periods"
            IsNumberedDesignatorWord = True
        Case Else
            IsNumberedDesignatorWord = False
    End Select
End Function

Private Function NextTokenStartsWithDigit(ByVal s As String, ByVal startPos As Long) As Boolean
    Dim idx As Long
    Dim ch As String
    NextTokenStartsWithDigit = False
    idx = startPos
    Do While idx <= Len(s)
        ch = Mid(s, idx, 1)
        If ch = " " Or ch = Chr(160) Or ch = vbTab Then
            idx = idx + 1
        Else
            Exit Do
        End If
    Loop
    If idx <= Len(s) Then
        ch = Mid(s, idx, 1)
        If ch >= "0" And ch <= "9" Then
            NextTokenStartsWithDigit = True
        End If
    End If
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
        ' Preserve plural acronyms and mixed-case abbreviations with 2+ uppercase letters (e.g. TEAEs, SAEs, AEs, AESIs, NSAIDs)
        If numUpper >= 2 And InStr(baseW, "-") = 0 Then
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
                ElseIf HasMultipleUpperLetters(pStr) Then
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

Private Function HasMultipleUpperLetters(ByVal s As String) As Boolean
    Dim i As Long, ch As String, uCount As Long
    uCount = 0
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "A" And ch <= "Z" Then
            uCount = uCount + 1
            If uCount >= 2 Then
                HasMultipleUpperLetters = True
                Exit Function
            End If
        End If
    Next i
    HasMultipleUpperLetters = False
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
        Case "teae": GetProperNoun = "TEAE"
        Case "teaes": GetProperNoun = "TEAEs"
        Case "tesae": GetProperNoun = "TESAE"
        Case "tesaes": GetProperNoun = "TESAEs"
        Case "sae": GetProperNoun = "SAE"
        Case "saes": GetProperNoun = "SAEs"
        Case "ae": GetProperNoun = "AE"
        Case "aes": GetProperNoun = "AEs"
        Case "aesi": GetProperNoun = "AESI"
        Case "aesis": GetProperNoun = "AESIs"
        Case "imp": GetProperNoun = "IMP"
        Case "imps": GetProperNoun = "IMPs"
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

' ==============================================================================
' MULTI-SELECT TABLE SELECTION & LONG INCIDENCE TABLE THRESHOLD FILTERING
' ==============================================================================

Private Function IsContinuationStubRow(rowData() As String, ByVal colCount As Long) As Boolean
    Dim stub As String
    Dim c As Long
    IsContinuationStubRow = False
    stub = LCase(Trim(rowData(1)))
    If Right(stub, 7) = "(cont.)" Or Right(stub, 11) = "(continued)" Then
        For c = 2 To colCount
            If Len(Trim(rowData(c))) > 0 Then
                Exit Function
            End If
        Next c
        IsContinuationStubRow = True
    End If
End Function

Private Function PromptMultiTableSelection(ByVal allTables As Collection) As Collection
    Dim res As Collection
    Dim promptMsg As String
    Dim pIdx As Long
    Dim ptDict As Object
    Dim totalCount As Long
    Dim userChoice As String
    
    Set res = New Collection
    totalCount = allTables.Count
    promptMsg = "Found " & totalCount & " distinct tables/subgroup panels:" & vbCrLf & _
                "Enter number(s)/range (e.g. '1', '5-8', '1-2, 5-8'), keyword (e.g. 'Race'), or 'All':" & vbCrLf & vbCrLf
    
    If totalCount <= 15 Then
        For pIdx = 1 To totalCount
            Set ptDict = allTables(pIdx)
            Dim panelLabel As String
            If Len(ptDict("Subtitle")) > 0 Then
                panelLabel = SanitizeInputBoxText(CStr(ptDict("Subtitle"))) & " [" & ptDict("BodyRows").Count & " rows]"
            Else
                panelLabel = SanitizeInputBoxText(CStr(ptDict("FullTitle"))) & " [" & ptDict("BodyRows").Count & " rows]"
            End If
            promptMsg = promptMsg & pIdx & ": " & panelLabel & vbCrLf
        Next pIdx
    Else
        ' Group by category prefix (before ':') and keep concise so InputBox never truncates at 1023 chars
        Dim currCat As String: currCat = ""
        Dim catStart As Long: catStart = 1
        Dim catEnd As Long: catEnd = 1
        Dim catVals As String: catVals = ""
        
        For pIdx = 1 To totalCount
            Set ptDict = allTables(pIdx)
            Dim subStr As String: subStr = SanitizeInputBoxText(Trim(CStr(ptDict("Subtitle"))))
            Dim colonPos As Long: colonPos = InStr(subStr, ":")
            Dim thisCat As String, thisVal As String
            If colonPos > 0 Then
                thisCat = ShortenPromptCategory(Trim(Left(subStr, colonPos - 1)))
                thisVal = ShortenPromptValue(Trim(Mid(subStr, colonPos + 1)))
            Else
                thisCat = ShortenPromptCategory(subStr)
                thisVal = ""
            End If
            
            If pIdx = 1 Then
                currCat = thisCat
                catStart = 1
                catEnd = 1
                catVals = "1: " & thisVal
            ElseIf StrComp(thisCat, currCat, vbTextCompare) = 0 Then
                catEnd = pIdx
                catVals = catVals & " | " & pIdx & ": " & thisVal
            Else
                If catStart = catEnd Then
                    promptMsg = promptMsg & "[" & catStart & "] " & currCat & " (" & catVals & ")" & vbCrLf
                Else
                    promptMsg = promptMsg & "[" & catStart & "-" & catEnd & "] " & currCat & " (" & catVals & ")" & vbCrLf
                End If
                currCat = thisCat
                catStart = pIdx
                catEnd = pIdx
                catVals = pIdx & ": " & thisVal
            End If
        Next pIdx
        If catStart = catEnd Then
            promptMsg = promptMsg & "[" & catStart & "] " & currCat & " (" & catVals & ")" & vbCrLf
        Else
            promptMsg = promptMsg & "[" & catStart & "-" & catEnd & "] " & currCat & " (" & catVals & ")" & vbCrLf
        End If
    End If
    
    userChoice = Trim(InputBox(promptMsg, "Select CSR Table(s) to Insert", "All"))
    If Len(userChoice) = 0 Then
        Set PromptMultiTableSelection = res
        Exit Function
    End If
    
    Set res = ParseMultiTableSelection(userChoice, allTables)
    If res.Count = 0 Then
        MsgBox "Invalid table selection: '" & userChoice & "'.", vbExclamation, "Takeda CSR Tools"
    End If
    Set PromptMultiTableSelection = res
End Function

Private Function SanitizeInputBoxText(ByVal s As String) As String
    Dim res As String
    res = Replace(s, ChrW(8805), ">=")
    res = Replace(res, ChrW(8804), "<=")
    SanitizeInputBoxText = res
End Function

Private Function ShortenPromptCategory(ByVal s As String) As String
    Dim res As String
    res = s
    res = Replace(res, "Current use of ", "", 1, -1, vbTextCompare)
    res = Replace(res, "Baseline presence of ", "Baseline ", 1, -1, vbTextCompare)
    res = Replace(res, "Geographic Region", "Region", 1, -1, vbTextCompare)
    ShortenPromptCategory = res
End Function

Private Function ShortenPromptValue(ByVal s As String) As String
    Dim res As String
    res = s
    res = Replace(res, "BSA Affected by PsO", "BSA", 1, -1, vbTextCompare)
    res = Replace(res, "Black Or African American", "Black/AA", 1, -1, vbTextCompare)
    res = Replace(res, "Median of TJC (median=12)", "Med(12)", 1, -1, vbTextCompare)
    res = Replace(res, "Median of SJC (median=7)", "Med(7)", 1, -1, vbTextCompare)
    If Len(res) > 22 Then
        res = Left(res, 20) & ".."
    End If
    ShortenPromptValue = res
End Function

Private Function ParseMultiTableSelection(ByVal userInput As String, ByVal allTables As Collection) As Collection
    Dim res As Collection
    Dim seen As Object
    Dim cleaned As String
    Dim tokens() As String
    Dim tok As Variant
    Dim tStr As String
    Dim dashPos As Long
    Dim startIdx As Long, endIdx As Long, k As Long
    Dim maxCount As Long
    
    Set res = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    maxCount = allTables.Count
    
    cleaned = Trim(userInput)
    If UCase(cleaned) = "ALL" Then
        For k = 1 To maxCount
            res.Add allTables(k)
        Next k
        Set ParseMultiTableSelection = res
        Exit Function
    End If
    
    cleaned = Replace(cleaned, ";", ",")
    tokens = Split(cleaned, ",")
    
    For Each tok In tokens
        tStr = Trim(CStr(tok))
        ' Strip brackets if user typed '[1]' or '[5-8]'
        If Left(tStr, 1) = "[" And Right(tStr, 1) = "]" And Len(tStr) > 2 Then
            tStr = Trim(Mid(tStr, 2, Len(tStr) - 2))
        End If
        
        ' Strip '1:...' or '1=...' if user copied value suffix along with index
        Dim eqOrColPos As Long
        eqOrColPos = InStr(tStr, "=")
        If eqOrColPos = 0 Then eqOrColPos = InStr(tStr, ":")
        If eqOrColPos > 1 Then
            If IsNumeric(Trim(Left(tStr, eqOrColPos - 1))) Then
                tStr = Trim(Left(tStr, eqOrColPos - 1))
            End If
        End If
        
        If Len(tStr) > 0 Then
            dashPos = InStr(tStr, "-")
            If dashPos > 1 And dashPos < Len(tStr) Then
                Dim leftStr As String, rightStr As String
                leftStr = Trim(Left(tStr, dashPos - 1))
                rightStr = Trim(Mid(tStr, dashPos + 1))
                If IsNumeric(leftStr) And IsNumeric(rightStr) Then
                    startIdx = CLng(leftStr)
                    endIdx = CLng(rightStr)
                    If startIdx <= endIdx Then
                        For k = startIdx To endIdx
                            If k >= 1 And k <= maxCount Then
                                If Not seen.Exists(CStr(k)) Then
                                    seen.Add CStr(k), True
                                    res.Add allTables(k)
                                End If
                            End If
                        Next k
                    End If
                End If
            ElseIf IsNumeric(tStr) Then
                k = CLng(tStr)
                If k >= 1 And k <= maxCount Then
                    If Not seen.Exists(CStr(k)) Then
                        seen.Add CStr(k), True
                        res.Add allTables(k)
                    End If
                End If
            Else
                ' Match by subgroup/subtitle keyword (e.g. 'Race', 'Gender', 'Age', 'Placebo')
                For k = 1 To maxCount
                    Dim tblObj As Object: Set tblObj = allTables(k)
                    Dim subText As String: subText = CStr(tblObj("Subtitle"))
                    If InStr(1, subText, tStr, vbTextCompare) > 0 Then
                        If Not seen.Exists(CStr(k)) Then
                            seen.Add CStr(k), True
                            res.Add tblObj
                        End If
                    End If
                Next k
            End If
        End If
    Next tok
    
    Set ParseMultiTableSelection = res
End Function

Private Sub PromptAndFilterSelectedTables(ByVal tablesToInsert As Collection)
    Dim tblDict As Variant
    Dim hasLongIncidenceTable As Boolean
    Dim maxRows As Long
    
    hasLongIncidenceTable = False
    maxRows = 0
    For Each tblDict In tablesToInsert
        If tblDict("BodyRows").Count >= 40 Then
            If IsIncidenceCountTable(tblDict) Then
                hasLongIncidenceTable = True
                If tblDict("BodyRows").Count > maxRows Then
                    maxRows = tblDict("BodyRows").Count
                End If
            End If
        End If
    Next tblDict
    
    If Not hasLongIncidenceTable Then Exit Sub
    
    Dim filterPrompt As String
    filterPrompt = "Long table detected (" & maxRows & " data rows per panel)." & vbCrLf & vbCrLf & _
                   "Choose a row filtering option (applied independently to each selected treatment/table):" & vbCrLf & _
                   "  - '0' or Blank : Keep all rows (no filtering)" & vbCrLf & _
                   "  - '>0' or '1'  : Remove rows with 0 across all columns/grades" & vbCrLf & _
                   "  - 'n' (e.g. 2, 5) : Keep rows with >= n participants in any column/grade" & vbCrLf & _
                   "  - 'x%' (e.g. 1%, 2%) : Keep rows with >= x% of participants in any column/grade" & vbCrLf & vbCrLf & _
                   "Enter filter threshold:"
    
    Dim filterInput As String
    filterInput = Trim(InputBox(filterPrompt, "Long CSR Table Row Filter", "0"))
    
    If Len(filterInput) = 0 Or filterInput = "0" Or UCase(filterInput) = "ALL" Or UCase(filterInput) = "NONE" Then
        Exit Sub
    End If
    
    Dim filterMode As String ' "GT_ZERO", "MIN_COUNT", "MIN_PCT"
    Dim minCount As Long
    Dim minPct As Double
    
    ParseFilterThresholdInput filterInput, filterMode, minCount, minPct
    If Len(filterMode) = 0 Then Exit Sub
    
    For Each tblDict In tablesToInsert
        If tblDict("BodyRows").Count >= 40 And IsIncidenceCountTable(tblDict) Then
            ApplyHierarchyRowFilterToTable tblDict, filterMode, minCount, minPct
            UpdateFilteredTableTitle tblDict, filterMode, minCount, minPct
        End If
    Next tblDict
End Sub

Private Function IsIncidenceCountTable(ByVal tblDict As Object) As Boolean
    Dim hVals As Variant
    Dim r As Long, c As Long
    Dim hText As String
    Dim bodyRows As Collection
    Dim rItem As Variant, rData As Variant
    Dim stub As String
    
    IsIncidenceCountTable = False
    On Error Resume Next
    
    ' 1. If body rows contain continuous summary statistic stubs (Mean (SD), Median, Q1, Q3, LS Mean), it is NOT an incidence table
    Set bodyRows = tblDict("BodyRows")
    If Not bodyRows Is Nothing Then
        For Each rItem In bodyRows
            rData = rItem(0)
            stub = LCase(Trim(CStr(rData(1))))
            If Left(stub, 9) = "mean (sd)" Or Left(stub, 6) = "median" Or _
               Left(stub, 6) = "q1, q3" Or Left(stub, 8) = "min, max" Or _
               Left(stub, 7) = "ls mean" Then
                On Error GoTo 0
                Exit Function
            End If
        Next rItem
    End If
    
    ' 2. Check header cells for 'n (%)' or 'Grade '
    hVals = tblDict("HeaderValues")
    For r = LBound(hVals, 1) To UBound(hVals, 1)
        For c = LBound(hVals, 2) To UBound(hVals, 2)
            hText = LCase(Trim(CStr(hVals(r, c))))
            If InStr(hText, "n (%)") > 0 Or InStr(hText, "grade ") > 0 Then
                IsIncidenceCountTable = True
                On Error GoTo 0
                Exit Function
            End If
        Next c
    Next r
    On Error GoTo 0
End Function

Private Sub ParseFilterThresholdInput( _
    ByVal rawInput As String, _
    ByRef outMode As String, _
    ByRef outMinCount As Long, _
    ByRef outMinPct As Double _
)
    Dim s As String
    outMode = ""
    outMinCount = 0
    outMinPct = 0#
    
    s = Trim(rawInput)
    s = Replace(s, ChrW(8805), "")
    s = Replace(s, ">=", "")
    s = Replace(s, "=>", "")
    s = Trim(s)
    
    If s = ">0" Or s = "> 0" Or s = "1" Then
        outMode = "GT_ZERO"
        outMinCount = 1
        Exit Sub
    End If
    
    If Left(s, 1) = ">" Then
        s = Trim(Mid(s, 2))
        If Right(s, 1) = "%" Then
            Dim pctPartGt As String
            pctPartGt = Trim(Left(s, Len(s) - 1))
            If IsNumeric(pctPartGt) Then
                outMode = "MIN_PCT"
                outMinPct = CDbl(pctPartGt)
            End If
        ElseIf IsNumeric(s) Then
            Dim valGt As Long
            valGt = CLng(s)
            If valGt <= 0 Then
                outMode = "GT_ZERO"
                outMinCount = 1
            Else
                outMode = "MIN_COUNT"
                outMinCount = valGt + 1
            End If
        End If
        Exit Sub
    End If
    
    If Right(s, 1) = "%" Then
        Dim pctPart As String
        pctPart = Trim(Left(s, Len(s) - 1))
        If IsNumeric(pctPart) Then
            outMode = "MIN_PCT"
            outMinPct = CDbl(pctPart)
        End If
        Exit Sub
    End If
    
    If IsNumeric(s) Then
        Dim nVal As Long
        nVal = CLng(s)
        If nVal = 1 Then
            outMode = "GT_ZERO"
            outMinCount = 1
        ElseIf nVal > 1 Then
            outMode = "MIN_COUNT"
            outMinCount = nVal
        End If
    End If
End Sub

Private Sub ApplyHierarchyRowFilterToTable( _
    ByVal tblDict As Object, _
    ByVal filterMode As String, _
    ByVal minCount As Long, _
    ByVal minPct As Double _
)
    Dim srcRows As Collection
    Dim filteredRows As Collection
    Dim colCount As Long
    Dim totalRows As Long
    Dim i As Long, j As Long
    
    Set srcRows = tblDict("BodyRows")
    colCount = CLng(tblDict("ColCount"))
    totalRows = srcRows.Count
    If totalRows = 0 Then Exit Sub
    
    Set filteredRows = New Collection
    i = 1
    Do While i <= totalRows
        Dim rItem As Variant
        Dim rData As Variant
        Dim rInds As Variant
        Dim indentPts As Single
        
        rItem = srcRows(i)
        rData = rItem(0)
        rInds = rItem(1)
        indentPts = CSng(rInds(1))
        
        If indentPts <= 0.1 Then
            ' Unindented row: check if followed by indented child rows (e.g. SOC heading with Preferred Terms)
            Dim childrenIndices As Collection
            Set childrenIndices = New Collection
            j = i + 1
            Do While j <= totalRows
                Dim childItem As Variant
                Dim childInds As Variant
                childItem = srcRows(j)
                childInds = childItem(1)
                If CSng(childInds(1)) > 0.1 Then
                    childrenIndices.Add j
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            
            If childrenIndices.Count = 0 Then
                ' Standalone unindented row (e.g. 'Participants with any TEAE' or unindented SOC with 0s)
                If RowMeetsThreshold(rData, colCount, filterMode, minCount, minPct) Then
                    filteredRows.Add rItem
                End If
            Else
                ' Parent SOC heading with child Preferred Terms: keep parent if >= 1 child survives
                Dim keptChildren As Collection
                Set keptChildren = New Collection
                Dim cIdxVar As Variant
                For Each cIdxVar In childrenIndices
                    Dim cItem As Variant
                    Dim cData As Variant
                    cItem = srcRows(CLng(cIdxVar))
                    cData = cItem(0)
                    If RowMeetsThreshold(cData, colCount, filterMode, minCount, minPct) Then
                        keptChildren.Add cItem
                    End If
                Next cIdxVar
                
                If keptChildren.Count > 0 Then
                    filteredRows.Add rItem
                    Dim kcItem As Variant
                    For Each kcItem In keptChildren
                        filteredRows.Add kcItem
                    Next kcItem
                End If
            End If
            i = j
        Else
            ' Fallback if an indented row appears without an unindented parent
            If RowMeetsThreshold(rData, colCount, filterMode, minCount, minPct) Then
                filteredRows.Add rItem
            End If
            i = i + 1
        End If
    Loop
    
    tblDict.Remove "BodyRows"
    Set tblDict("BodyRows") = filteredRows
End Sub

Private Function RowMeetsThreshold( _
    ByRef rowData As Variant, _
    ByVal colCount As Long, _
    ByVal filterMode As String, _
    ByVal minCount As Long, _
    ByVal minPct As Double _
) As Boolean
    Dim c As Long
    Dim cellVal As String
    Dim nCount As Long
    Dim pctVal As Double
    
    RowMeetsThreshold = False
    For c = 2 To colCount
        cellVal = Trim(CStr(rowData(c)))
        ParseCountAndPercent cellVal, nCount, pctVal
        If filterMode = "GT_ZERO" Then
            If nCount >= 1 Or pctVal > 0# Then
                RowMeetsThreshold = True
                Exit Function
            End If
        ElseIf filterMode = "MIN_COUNT" Then
            If nCount >= minCount Then
                RowMeetsThreshold = True
                Exit Function
            End If
        ElseIf filterMode = "MIN_PCT" Then
            If nCount >= 1 And pctVal >= minPct Then
                RowMeetsThreshold = True
                Exit Function
            End If
        End If
    Next c
End Function

Private Sub ParseCountAndPercent(ByVal cellText As String, ByRef outCount As Long, ByRef outPct As Double)
    Dim s As String
    Dim openP As Long, closeP As Long
    Dim countPart As String, pctPart As String
    
    outCount = 0
    outPct = 0#
    s = Trim(cellText)
    If Len(s) = 0 Or s = "0" Or s = "-" Or s = "NA" Then Exit Sub
    
    openP = InStr(s, "(")
    If openP > 0 Then
        countPart = Trim(Left(s, openP - 1))
        closeP = InStr(openP, s, ")")
        If closeP > openP Then
            pctPart = Trim(Mid(s, openP + 1, closeP - openP - 1))
        Else
            pctPart = Trim(Mid(s, openP + 1))
        End If
        pctPart = Replace(pctPart, "%", "")
        pctPart = Trim(pctPart)
        
        If IsNumeric(countPart) Then
            outCount = CLng(CDbl(countPart))
        End If
        
        If Left(pctPart, 1) = "<" Or Left(pctPart, 1) = ChrW(8804) Then
            ' e.g. '<1' means > 0% but < 1%
            outPct = 0.5
        ElseIf Left(pctPart, 1) = ">" Or Left(pctPart, 1) = ChrW(8805) Then
            Dim numAfter As String
            numAfter = Trim(Mid(pctPart, 2))
            If IsNumeric(numAfter) Then outPct = CDbl(numAfter)
        ElseIf IsNumeric(pctPart) Then
            outPct = CDbl(pctPart)
        End If
    ElseIf IsNumeric(s) Then
        outCount = CLng(CDbl(s))
    End If
End Sub

Private Sub UpdateFilteredTableTitle( _
    ByVal tblDict As Object, _
    ByVal filterMode As String, _
    ByVal minCount As Long, _
    ByVal minPct As Double _
)
    Dim baseT As String
    Dim subT As String
    Dim popT As String
    Dim hasGrades As Boolean
    Dim suffixUnit As String
    Dim filterClause As String
    Dim hVals As Variant
    Dim r As Long, c As Long
    
    baseT = Trim(CStr(tblDict("Title")))
    subT = Trim(CStr(tblDict("Subtitle")))
    popT = Trim(CStr(tblDict("Population")))
    
    If Len(popT) > 0 Then
        If Right(baseT, Len(popT)) = popT Then
            baseT = Trim(Left(baseT, Len(baseT) - Len(popT)))
        End If
    End If
    
    hasGrades = (InStr(1, baseT, "Grade", vbTextCompare) > 0)
    If Not hasGrades Then
        On Error Resume Next
        hVals = tblDict("HeaderValues")
        For r = LBound(hVals, 1) To UBound(hVals, 1)
            For c = LBound(hVals, 2) To UBound(hVals, 2)
                If InStr(1, CStr(hVals(r, c)), "Grade", vbTextCompare) > 0 Then
                    hasGrades = True
                    Exit For
                End If
            Next c
            If hasGrades Then Exit For
        Next r
        On Error GoTo 0
    End If
    
    If hasGrades Then
        suffixUnit = " in Any Grade"
    Else
        suffixUnit = " in Any Group"
    End If
    
    If filterMode = "GT_ZERO" Then
        filterClause = " (Occurring in " & ChrW(8805) & "1 Participant" & suffixUnit & ")"
    ElseIf filterMode = "MIN_COUNT" Then
        filterClause = " (Occurring in " & ChrW(8805) & CStr(minCount) & " Participants" & suffixUnit & ")"
    ElseIf filterMode = "MIN_PCT" Then
        filterClause = " (Occurring in " & ChrW(8805) & CStr(minPct) & "% of Participants" & suffixUnit & ")"
    Else
        filterClause = ""
    End If
    
    Dim fullT As String
    fullT = baseT & filterClause
    If Len(subT) > 0 Then fullT = fullT & " - " & subT
    If Len(popT) > 0 Then fullT = fullT & " " & popT
    tblDict("FullTitle") = fullT
End Sub

' ==============================================================================
' PRE-FLIGHT RAW RTF GUARD & PRE-INSERTION CONFIRMATION HELPERS
' ==============================================================================

Private Function ExtractFolderPath(ByVal fullPath As String) As String
    Dim lastSep As Long
    lastSep = InStrRev(fullPath, "\")
    If lastSep = 0 Then lastSep = InStrRev(fullPath, "/")
    If lastSep > 0 Then
        ExtractFolderPath = Left(fullPath, lastSep)
    Else
        ExtractFolderPath = ""
    End If
End Function

Private Function ExtractFileName(ByVal fullPath As String) As String
    Dim lastSep As Long
    lastSep = InStrRev(fullPath, "\")
    If lastSep = 0 Then lastSep = InStrRev(fullPath, "/")
    If lastSep > 0 Then
        ExtractFileName = Mid(fullPath, lastSep + 1)
    Else
        ExtractFileName = fullPath
    End If
End Function

Private Function CheckRTFPreFlightBeforeOpen(ByVal rtfPath As String) As Boolean
    Dim fileBytes As Long
    Dim estSourcePages As Long
    Dim estRawRows As Long
    Dim quickTitle As String
    
    CheckRTFPreFlightBeforeOpen = True
    On Error Resume Next
    
    ' Only run raw RTF inspection on .rtf/.doc text streams (skip binary .docx zip archives)
    If LCase(Right(rtfPath, 5)) = ".docx" Then
        On Error GoTo 0
        Exit Function
    End If
    
    QuickInspectRTFFile rtfPath, fileBytes, estSourcePages, estRawRows, quickTitle
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' Trigger the pre-open warning if the source RTF is large (>= 30 pages, >= 500 raw rows, or >= 750 KB)
    If estSourcePages >= 30 Or estRawRows >= 500 Or fileBytes >= 768000 Then
        Dim sizeStr As String
        If fileBytes >= 1048576 Then
            sizeStr = Format(CDbl(fileBytes) / 1048576#, "0.00") & " MB"
        Else
            sizeStr = CLng(fileBytes / 1024) & " KB"
        End If
        
        Dim scanSecLow As Long
        Dim scanSecHigh As Long
        scanSecLow = CLng(estSourcePages * 0.1) + CLng(estRawRows * 0.002)
        If scanSecLow < 2 Then scanSecLow = 2
        scanSecHigh = CLng(scanSecLow * 1.6) + 2
        
        Dim warnMsg As String
        warnMsg = "Large RTF file detected before opening in Word:" & vbCrLf & vbCrLf & _
                  "File: " & ExtractFileName(rtfPath) & " (" & sizeStr & ")" & vbCrLf & _
                  "Detected Table: " & quickTitle & vbCrLf & _
                  "Source RTF Length: ~" & estSourcePages & " RTF pages (~" & estRawRows & " raw table rows)" & vbCrLf & _
                  "Estimated Scan Time: " & FormatEstimatedSeconds(scanSecLow, scanSecHigh) & vbCrLf & vbCrLf & _
                  "Note: If this table has multiple treatment/subgroup panels or is a long incidence table, " & _
                  "you will be able to select specific panels or filter rows after scanning." & vbCrLf & _
                  "If you selected the wrong file (e.g., a massive shift or by-visit table), click Cancel now to abort immediately." & vbCrLf & vbCrLf & _
                  "Proceed with opening and scanning this file?"
        
        Dim resp As VbMsgBoxResult
        resp = MsgBox(warnMsg, vbOKCancel + vbExclamation, "Large RTF File Detected - Confirm Scan")
        If resp <> vbOK Then
            CheckRTFPreFlightBeforeOpen = False
        End If
    End If
End Function

Private Sub QuickInspectRTFFile( _
    ByVal rtfPath As String, _
    ByRef outFileBytes As Long, _
    ByRef outPages As Long, _
    ByRef outRawRows As Long, _
    ByRef outQuickTitle As String _
)
    Dim fNum As Integer
    Dim headLen As Long
    Dim headBuf As String
    Dim scanLen As Long
    Dim scanBuf As String
    Dim sectdCount As Long
    Dim pageCount As Long
    
    outFileBytes = 0
    outPages = 1
    outRawRows = 0
    outQuickTitle = "Clinical Summary Table"
    
    On Error GoTo QuickCleanup
    outFileBytes = FileLen(rtfPath)
    If outFileBytes <= 0 Then Exit Sub
    
    fNum = FreeFile
    Open rtfPath For Binary Access Read Shared As #fNum
    
    ' Read first 32 KB for instant title extraction
    headLen = outFileBytes
    If headLen > 32768 Then headLen = 32768
    headBuf = Space$(headLen)
    Get #fNum, 1, headBuf
    outQuickTitle = ExtractQuickTitleFromRawRTF(headBuf)
    
    ' Read up to 8 MB in a single fast block to count \sectd and \row tokens
    scanLen = outFileBytes
    If scanLen > 8388608 Then scanLen = 8388608
    If scanLen <= headLen Then
        scanBuf = headBuf
    Else
        scanBuf = Space$(scanLen)
        Get #fNum, 1, scanBuf
    End If
    Close #fNum
    fNum = 0
    
    sectdCount = CountSubstringOccurrences(scanBuf, "\sectd")
    pageCount = CountSubstringOccurrences(scanBuf, "\page")
    If sectdCount >= pageCount And sectdCount > 0 Then
        outPages = sectdCount
    ElseIf pageCount > 0 Then
        outPages = pageCount + 1
    Else
        outPages = 1
    End If
    
    outRawRows = CountSubstringOccurrences(scanBuf, "\row")
    
    ' Scale estimate proportionally if file exceeds 8 MB
    If outFileBytes > scanLen And scanLen > 0 Then
        Dim scaleFactor As Double
        scaleFactor = CDbl(outFileBytes) / CDbl(scanLen)
        outPages = CLng(outPages * scaleFactor)
        outRawRows = CLng(outRawRows * scaleFactor)
    End If
    Exit Sub
    
QuickCleanup:
    If fNum <> 0 Then
        On Error Resume Next
        Close #fNum
    End If
End Sub

Private Function CountSubstringOccurrences(ByRef haystack As String, ByVal needle As String) As Long
    Dim cnt As Long
    Dim pos As Long
    Dim nLen As Long
    
    cnt = 0
    nLen = Len(needle)
    If nLen = 0 Or Len(haystack) = 0 Then
        CountSubstringOccurrences = 0
        Exit Function
    End If
    
    pos = InStr(1, haystack, needle, vbBinaryCompare)
    Do While pos > 0
        cnt = cnt + 1
        pos = InStr(pos + nLen, haystack, needle, vbBinaryCompare)
    Loop
    CountSubstringOccurrences = cnt
End Function

Private Function ExtractQuickTitleFromRawRTF(ByRef rawHead As String) As String
    Dim pos As Long
    Dim kw As Variant
    Dim idx As Long
    Dim p As Long
    Dim nxtCh As String
    
    pos = 0
    For Each kw In Array("Table ", "Listing ")
        idx = 1
        Do
            p = InStr(idx, rawHead, CStr(kw), vbBinaryCompare)
            If p = 0 Then Exit Do
            If p + Len(CStr(kw)) <= Len(rawHead) Then
                nxtCh = Mid$(rawHead, p + Len(CStr(kw)), 1)
                If nxtCh >= "0" And nxtCh <= "9" Then
                    If pos = 0 Or p < pos Then
                        pos = p
                    End If
                    Exit Do
                End If
            End If
            idx = p + Len(CStr(kw))
        Loop
    Next kw
    
    If pos = 0 Then
        ExtractQuickTitleFromRawRTF = "Clinical Summary Table"
        Exit Function
    End If
    
    Dim endCellOrRow As Long
    endCellOrRow = InStr(pos, rawHead, "\cell}", vbBinaryCompare)
    Dim endBrace As Long
    endBrace = InStr(pos, rawHead, "}", vbBinaryCompare)
    
    Dim firstEnd As Long
    If endCellOrRow > 0 And endCellOrRow < pos + 600 Then
        firstEnd = endCellOrRow
    ElseIf endBrace > 0 And endBrace < pos + 600 Then
        firstEnd = endBrace - 1
    Else
        firstEnd = pos + 300
    End If
    If firstEnd > Len(rawHead) Then firstEnd = Len(rawHead)
    
    Dim firstSeg As String
    firstSeg = Mid$(rawHead, pos, firstEnd - pos + 1)
    
    If InStr(1, firstSeg, "\line", vbBinaryCompare) > 0 Then
        ExtractQuickTitleFromRawRTF = StripRawRTFControls(firstSeg)
        Exit Function
    End If
    
    ' Multi-row header table (each title line is in its own {\cell} row)
    Dim combined As String
    combined = StripRawRTFControls(firstSeg)
    
    Dim scanPos As Long
    scanPos = firstEnd + 1
    Dim stepIdx As Long
    For stepIdx = 1 To 3
        Dim nextCell As Long
        nextCell = InStr(scanPos, rawHead, "\cell}", vbBinaryCompare)
        If nextCell = 0 Or (nextCell - scanPos) > 650 Then
            Exit For
        End If
        Dim openB As Long
        openB = InStrRev(rawHead, "{", nextCell, vbBinaryCompare)
        If openB >= scanPos Then
            Dim segTxt As String
            segTxt = StripRawRTFControls(Mid$(rawHead, openB + 1, nextCell - openB - 1))
            If Len(segTxt) > 0 And Left$(segTxt, 5) <> "Page " And Left$(segTxt, 8) <> "Takeda: " Then
                combined = combined & " - " & segTxt
            End If
        End If
        scanPos = nextCell + 6
    Next stepIdx
    
    If Len(combined) = 0 Then combined = "Clinical Summary Table"
    ExtractQuickTitleFromRawRTF = combined
End Function

Private Function StripRawRTFControls(ByVal s As String) As String
    Dim work As String
    Dim out As String
    Dim i As Long
    Dim n As Long
    Dim ch As String
    
    work = Replace(s, "\line ", " - ")
    work = Replace(work, "\line", " - ")
    work = Replace(work, "\cell", "")
    
    out = ""
    n = Len(work)
    i = 1
    Do While i <= n
        ch = Mid$(work, i, 1)
        If ch = "\" Then
            i = i + 1
            Do While i <= n
                ch = Mid$(work, i, 1)
                If (ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "*" Then
                    i = i + 1
                Else
                    Exit Do
                End If
            Loop
            If i <= n Then
                If Mid$(work, i, 1) = " " Then i = i + 1
            End If
        ElseIf ch = "{" Or ch = "}" Or ch = vbCr Or ch = vbLf Then
            i = i + 1
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    
    out = Trim(out)
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    StripRawRTFControls = out
End Function

Private Function ConfirmTableInsertionWithUser( _
    ByVal rtfPath As String, _
    ByVal tablesToInsert As Collection, _
    ByVal defaultSourceNumber As String _
) As Boolean
    Dim tblCount As Long
    Dim totalDataRows As Long
    Dim totalOrigRows As Long
    Dim totalCells As Long
    Dim totalEstPages As Long
    Dim anyLandscape As Boolean
    Dim anyPortrait As Boolean
    Dim firstTbl As Object
    Dim tItem As Variant
    Dim tObj As Object
    
    ConfirmTableInsertionWithUser = True
    tblCount = tablesToInsert.Count
    If tblCount = 0 Then
        ConfirmTableInsertionWithUser = False
        Exit Function
    End If
    
    totalDataRows = 0
    totalOrigRows = 0
    totalCells = 0
    totalEstPages = 0
    anyLandscape = False
    anyPortrait = False
    
    Set firstTbl = tablesToInsert(1)
    For Each tItem In tablesToInsert
        Set tObj = tItem
        Dim rCnt As Long: rCnt = tObj("BodyRows").Count
        Dim oCnt As Long: oCnt = rCnt
        If tObj.Exists("OriginalRowCount") Then
            oCnt = CLng(tObj("OriginalRowCount"))
        End If
        Dim cCnt As Long: cCnt = CLng(tObj("ColCount"))
        Dim hCnt As Long: hCnt = CLng(tObj("NumHeaderRows"))
        Dim wPts As Single: wPts = CSng(tObj("TotalWidthPts"))
        Dim isLand As Boolean
        isLand = (cCnt >= CSR_AUTO_LANDSCAPE_COL_THRESHOLD Or wPts > CSR_PORTRAIT_MAX_WIDTH_PTS)
        If isLand Then
            anyLandscape = True
        Else
            anyPortrait = True
        End If
        
        totalDataRows = totalDataRows + rCnt
        totalOrigRows = totalOrigRows + oCnt
        totalCells = totalCells + ((rCnt + hCnt + 2) * cCnt)
        totalEstPages = totalEstPages + EstimateCSROutputPages(rCnt, hCnt, isLand)
    Next tItem
    
    Dim orientLabel As String
    If anyLandscape And anyPortrait Then
        orientLabel = "Mixed Landscape & Portrait"
    ElseIf anyLandscape Then
        orientLabel = "Landscape"
    Else
        orientLabel = "Portrait"
    End If
    
    Dim srcNum As String
    srcNum = CStr(firstTbl("SourceNumber"))
    If Len(srcNum) = 0 Then srcNum = defaultSourceNumber
    If Len(srcNum) = 0 Then srcNum = "N/A"
    
    Dim titleSummary As String
    If tblCount = 1 Then
        titleSummary = CStr(firstTbl("FullTitle"))
    Else
        titleSummary = CStr(firstTbl("Title"))
        If Len(CStr(firstTbl("Population"))) > 0 Then
            If InStr(1, titleSummary, CStr(firstTbl("Population")), vbTextCompare) = 0 Then
                titleSummary = titleSummary & " " & CStr(firstTbl("Population"))
            End If
        End If
        
        Dim panelPreview As String: panelPreview = ""
        Dim pIdx As Long
        Dim maxPreview As Long: maxPreview = tblCount
        If maxPreview > 4 Then maxPreview = 4
        For pIdx = 1 To maxPreview
            Set tObj = tablesToInsert(pIdx)
            Dim subName As String: subName = Trim(CStr(tObj("Subtitle")))
            If Len(subName) = 0 Then subName = "Panel " & pIdx
            If pIdx > 1 Then panelPreview = panelPreview & "; "
            panelPreview = panelPreview & subName & " (" & tObj("BodyRows").Count & " rows)"
        Next pIdx
        If tblCount > 4 Then
            panelPreview = panelPreview & "; ... (+" & (tblCount - 4) & " more)"
        End If
        titleSummary = titleSummary & vbCrLf & "Selected Panels (" & tblCount & "): " & panelPreview
    End If
    
    Dim rowsSummary As String
    If totalOrigRows > totalDataRows Then
        rowsSummary = totalDataRows & " data rows (filtered down from " & totalOrigRows & " original rows) x " & firstTbl("ColCount") & " columns"
    Else
        rowsSummary = totalDataRows & " data rows x " & firstTbl("ColCount") & " columns"
    End If
    
    ' Estimate insertion & formatting duration in Word
    Dim estSecLow As Long
    Dim estSecHigh As Long
    estSecLow = CLng(totalCells * 0.0025) + CLng(tblCount * 0.4)
    If estSecLow < 1 Then estSecLow = 1
    estSecHigh = CLng(estSecLow * 1.6) + 1
    
    Dim msg As String
    msg = "Ready to insert into CSR document:" & vbCrLf & vbCrLf & _
          "File: " & ExtractFileName(rtfPath) & vbCrLf & _
          "Source: " & srcNum & vbCrLf & _
          "Title: " & titleSummary & vbCrLf & vbCrLf & _
          "Table Size: " & rowsSummary & vbCrLf & _
          "Estimated CSR Length: ~" & totalEstPages & " page(s) (" & orientLabel & ")" & vbCrLf & _
          "Estimated Insertion Time: " & FormatEstimatedSeconds(estSecLow, estSecHigh)
    
    Dim boxStyle As VbMsgBoxStyle
    If totalDataRows >= 150 Or totalEstPages >= 6 Then
        msg = msg & vbCrLf & vbCrLf & _
              "WARNING: This insertion will add ~" & totalEstPages & " pages (" & totalDataRows & " rows) to your CSR document."
        boxStyle = vbOKCancel + vbExclamation
    Else
        boxStyle = vbOKCancel + vbInformation
    End If
    
    msg = msg & vbCrLf & vbCrLf & "Click OK to insert into the CSR, or Cancel to abort."
    
    Dim choice As VbMsgBoxResult
    choice = MsgBox(msg, boxStyle, "Confirm CSR Table Insertion")
    If choice <> vbOK Then
        ConfirmTableInsertionWithUser = False
    End If
End Function

Private Function EstimateCSROutputPages(ByVal dataRows As Long, ByVal headerRows As Long, ByVal isLandscape As Boolean) As Long
    Dim rowsPerPage As Long
    If isLandscape Then
        rowsPerPage = 26 - headerRows
    Else
        rowsPerPage = 38 - headerRows
    End If
    If rowsPerPage < 12 Then rowsPerPage = 12
    
    Dim pages As Long
    pages = CLng((dataRows + 3 + rowsPerPage - 1) \ rowsPerPage)
    If pages < 1 Then pages = 1
    EstimateCSROutputPages = pages
End Function

Private Function FormatEstimatedSeconds(ByVal secLow As Long, ByVal secHigh As Long) As String
    If secHigh <= 2 Then
        FormatEstimatedSeconds = "< 2 seconds"
    ElseIf secHigh < 60 Then
        FormatEstimatedSeconds = "~" & secLow & "-" & secHigh & " seconds"
    Else
        Dim minLow As Long: minLow = secLow \ 60
        Dim minHigh As Long: minHigh = (secHigh + 59) \ 60
        If minLow < 1 Then minLow = 1
        FormatEstimatedSeconds = "~" & minLow & "-" & minHigh & " minute(s)"
    End If
End Function



