"""Programmatic MSForms builder for the Takeda CSR Table Import dialog."""

FRM_CSR_TABLE_VBA = r'''Option Explicit

Private mAllTables As Collection
Private mResultTables As Collection
Private mCancelled As Boolean
Private mRtfPath As String
Private mSourceNum As String
Private mUpdatingUI As Boolean
Private mRefColCount As Long
Private mHasIncidence As Boolean
Private mHasTotalCol As Boolean

Public Property Get Cancelled() As Boolean
    Cancelled = mCancelled
End Property

Public Property Get ResultTables() As Collection
    Set ResultTables = mResultTables
End Property

Public Sub InitializeDialog( _
    ByVal rtfPath As String, _
    ByVal allTables As Collection, _
    ByVal defaultSourceNum As String _
)
    Dim i As Long
    Dim tblObj As Object
    Dim srcNum As String
    Dim canCombine As Boolean
    Dim cIdx As Long
    
    mUpdatingUI = True
    mCancelled = True
    mRtfPath = rtfPath
    Set mAllTables = allTables
    Set mResultTables = New Collection
    mSourceNum = defaultSourceNum
    
    Set tblObj = allTables(1)
    srcNum = CStr(tblObj("SourceNumber"))
    If Len(srcNum) = 0 Then srcNum = defaultSourceNum
    If Len(srcNum) = 0 Then srcNum = "N/A"
    mSourceNum = srcNum
    
    Me.lblHeaderSub.Caption = "Source: " & srcNum & "  |  File: " & Mod_CSR_Table.CSR_ExtractFileName(rtfPath)
    
    ' 1. Populate Panels list (all checked by default)
    Me.lstPanels.Clear
    For i = 1 To allTables.Count
        Me.lstPanels.AddItem Mod_CSR_Table.CSR_GetPanelDisplayLabel(allTables(i), i)
        Me.lstPanels.Selected(i - 1) = True
    Next i
    
    canCombine = Mod_CSR_Table.CSR_CanCombineTables(allTables)
    If allTables.Count <= 1 Then
        Me.fraPanels.Caption = " 1. Source Table (1 table detected) "
        Me.txtKeyword.Enabled = False
        Me.btnMatchKeyword.Enabled = False
        Me.btnSelectAllPanels.Enabled = False
        Me.btnClearAllPanels.Enabled = False
        Me.optSeparate.Value = True
        Me.optSeparate.Enabled = False
        Me.optCombine.Enabled = False
    Else
        Me.fraPanels.Caption = " 1. Subtables / Subgroup Panels (" & allTables.Count & " detected) "
        Me.txtKeyword.Enabled = True
        Me.btnMatchKeyword.Enabled = True
        Me.btnSelectAllPanels.Enabled = True
        Me.btnClearAllPanels.Enabled = True
        Me.optSeparate.Value = True
        Me.optSeparate.Enabled = canCombine
        Me.optCombine.Enabled = canCombine
    End If
    
    ' 2. Populate Data Columns list (Columns 2..ColCount, all checked by default)
    Me.lstColumns.Clear
    mRefColCount = CLng(tblObj("ColCount"))
    For cIdx = 2 To mRefColCount
        Me.lstColumns.AddItem Mod_CSR_Table.CSR_GetColumnDisplayLabel(tblObj, cIdx)
        Me.lstColumns.Selected(cIdx - 2) = True
    Next cIdx
    Me.btnAllCols.Enabled = (mRefColCount > 2)
    
    ' 3. Configure Row Filtering section
    mHasIncidence = Mod_CSR_Table.CSR_AnyTableIsIncidence(allTables)
    mHasTotalCol = Mod_CSR_Table.CSR_AnyTableHasTotalCol(allTables)
    
    Me.optFilterNone.Value = True
    Me.txtMinCount.Text = "5"
    Me.txtMinPct.Text = "5"
    Me.optScopeGroup.Value = True
    
    Me.optFilterNone.Enabled = mHasIncidence
    Me.optFilterGtZero.Enabled = mHasIncidence
    Me.optFilterCount.Enabled = mHasIncidence
    Me.txtMinCount.Enabled = mHasIncidence
    Me.optFilterPct.Enabled = mHasIncidence
    Me.txtMinPct.Enabled = mHasIncidence
    Me.lblPctSign.Enabled = mHasIncidence
    
    If mHasIncidence Then
        Me.fraFilter.Caption = " 3. Row Threshold Filtering (Incidence / Grade Table) "
    Else
        Me.fraFilter.Caption = " 3. Row Threshold Filtering (N/A for Continuous / Listing Tables) "
    End If
    
    Me.lblScope.Enabled = (mHasIncidence And mHasTotalCol)
    Me.optScopeGroup.Enabled = (mHasIncidence And mHasTotalCol)
    Me.optScopeTotal.Enabled = (mHasIncidence And mHasTotalCol)
    If Not mHasTotalCol Then
        Me.optScopeGroup.Caption = "Any group (no Total col detected)"
    Else
        Me.optScopeGroup.Caption = "Any group (exclude Total column)"
    End If
    
    mUpdatingUI = False
    RefreshPreview
End Sub

Private Function GetSelectedPanelIndices() As Collection
    Dim coll As Collection
    Dim i As Long
    Set coll = New Collection
    For i = 0 To Me.lstPanels.ListCount - 1
        If Me.lstPanels.Selected(i) Then
            coll.Add (i + 1)
        End If
    Next i
    Set GetSelectedPanelIndices = coll
End Function

Private Function GetSelectedDataColumnIndices() As Collection
    Dim coll As Collection
    Dim i As Long
    Set coll = New Collection
    For i = 0 To Me.lstColumns.ListCount - 1
        If Me.lstColumns.Selected(i) Then
            coll.Add (i + 2)
        End If
    Next i
    Set GetSelectedDataColumnIndices = coll
End Function

Private Sub ReadFilterSettings( _
    ByRef outMode As String, _
    ByRef outMinCount As Long, _
    ByRef outMinPct As Double, _
    ByRef outScope As String _
)
    Dim cleanNum As String
    outMode = ""
    outMinCount = 0
    outMinPct = 0#
    outScope = "NON_TOTAL"
    
    If Not mHasIncidence Then Exit Sub
    If Me.optFilterGtZero.Value Then
        outMode = "GT_ZERO"
        outMinCount = 1
    ElseIf Me.optFilterCount.Value Then
        cleanNum = Trim(Me.txtMinCount.Text)
        If IsNumeric(cleanNum) Then
            outMinCount = CLng(CDbl(cleanNum))
            If outMinCount = 1 Then
                outMode = "GT_ZERO"
            ElseIf outMinCount > 1 Then
                outMode = "MIN_COUNT"
            End If
        End If
    ElseIf Me.optFilterPct.Value Then
        cleanNum = Replace(Trim(Me.txtMinPct.Text), "%", "")
        If IsNumeric(cleanNum) Then
            outMinPct = CDbl(cleanNum)
            If outMinPct > 0# Then
                outMode = "MIN_PCT"
            End If
        End If
    End If
    
    If mHasTotalCol And Me.optScopeTotal.Value Then
        outScope = "TOTAL_ONLY"
    Else
        outScope = "NON_TOTAL"
    End If
End Sub

Private Sub RefreshPreview()
    Dim panelIdxs As Collection
    Dim colIdxs As Collection
    Dim filterMode As String
    Dim minCount As Long
    Dim minPct As Double
    Dim colScope As String
    Dim titlePrev As String
    Dim statsPrev As String
    Dim combineFlag As Boolean
    Dim configured As Collection
    
    If mUpdatingUI Then Exit Sub
    If mAllTables Is Nothing Then Exit Sub
    
    On Error GoTo PreviewSafe
    Set panelIdxs = GetSelectedPanelIndices()
    Set colIdxs = GetSelectedDataColumnIndices()
    
    ' Update combine option enablement based on current panel count selection
    If panelIdxs.Count >= 2 And Mod_CSR_Table.CSR_CanCombineTables(mAllTables) Then
        Me.optSeparate.Enabled = True
        Me.optCombine.Enabled = True
    Else
        Me.optSeparate.Enabled = False
        Me.optCombine.Enabled = False
    End If
    
    ' Update Total-scope enablement when a filter other than Keep All is active
    Dim hasActiveFilter As Boolean
    hasActiveFilter = (mHasIncidence And Not Me.optFilterNone.Value)
    Dim totalCheckedInUI As Boolean
    totalCheckedInUI = (mHasIncidence And mHasTotalCol And hasActiveFilter)
    Me.lblScope.Enabled = totalCheckedInUI
    Me.optScopeGroup.Enabled = totalCheckedInUI
    Me.optScopeTotal.Enabled = totalCheckedInUI
    
    If panelIdxs.Count = 0 Or colIdxs.Count = 0 Then
        Me.lblPreviewTitle.Caption = "Please select at least 1 panel and 1 data column."
        Me.lblPreviewStats.Caption = "0 tables ready to insert"
        Me.btnInsert.Enabled = False
        Exit Sub
    End If
    
    ReadFilterSettings filterMode, minCount, minPct, colScope
    combineFlag = (Me.optCombine.Enabled And Me.optCombine.Value)
    
    Set configured = Mod_CSR_Table.CSR_BuildConfiguredTables( _
        mAllTables, panelIdxs, combineFlag, colIdxs, filterMode, minCount, minPct, colScope)
    Set mResultTables = configured
    
    statsPrev = Mod_CSR_Table.CSR_FormatPreviewSummary(configured, titlePrev)
    If Len(titlePrev) > 105 Then titlePrev = Left(titlePrev, 102) & "..."
    Me.lblPreviewTitle.Caption = titlePrev
    Me.lblPreviewStats.Caption = statsPrev
    Me.btnInsert.Enabled = (configured.Count > 0)
    Exit Sub

PreviewSafe:
    Err.Clear
End Sub

Private Sub lstPanels_Change()
    RefreshPreview
End Sub

Private Sub lstColumns_Change()
    RefreshPreview
End Sub

Private Sub optSeparate_Click()
    RefreshPreview
End Sub

Private Sub optCombine_Click()
    RefreshPreview
End Sub

Private Sub optFilterNone_Click()
    RefreshPreview
End Sub

Private Sub optFilterGtZero_Click()
    RefreshPreview
End Sub

Private Sub optFilterCount_Click()
    RefreshPreview
End Sub

Private Sub optFilterPct_Click()
    RefreshPreview
End Sub

Private Sub optScopeGroup_Click()
    RefreshPreview
End Sub

Private Sub optScopeTotal_Click()
    RefreshPreview
End Sub

Private Sub txtMinCount_Change()
    If mUpdatingUI Then Exit Sub
    If Not Me.optFilterCount.Value Then
        mUpdatingUI = True
        Me.optFilterCount.Value = True
        mUpdatingUI = False
    End If
    RefreshPreview
End Sub

Private Sub txtMinPct_Change()
    If mUpdatingUI Then Exit Sub
    If Not Me.optFilterPct.Value Then
        mUpdatingUI = True
        Me.optFilterPct.Value = True
        mUpdatingUI = False
    End If
    RefreshPreview
End Sub

Private Sub btnSelectAllPanels_Click()
    Dim i As Long
    mUpdatingUI = True
    For i = 0 To Me.lstPanels.ListCount - 1
        Me.lstPanels.Selected(i) = True
    Next i
    mUpdatingUI = False
    RefreshPreview
End Sub

Private Sub btnClearAllPanels_Click()
    Dim i As Long
    mUpdatingUI = True
    For i = 0 To Me.lstPanels.ListCount - 1
        Me.lstPanels.Selected(i) = False
    Next i
    mUpdatingUI = False
    RefreshPreview
End Sub

Private Sub btnMatchKeyword_Click()
    Dim kw As String
    Dim i As Long
    Dim matched As Long
    kw = Trim(Me.txtKeyword.Text)
    If Len(kw) = 0 Then
        btnSelectAllPanels_Click
        Exit Sub
    End If
    
    mUpdatingUI = True
    matched = 0
    For i = 0 To Me.lstPanels.ListCount - 1
        If InStr(1, CStr(Me.lstPanels.List(i)), kw, vbTextCompare) > 0 Then
            Me.lstPanels.Selected(i) = True
            matched = matched + 1
        Else
            Me.lstPanels.Selected(i) = False
        End If
    Next i
    mUpdatingUI = False
    RefreshPreview
    If matched = 0 Then
        Me.lblPreviewTitle.Caption = "No panels matched keyword '" & kw & "'. Try another keyword or click Select All."
    End If
End Sub

Private Sub txtKeyword_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        btnMatchKeyword_Click
    End If
End Sub

Private Sub btnAllCols_Click()
    Dim i As Long
    mUpdatingUI = True
    For i = 0 To Me.lstColumns.ListCount - 1
        Me.lstColumns.Selected(i) = True
    Next i
    mUpdatingUI = False
    RefreshPreview
End Sub

Private Sub btnInsert_Click()
    RefreshPreview
    If mResultTables Is Nothing Then Exit Sub
    If mResultTables.Count = 0 Then
        MsgBox "Please select at least 1 panel and 1 data column before inserting.", vbExclamation, "Takeda CSR Tools"
        Exit Sub
    End If
    mCancelled = False
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    mCancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then ' vbFormControlMenu (X button)
        Cancel = 1
        mCancelled = True
        Me.Hide
    End If
End Sub
'''


def _ole_rgb(r: int, g: int, b: int) -> int:
    return (b << 16) | (g << 8) | r


def _style_font(ctrl, size: float = 9.0, bold: bool = False):
    ctrl.Font.Name = "Segoe UI"
    ctrl.Font.Size = size
    ctrl.Font.Bold = bold


def _add_label(container, name: str, left: float, top: float, width: float, height: float,
               caption: str = "", size: float = 9.0, bold: bool = False,
               fore_color: int = None, back_color: int = None,
               transparent: bool = True, word_wrap: bool = False):
    lbl = container.Controls.Add("Forms.Label.1", name, True)
    lbl.Left = left
    lbl.Top = top
    lbl.Width = width
    lbl.Height = height
    lbl.Caption = caption
    _style_font(lbl, size=size, bold=bold)
    lbl.BackStyle = 0 if transparent else 1
    lbl.WordWrap = word_wrap
    if fore_color is not None:
        lbl.ForeColor = fore_color
    if back_color is not None:
        lbl.BackColor = back_color
    return lbl


def _add_button(container, name: str, left: float, top: float, width: float, height: float,
                caption: str, size: float = 8.5, bold: bool = False,
                back_color: int = None, fore_color: int = None,
                is_default: bool = False, is_cancel: bool = False, tip: str = ""):
    btn = container.Controls.Add("Forms.CommandButton.1", name, True)
    btn.Left = left
    btn.Top = top
    btn.Width = width
    btn.Height = height
    btn.Caption = caption
    _style_font(btn, size=size, bold=bold)
    if back_color is not None:
        btn.BackColor = back_color
    if fore_color is not None:
        btn.ForeColor = fore_color
    if is_default:
        btn.Default = True
    if is_cancel:
        btn.Cancel = True
    if tip:
        btn.ControlTipText = tip
    return btn


def _add_frame(designer, name: str, left: float, top: float, width: float, height: float, caption: str):
    fra = designer.Controls.Add("Forms.Frame.1", name, True)
    fra.Left = left
    fra.Top = top
    fra.Width = width
    fra.Height = height
    fra.Caption = caption
    fra.BackColor = _ole_rgb(255, 255, 255)
    fra.ForeColor = _ole_rgb(24, 58, 102)
    fra.SpecialEffect = 0  # fmSpecialEffectFlat
    fra.BorderStyle = 1    # fmBorderStyleSingle
    fra.BorderColor = _ole_rgb(208, 215, 224)
    _style_font(fra, size=9.0, bold=True)
    return fra


def _add_option(container, name: str, left: float, top: float, width: float, height: float,
                caption: str, group_name: str, back_color: int):
    opt = container.Controls.Add("Forms.OptionButton.1", name, True)
    opt.Left = left
    opt.Top = top
    opt.Width = width
    opt.Height = height
    opt.Caption = caption
    opt.GroupName = group_name
    opt.BackColor = back_color
    opt.ForeColor = _ole_rgb(28, 36, 48)
    _style_font(opt, size=8.5, bold=False)
    return opt


def _add_textbox(container, name: str, left: float, top: float, width: float, height: float,
                 text: str = "", tip: str = ""):
    tb = container.Controls.Add("Forms.TextBox.1", name, True)
    tb.Left = left
    tb.Top = top
    tb.Width = width
    tb.Height = height
    tb.Text = text
    tb.SpecialEffect = 0   # fmSpecialEffectFlat
    tb.BorderStyle = 1     # fmBorderStyleSingle
    tb.BorderColor = _ole_rgb(190, 200, 212)
    tb.BackColor = _ole_rgb(255, 255, 255)
    tb.ForeColor = _ole_rgb(28, 36, 48)
    _style_font(tb, size=8.5, bold=False)
    if tip:
        tb.ControlTipText = tip
    return tb


def add_csr_table_userform(vb_comp) -> None:
    """Create and configure the FrmCSRTable UserForm inside a Word VBProject."""
    comp = vb_comp.Add(3)  # 3 = vbext_ct_MSForm
    comp.Name = "FrmCSRTable"

    bg_surface = _ole_rgb(244, 246, 249)
    navy_banner = _ole_rgb(24, 58, 102)
    accent_blue = _ole_rgb(0, 102, 184)
    card_white = _ole_rgb(255, 255, 255)
    text_dark = _ole_rgb(28, 36, 48)
    text_muted = _ole_rgb(95, 106, 122)

    comp.Properties("Caption").Value = "Takeda CSR Table Importer"
    comp.Properties("Width").Value = 560
    comp.Properties("Height").Value = 516
    comp.Properties("BackColor").Value = bg_surface
    comp.Properties("StartUpPosition").Value = 1  # CenterOwner

    designer = comp.Designer

    # --- Header Banner ---
    _add_label(
        designer, "lblHeaderBg", 0, 0, 560, 50,
        caption="", back_color=navy_banner, transparent=False
    )
    _add_label(
        designer, "lblHeaderTitle", 14, 7, 520, 18,
        caption="Insert Takeda CSR Table",
        size=12.0, bold=True, fore_color=_ole_rgb(255, 255, 255), transparent=True
    )
    _add_label(
        designer, "lblHeaderSub", 14, 27, 520, 16,
        caption="",
        size=8.5, bold=False, fore_color=_ole_rgb(212, 224, 238), transparent=True
    )

    # --- Section 1: Subtables / Panels ---
    fra_panels = _add_frame(
        designer, "fraPanels", 12, 58, 520, 148,
        caption=" 1. Subtables / Subgroup Panels "
    )
    _add_label(
        fra_panels, "lblSearch", 10, 17, 30, 14,
        caption="Filter:", size=8.5, fore_color=text_dark, transparent=True
    )
    _add_textbox(
        fra_panels, "txtKeyword", 42, 14, 140, 19,
        text="", tip="Type a panel keyword (e.g. Race, Cytopenia) and click Select Matching"
    )
    _add_button(
        fra_panels, "btnMatchKeyword", 188, 14, 96, 19,
        caption="Select Matching", back_color=bg_surface, fore_color=text_dark
    )
    _add_button(
        fra_panels, "btnSelectAllPanels", 368, 14, 66, 19,
        caption="Select All", back_color=bg_surface, fore_color=text_dark
    )
    _add_button(
        fra_panels, "btnClearAllPanels", 440, 14, 68, 19,
        caption="Clear All", back_color=bg_surface, fore_color=text_dark
    )

    lst_panels = fra_panels.Controls.Add("Forms.ListBox.1", "lstPanels", True)
    lst_panels.Left = 10
    lst_panels.Top = 38
    lst_panels.Width = 498
    lst_panels.Height = 82
    lst_panels.ListStyle = 1       # fmListStyleOption (checkboxes)
    lst_panels.MultiSelect = 1     # fmMultiSelectMulti
    lst_panels.IntegralHeight = False
    lst_panels.SpecialEffect = 0
    lst_panels.BorderStyle = 1
    lst_panels.BorderColor = _ole_rgb(190, 200, 212)
    lst_panels.ForeColor = text_dark
    _style_font(lst_panels, size=9.0, bold=False)

    _add_label(
        fra_panels, "lblCombineMode", 10, 126, 98, 15,
        caption="Multi-Panel Mode:", size=8.5, bold=True, fore_color=text_dark, transparent=True
    )
    _add_option(
        fra_panels, "optSeparate", 112, 124, 156, 18,
        caption="Insert as separate tables", group_name="grpCombine", back_color=card_white
    )
    _add_option(
        fra_panels, "optCombine", 272, 124, 238, 18,
        caption="Combine into 1 table (panels as parent rows)", group_name="grpCombine", back_color=card_white
    )

    # --- Section 2: Data Columns ---
    fra_cols = _add_frame(
        designer, "fraColumns", 12, 212, 520, 84,
        caption=" 2. Data Columns to Include "
    )
    lst_cols = fra_cols.Controls.Add("Forms.ListBox.1", "lstColumns", True)
    lst_cols.Left = 10
    lst_cols.Top = 14
    lst_cols.Width = 420
    lst_cols.Height = 62
    lst_cols.ListStyle = 1         # fmListStyleOption (checkboxes)
    lst_cols.MultiSelect = 1       # fmMultiSelectMulti
    lst_cols.IntegralHeight = False
    lst_cols.SpecialEffect = 0
    lst_cols.BorderStyle = 1
    lst_cols.BorderColor = _ole_rgb(190, 200, 212)
    lst_cols.ForeColor = text_dark
    _style_font(lst_cols, size=8.5, bold=False)

    _add_button(
        fra_cols, "btnAllCols", 438, 16, 70, 20,
        caption="All Columns", back_color=bg_surface, fore_color=text_dark
    )
    _add_label(
        fra_cols, "lblColHint", 438, 40, 72, 36,
        caption="Row stub (Col 1) always kept",
        size=7.5, fore_color=text_muted, transparent=True, word_wrap=True
    )

    # --- Section 3: Row Filtering ---
    fra_filter = _add_frame(
        designer, "fraFilter", 12, 302, 520, 76,
        caption=" 3. Row Threshold Filtering "
    )
    _add_option(
        fra_filter, "optFilterNone", 10, 16, 90, 18,
        caption="Keep all rows", group_name="grpFilter", back_color=card_white
    )
    _add_option(
        fra_filter, "optFilterGtZero", 104, 16, 106, 18,
        caption="Remove all-0 (>0)", group_name="grpFilter", back_color=card_white
    )
    _add_option(
        fra_filter, "optFilterCount", 216, 16, 112, 18,
        caption="Min participants >=", group_name="grpFilter", back_color=card_white
    )
    _add_textbox(
        fra_filter, "txtMinCount", 330, 16, 34, 18,
        text="5", tip="Minimum participant count n"
    )
    _add_option(
        fra_filter, "optFilterPct", 376, 16, 76, 18,
        caption="Min % >=", group_name="grpFilter", back_color=card_white
    )
    _add_textbox(
        fra_filter, "txtMinPct", 454, 16, 34, 18,
        text="5", tip="Minimum percentage of participants"
    )
    _add_label(
        fra_filter, "lblPctSign", 491, 18, 14, 14,
        caption="%", size=8.5, fore_color=text_dark, transparent=True
    )

    _add_label(
        fra_filter, "lblScope", 10, 46, 110, 16,
        caption="Total Column Scope:", size=8.5, bold=True, fore_color=text_dark, transparent=True
    )
    _add_option(
        fra_filter, "optScopeGroup", 122, 44, 188, 18,
        caption="Any group (exclude Total column)", group_name="grpScope", back_color=card_white
    )
    _add_option(
        fra_filter, "optScopeTotal", 316, 44, 160, 18,
        caption="Total column only", group_name="grpScope", back_color=card_white
    )

    # --- Preview & Action Buttons ---
    fra_prev = _add_frame(
        designer, "fraPreview", 12, 384, 520, 56,
        caption=" Output Preview "
    )
    _add_label(
        fra_prev, "lblPreviewTitle", 10, 15, 498, 18,
        caption="", size=8.5, bold=True, fore_color=text_dark, transparent=True
    )
    _add_label(
        fra_prev, "lblPreviewStats", 10, 35, 498, 16,
        caption="", size=8.5, bold=False, fore_color=navy_banner, transparent=True
    )

    _add_button(
        designer, "btnCancel", 354, 448, 84, 26,
        caption="Cancel", size=9.5, bold=False,
        back_color=card_white, fore_color=text_dark, is_cancel=True
    )
    _add_button(
        designer, "btnInsert", 446, 448, 86, 26,
        caption="Insert Table", size=9.5, bold=True,
        back_color=accent_blue, fore_color=_ole_rgb(255, 255, 255), is_default=True
    )

    code_mod = comp.CodeModule
    if code_mod.CountOfLines > 0:
        code_mod.DeleteLines(1, code_mod.CountOfLines)
    code_mod.AddFromString(FRM_CSR_TABLE_VBA.strip() + "\r\n")
