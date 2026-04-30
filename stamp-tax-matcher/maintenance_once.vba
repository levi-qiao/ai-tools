'========================================
' 印花税工具 - 一次性维护/清理宏
'========================================
' 用途：维护人员交付前或排查问题时导入使用。
' 日常用户只需要 base1.vba 里的 RunMatch。
'========================================

Sub SetupWorkbookOnce()
    Dim wsMapping As Worksheet
    Dim wsSummary As Worksheet
    Dim wsInvoice As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrorHandler

    Set wsMapping = ThisWorkbook.Worksheets("税目映射规则")
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")
    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Call NormalizeMappingNonTaxableLabels(wsMapping)
    Call EnsureTaxCatalogSheet(wsMapping, wsSummary)
    Call SetMappingRuleValidation(wsMapping)
    Call SetSummaryInputValidation(wsSummary)

    lastRow = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 3 Then
        Call SetManualConfirmValidation(wsInvoice, lastRow - 2)
        Call InstallManualConfirmationFormulas(wsInvoice, lastRow - 2)
        Call SetInvoiceConditionalFormatting(wsInvoice, lastRow - 2)
    End If

    Call GenerateSummaryConfirmation(wsInvoice, wsSummary)
    Application.Calculate

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "初始化完成。后续日常使用只需点击开始匹配。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "初始化失败：" & Err.Description, vbCritical, "错误"
End Sub

Sub OptimizeWorkbookLayout()
    Call ClearInvoiceColors(False)
    Call ClearSummaryColors(False)
    Call RefreshAfterManualConfirmation(False)
    MsgBox "优化完成：已清理旧颜色，并重建公式、下拉和条件格式。", vbInformation, "完成"
End Sub

Sub ClearInvoiceColors(Optional showMessage As Boolean = True)
    Dim wsInvoice As Worksheet
    Dim lastRow As Long

    On Error GoTo ErrorHandler
    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
    lastRow = LastUsedRowBetweenColumns(wsInvoice, 3, 1, 42)

    If lastRow < 3 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    wsInvoice.Rows("3:" & lastRow).FormatConditions.Delete
    wsInvoice.Rows("3:" & lastRow).Interior.Pattern = xlNone
    Call SafeHideColumns(wsInvoice, 36, 42)

    Call SetInvoiceConditionalFormatting(wsInvoice, lastRow - 2)

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If showMessage Then MsgBox "发票明细颜色已清理并重新按规则标色。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "清理发票明细颜色失败：" & Err.Description, vbCritical, "错误"
End Sub

Sub ClearSummaryColors(Optional showMessage As Boolean = True)
    Dim wsSummary As Worksheet

    On Error GoTo ErrorHandler
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    wsSummary.Rows("5:300").FormatConditions.Delete
    wsSummary.Rows("5:300").Interior.Pattern = xlNone
    Call RefreshQuarterSummary

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If showMessage Then MsgBox "季度汇总颜色已清理并重新生成。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "清理季度汇总颜色失败：" & Err.Description, vbCritical, "错误"
End Sub

Sub CheckManualConfirmationSetup()
    Dim msg As String
    Dim wsInvoice As Worksheet
    Dim wsSummary As Worksheet
    Dim aiHeader As String

    On Error Resume Next
    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")
    On Error GoTo 0

    msg = "人工确认联动检查：" & vbCrLf & vbCrLf

    If wsInvoice Is Nothing Then
        msg = msg & "× 找不到【发票明细】工作表" & vbCrLf
    Else
        aiHeader = Trim(CStr(wsInvoice.Cells(2, 35).Value))
        msg = msg & "√ 找到【发票明细】工作表" & vbCrLf
        msg = msg & "AI列标题：" & IIf(aiHeader = "", "空", aiHeader) & vbCrLf
    End If

    If wsSummary Is Nothing Then
        msg = msg & "× 找不到【季度汇总】工作表" & vbCrLf
    Else
        msg = msg & "√ 找到【季度汇总】工作表" & vbCrLf
    End If

    msg = msg & vbCrLf & "日常使用：运行 RunMatch。" & vbCrLf & _
          "维护使用：运行 SetupWorkbookOnce 或 OptimizeWorkbookLayout。"

    MsgBox msg, vbInformation, "检查结果"
End Sub

Sub CheckMappingConsistency()
    Dim wsCatalog As Worksheet
    Dim wsMapping As Worksheet
    Dim dict As Object
    Dim r As Long
    Dim lastRow As Long
    Dim taxName As String
    Dim badRules As String
    Dim badCount As Long

    On Error GoTo ErrorHandler

    Set wsCatalog = ThisWorkbook.Worksheets("税目维护")
    Set wsMapping = ThisWorkbook.Worksheets("税目映射规则")
    Set dict = CreateObject("Scripting.Dictionary")

    For r = 2 To wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row
        taxName = Trim(CStr(wsCatalog.Cells(r, 1).Value))
        If Len(taxName) > 0 And Trim(CStr(wsCatalog.Cells(r, 3).Value)) <> "否" Then
            dict(taxName) = True
        End If
    Next r
    dict("不征收") = True
    dict("不征税") = True

    lastRow = wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        taxName = Trim(CStr(wsMapping.Cells(r, 6).Value))
        If Len(taxName) > 0 Then
            If Not dict.Exists(taxName) Then
                badCount = badCount + 1
                If Len(badRules) < 800 Then
                    badRules = badRules & "第" & r & "行：" & taxName & vbCrLf
                End If
            End If
        End If
    Next r

    If badCount = 0 Then
        MsgBox "检查通过：税目映射规则 F列均能在【税目维护】或“不征收”口径中找到。", vbInformation, "检查结果"
    Else
        MsgBox "发现 " & badCount & " 条未映射口径，请补充到【税目维护】或改为“不征收”：" & vbCrLf & vbCrLf & badRules, vbExclamation, "检查结果"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "检查失败：" & Err.Description, vbCritical, "错误"
End Sub

Sub RepairTaxCatalog()
    Dim wsCatalog As Worksheet
    Dim wsMapping As Worksheet
    Dim wsSummary As Worksheet

    On Error GoTo ErrorHandler

    Set wsCatalog = ThisWorkbook.Worksheets("税目维护")
    Set wsMapping = ThisWorkbook.Worksheets("税目映射规则")
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Call NormalizeMappingNonTaxableLabels(wsMapping)
    Call CleanTaxCatalogSheet(wsCatalog)
    Call EnsureTaxCatalogSheet(wsMapping, wsSummary)
    Call SetMappingRuleValidation(wsMapping)
    Call RefreshQuarterSummary

CleanUp:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "税目维护已修复：已删除非税目展示项，并重建下拉。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "修复税目维护失败：" & Err.Description, vbCritical, "错误"
End Sub

Sub ExplainColorRules()
    MsgBox "发票明细颜色说明：" & vbCrLf & vbCrLf & _
           "红色：匹配税目=未匹配、匹配结果不是不征收、排除原因=无匹配规则，且【人工确认】为空，需要补规则或人工确认。" & vbCrLf & _
           "黄色：【人工确认】为空，且匹配税目/结果/原因提示争议、待确认、需确认或待补税目，需要复核。" & vbCrLf & _
           "绿色：【人工确认】已选择不征收或具体税目，且匹配结果有效，表示已人工处理。" & vbCrLf & _
           "橙色：【人工确认】填了非标准内容，只标该单元格，请改用下拉值。" & vbCrLf & vbCrLf & _
           "列位置按表头识别，不再依赖肉眼判断 AI/AB 等列号。" & vbCrLf & _
           "如出现旧颜色残留，请运行 OptimizeWorkbookLayout。", vbInformation, "颜色说明"
End Sub

Sub DiagnoseSelectedInvoiceColor()
    Dim wsInvoice As Worksheet
    Dim rowNum As Long
    Dim matchTaxCol As Long
    Dim matchResultCol As Long
    Dim excludeReasonCol As Long
    Dim manualCol As Long
    Dim matchTax As String
    Dim matchResult As String
    Dim excludeReason As String
    Dim manualConfirm As String
    Dim diagnosis As String

    On Error GoTo ErrorHandler

    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
    If ActiveSheet.Name <> wsInvoice.Name Then
        MsgBox "请先在【发票明细】中选中要诊断的行。", vbExclamation, "颜色诊断"
        Exit Sub
    End If

    rowNum = ActiveCell.Row
    If rowNum < 3 Then
        MsgBox "请选中第3行之后的发票数据行。", vbExclamation, "颜色诊断"
        Exit Sub
    End If

    matchTaxCol = FindInvoiceColumn(wsInvoice, "匹配税目", 28)
    matchResultCol = FindInvoiceColumn(wsInvoice, "匹配结果", 31)
    excludeReasonCol = FindInvoiceColumn(wsInvoice, "排除原因", 34)
    manualCol = FindInvoiceColumn(wsInvoice, "人工确认", 35)

    matchTax = Trim(CStr(wsInvoice.Cells(rowNum, matchTaxCol).Value))
    matchResult = Trim(CStr(wsInvoice.Cells(rowNum, matchResultCol).Value))
    excludeReason = Trim(CStr(wsInvoice.Cells(rowNum, excludeReasonCol).Value))
    manualConfirm = Trim(CStr(wsInvoice.Cells(rowNum, manualCol).Value))

    diagnosis = "当前行：" & rowNum & vbCrLf & vbCrLf & _
                "匹配税目（" & ColumnLetter(wsInvoice, matchTaxCol) & "列）：" & IIf(matchTax = "", "空", matchTax) & vbCrLf & _
                "匹配结果（" & ColumnLetter(wsInvoice, matchResultCol) & "列）：" & IIf(matchResult = "", "空", matchResult) & vbCrLf & _
                "排除原因（" & ColumnLetter(wsInvoice, excludeReasonCol) & "列）：" & IIf(excludeReason = "", "空", excludeReason) & vbCrLf & _
                "人工确认（" & ColumnLetter(wsInvoice, manualCol) & "列）：" & IIf(manualConfirm = "", "空", manualConfirm) & vbCrLf & vbCrLf

    If matchTax = "未匹配" And matchResult <> "不征收" And excludeReason = "无匹配规则" And manualConfirm = "" Then
        diagnosis = diagnosis & "应显示红色：未匹配、无匹配规则且人工确认为空。"
    ElseIf manualConfirm = "" And excludeReason <> "无匹配规则" And _
           (InStr(1, matchTax, "争议", vbTextCompare) > 0 Or _
            InStr(1, matchTax, "待确认", vbTextCompare) > 0 Or _
            InStr(1, matchTax, "需确认", vbTextCompare) > 0 Or _
            InStr(1, matchResult, "待确认", vbTextCompare) > 0 Or _
            InStr(1, excludeReason, "待确认", vbTextCompare) > 0 Or _
            matchResult = "待补税目") Then
        diagnosis = diagnosis & "应显示黄色：需要人工复核。"
    ElseIf manualConfirm <> "" And (IsNonTaxableText(manualConfirm) Or IsTaxableText(manualConfirm) Or IsKnownTaxCatalogItem(manualConfirm) Or matchResult = "应税" Or matchResult = "不征收") Then
        diagnosis = diagnosis & "应显示绿色：人工确认已生效。"
    ElseIf manualConfirm <> "" Then
        diagnosis = diagnosis & "人工确认有值但未形成有效结果，应检查是否用了下拉标准值。"
    Else
        diagnosis = diagnosis & "不应高亮；如仍有底色，请运行 OptimizeWorkbookLayout 清理旧颜色。"
    End If

    MsgBox diagnosis, vbInformation, "颜色诊断"
    Exit Sub

ErrorHandler:
    MsgBox "颜色诊断失败：" & Err.Description, vbCritical, "颜色诊断"
End Sub
