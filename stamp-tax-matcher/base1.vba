'========================================
' 印花税自动匹配模型 - 统一版 VBA
'========================================
' 使用方式：
' 1. 整个文件导入/粘贴到一个标准模块，例如 Module1。
' 2. 正常只需要运行 RunMatch。
' 3. RunMatch 会给【发票明细】写入联动公式和“人工确认”下拉框。
' 4. AI列可选择“不征收”或【税目维护】中的具体税目；选择具体税目后会进入上方对应税目汇总。
' 5. 若要【税目维护】改完后汇总布局自动增减行，请在 ThisWorkbook 中放置：
'    Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'        HandleWorkbookSheetChange Sh, Target
'    End Sub
'========================================

Const TAX_CATALOG_MAX_ROWS As Long = 300

Sub RunMatch()
    '========================================
    ' 印花税自动匹配 - V4优化版
    ' 优化：解决卡死问题，提升性能
    '========================================

    Dim wsRaw As Worksheet
    Dim wsInvoice As Worksheet
    Dim wsMapping As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRowRaw As Long
    Dim lastRowMapping As Long
    Dim i As Long
    Dim outputRow As Long
    Dim startTime As Double
    Dim oldCalculation As XlCalculation

    On Error GoTo ErrorHandler
    startTime = Timer
    oldCalculation = Application.Calculation

    ' 禁用屏幕更新和自动计算（性能优化）
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False  '禁用事件

    ' 显示进度提示
    Application.StatusBar = "正在初始化..."

    ' 1. 检查工作表
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Worksheets("原始数据")
    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
    Set wsMapping = ThisWorkbook.Worksheets("税目映射规则")
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")
    On Error GoTo ErrorHandler

    If wsRaw Is Nothing Or wsInvoice Is Nothing Or wsMapping Is Nothing Then
        MsgBox "找不到必需的工作表！", vbCritical, "错误"
        GoTo CleanUp
    End If

    ' 日常匹配只做必要检查；完整下拉/模板维护由 SetupWorkbookOnce 或 OptimizeWorkbookLayout 执行。
    If Not WorksheetExists("税目维护") Or Not WorkbookNameExists("ManualConfirmList") Or Not WorkbookNameExists("TaxCatalogItems") Then
        Call EnsureTaxCatalogSheet(wsMapping, wsSummary)
    End If

    ' 1.5 检查原始数据格式
    Call CheckTaxCodeFormat(wsRaw)

    ' 2. 清空发票明细旧数据（优化：只清除使用的区域）
    Application.StatusBar = "正在清空旧数据..."
    Dim lastRowInvoice As Long
    lastRowInvoice = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row
    If lastRowInvoice >= 3 Then
        ' 只清除实际使用的列（1-42列），避免清除整行
        wsInvoice.Range(wsInvoice.Cells(3, 1), wsInvoice.Cells(lastRowInvoice, 42)).Clear
    End If

    ' 3. 清空季度汇总表
    If Not wsSummary Is Nothing Then
        Dim lastRowSummary As Long
        lastRowSummary = wsSummary.Cells(wsSummary.Rows.Count, 1).End(xlUp).Row
        If lastRowSummary >= 21 Then
            wsSummary.Range(wsSummary.Cells(21, 1), wsSummary.Cells(lastRowSummary, 10)).Clear
        End If
    End If

    ' 4. 获取数据行数
    lastRowRaw = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    lastRowMapping = wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row

    If lastRowRaw < 3 Then
        MsgBox "【原始数据】中没有数据！", vbExclamation, "提示"
        GoTo CleanUp
    End If

    ' 5. 加载映射规则到数组（性能优化：一次性读取）
    Application.StatusBar = "正在加载匹配规则..."
    Dim rules() As Variant
    Dim ruleCount As Long
    ruleCount = lastRowMapping - 1

    If ruleCount > 0 Then
        ReDim rules(1 To ruleCount, 1 To 11)

        Dim j As Long
        For j = 2 To lastRowMapping
            rules(j - 1, 1) = wsMapping.Cells(j, 1).Value   '规则编号
            rules(j - 1, 2) = Trim(CStr(wsMapping.Cells(j, 2).Value))  '编码前缀
            rules(j - 1, 3) = wsMapping.Cells(j, 3).Value   '位置
            rules(j - 1, 4) = Trim(CStr(wsMapping.Cells(j, 4).Value))  '名称包含
            rules(j - 1, 5) = Trim(CStr(wsMapping.Cells(j, 5).Value))  '关键词判断
            rules(j - 1, 6) = wsMapping.Cells(j, 6).Value   '对应税目
            rules(j - 1, 7) = wsMapping.Cells(j, 7).Value   '税率
            rules(j - 1, 8) = wsMapping.Cells(j, 8).Value   '是否征税
            rules(j - 1, 9) = wsMapping.Cells(j, 9).Value   '优先级
            rules(j - 1, 10) = Trim(CStr(wsMapping.Cells(j, 10).Value)) '排除关键词
            rules(j - 1, 11) = wsMapping.Cells(j, 11).Value  '备注
        Next j

        ' 按优先级排序
        Call SortRulesByPriority(rules, ruleCount)
    End If

    ' 6. 批量读取原始数据到数组（性能优化：减少单元格访问）
    Application.StatusBar = "正在读取发票数据..."
    Dim rawData() As Variant
    rawData = wsRaw.Range(wsRaw.Cells(3, 1), wsRaw.Cells(lastRowRaw, 27)).Value

    ' 6.5 单独读取长数字列的Text属性（避免科学计数法）
    Dim textColumns(1 To 6) As Long
    textColumns(1) = 2: textColumns(2) = 3: textColumns(3) = 4
    textColumns(4) = 5: textColumns(5) = 7: textColumns(6) = 10

    Dim textData() As Variant
    ReDim textData(1 To UBound(rawData, 1), 1 To 6)

    Dim rowIdx As Long, colIdx As Long
    For rowIdx = 1 To UBound(rawData, 1)
        For colIdx = 1 To 6
            ' 使用.Text属性读取，保留原始文本
            textData(rowIdx, colIdx) = wsRaw.Cells(rowIdx + 2, textColumns(colIdx)).text
        Next colIdx
    Next rowIdx

    ' 7. 准备输出数组（性能优化：批量写入）
    Dim outputData() As Variant
    Dim validRows As Long
    validRows = 0

    ' 先计算有效行数
    For i = 1 To UBound(rawData, 1)
        If Not IsEmpty(rawData(i, 1)) And Trim(CStr(rawData(i, 1))) <> "" Then
            validRows = validRows + 1
        End If
    Next i

    If validRows = 0 Then
        MsgBox "【原始数据】中没有有效数据！", vbExclamation, "提示"
        GoTo CleanUp
    End If

    ReDim outputData(1 To validRows, 1 To 42)

    ' 8. 遍历数据进行匹配
    outputRow = 0
    Dim progressInterval As Long
    progressInterval = Application.Max(1, Int(UBound(rawData, 1) / 20))  '每5%更新一次进度

    For i = 1 To UBound(rawData, 1)
        ' 更新进度（每处理5%显示一次）
        If i Mod progressInterval = 0 Then
            Application.StatusBar = "正在匹配发票... " & Format(i / UBound(rawData, 1), "0%")
        End If

        ' 跳过空行
        If IsEmpty(rawData(i, 1)) Or Trim(CStr(rawData(i, 1))) = "" Then
            GoTo NextRow
        End If

        outputRow = outputRow + 1

        ' 复制前27列基础数据
        Dim k As Long
        For k = 1 To 27
            ' 对于长数字列，使用Text属性读取的数据
            If k = 2 Then
                outputData(outputRow, k) = "'" & textData(i, 1)  '发票代码
            ElseIf k = 3 Then
                outputData(outputRow, k) = "'" & textData(i, 2)  '发票号码
            ElseIf k = 4 Then
                outputData(outputRow, k) = "'" & textData(i, 3)  '数电发票号码
            ElseIf k = 5 Then
                outputData(outputRow, k) = "'" & textData(i, 4)  '销方识别号
            ElseIf k = 7 Then
                outputData(outputRow, k) = "'" & textData(i, 5)  '购方识别号
            ElseIf k = 10 Then
                outputData(outputRow, k) = "'" & textData(i, 6)  '税收分类编码
            Else
                outputData(outputRow, k) = rawData(i, k)
            End If
        Next k

        ' 获取关键字段（使用Text数据）
        Dim taxCode As String
        Dim goodsName As String
        Dim amount As Double
        Dim sellerTaxId As String

        taxCode = Trim(textData(i, 6))  '税收分类编码（使用Text）
        goodsName = Trim(CStr(rawData(i, 12)))
        sellerTaxId = Trim(textData(i, 4))  '销方税号（第5列，textData第4个）

        amount = 0
        If IsNumeric(rawData(i, 20)) Then
            amount = CDbl(rawData(i, 20))
        End If

        ' 执行匹配
        Dim matchResult As Variant
        matchResult = MatchInvoice(taxCode, goodsName, amount, sellerTaxId, rules, ruleCount)

        ' 写入匹配结果（列28-35）
        outputData(outputRow, 28) = matchResult(0)  '匹配税目
        outputData(outputRow, 29) = matchResult(1)  '税率
        outputData(outputRow, 30) = matchResult(2)  '应纳税额
        outputData(outputRow, 31) = matchResult(3)  '匹配结果
        outputData(outputRow, 32) = matchResult(4)  '匹配规则
        outputData(outputRow, 33) = matchResult(5)  '是否排除
        outputData(outputRow, 34) = matchResult(6)  '排除原因
        outputData(outputRow, 35) = ""              '人工确认
        outputData(outputRow, 36) = matchResult(5)  '自动判断-是否排除（隐藏辅助）
        outputData(outputRow, 37) = matchResult(6)  '自动判断-排除原因（隐藏辅助）
        outputData(outputRow, 38) = matchResult(2)  '自动判断-应纳税额（隐藏辅助）
        outputData(outputRow, 39) = matchResult(3)  '自动判断-匹配结果（隐藏辅助）
        outputData(outputRow, 41) = matchResult(0)  '自动判断-匹配税目（隐藏辅助）
        outputData(outputRow, 42) = matchResult(1)  '自动判断-税率（隐藏辅助）

NextRow:
    Next i

    ' 9. 批量写入结果（性能优化：一次性写入）
    Application.StatusBar = "正在写入结果..."
    If outputRow > 0 Then
        wsInvoice.Range(wsInvoice.Cells(3, 1), wsInvoice.Cells(2 + outputRow, 42)).Value = outputData

        ' 设置文本格式列（所有可能是长数字的列）
        ' 列2: 发票代码
        ' 列3: 发票号码
        ' 列4: 数电发票号码
        ' 列5: 销方识别号
        ' 列7: 购方识别号
        ' 列10: 税收分类编码
        wsInvoice.Range(wsInvoice.Cells(3, 2), wsInvoice.Cells(2 + outputRow, 2)).NumberFormat = "@"
        wsInvoice.Range(wsInvoice.Cells(3, 3), wsInvoice.Cells(2 + outputRow, 3)).NumberFormat = "@"
        wsInvoice.Range(wsInvoice.Cells(3, 4), wsInvoice.Cells(2 + outputRow, 4)).NumberFormat = "@"
        wsInvoice.Range(wsInvoice.Cells(3, 5), wsInvoice.Cells(2 + outputRow, 5)).NumberFormat = "@"
        wsInvoice.Range(wsInvoice.Cells(3, 7), wsInvoice.Cells(2 + outputRow, 7)).NumberFormat = "@"
        wsInvoice.Range(wsInvoice.Cells(3, 10), wsInvoice.Cells(2 + outputRow, 10)).NumberFormat = "@"

        Call SetManualConfirmValidation(wsInvoice, outputRow)

        Call InstallManualConfirmationFormulas(wsInvoice, outputRow)

        ' 使用条件格式替代逐行染色，避免大数据量时卡顿。
        Call SetInvoiceConditionalFormatting(wsInvoice, outputRow)
    End If

    ' 10. 生成季度汇总
    Application.StatusBar = "正在生成汇总..."
    Call GenerateSummaryConfirmation(wsInvoice, wsSummary)

    ' 11. 完成提示
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    MsgBox "匹配完成！" & vbCrLf & vbCrLf & _
           "共处理 " & outputRow & " 条发票数据" & vbCrLf & _
           "耗时 " & Format(elapsedTime, "0.0") & " 秒", vbInformation, "完成"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.Calculate
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "发生错误：" & Err.Description & vbCrLf & _
           "错误代码：" & Err.Number, vbCritical, "错误"
End Sub

Function EnsureTaxCatalogSheet(wsMapping As Worksheet, wsSummary As Worksheet) As Worksheet
    Dim wsCatalog As Worksheet
    Dim lastRow As Long
    Dim hasData As Boolean
    Dim dict As Object
    Dim r As Long
    Dim taxName As String
    Dim taxRate As Variant

    On Error Resume Next
    Set wsCatalog = ThisWorkbook.Worksheets("税目维护")
    On Error GoTo 0

    If wsCatalog Is Nothing Then
        Set wsCatalog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsCatalog.Name = "税目维护"
    End If

    wsCatalog.Range("A1:D1").Value = Array("税目", "税率(‰)", "启用", "备注")
    wsCatalog.Range("E1").Value = "人工确认下拉"
    wsCatalog.Range("F1").Value = "税目下拉"
    wsCatalog.Rows(1).Font.Bold = True

    lastRow = wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row
    hasData = (lastRow >= 2 And Len(Trim(CStr(wsCatalog.Cells(2, 1).Value))) > 0)

    Set dict = CreateObject("Scripting.Dictionary")
    If hasData Then
        For r = 2 To lastRow
            taxName = Trim(CStr(wsCatalog.Cells(r, 1).Value))
            If Len(taxName) > 0 Then dict(taxName) = True
            If Len(Trim(CStr(wsCatalog.Cells(r, 3).Value))) = 0 Then wsCatalog.Cells(r, 3).Value = "是"
        Next r
    End If

    ' 首次使用或规则新增税目时，从【税目映射规则】F/G列补充到维护表。
    If Not wsMapping Is Nothing Then
        For r = 2 To wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row
            taxName = Trim(CStr(wsMapping.Cells(r, 6).Value))
            taxRate = wsMapping.Cells(r, 7).Value
            If ShouldAddTaxCatalogItem(taxName, taxRate) Then
                If Not dict.Exists(taxName) Then
                    lastRow = wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row + 1
                    wsCatalog.Cells(lastRow, 1).Value = taxName
                    wsCatalog.Cells(lastRow, 2).Value = taxRate
                    wsCatalog.Cells(lastRow, 3).Value = "是"
                    wsCatalog.Cells(lastRow, 4).Value = "从税目映射规则自动补充"
                    dict(taxName) = True
                End If
            End If
        Next r
    End If

    ' 兼容老模板：如果映射规则里没有税目，则从【季度汇总】A/B列种子数据补充。
    If Not wsSummary Is Nothing Then
        For r = 5 To 50
            taxName = Trim(CStr(wsSummary.Cells(r, 1).Value))
            taxRate = wsSummary.Cells(r, 2).Value
            If InStr(1, taxName, "合计", vbTextCompare) > 0 Or _
               InStr(1, taxName, "人工确认", vbTextCompare) > 0 Or _
               InStr(1, taxName, "待确认", vbTextCompare) > 0 Then
                Exit For
            End If
            If ShouldAddTaxCatalogItem(taxName, taxRate) Then
                If Not dict.Exists(taxName) Then
                    lastRow = wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row + 1
                    wsCatalog.Cells(lastRow, 1).Value = taxName
                    wsCatalog.Cells(lastRow, 2).Value = taxRate
                    wsCatalog.Cells(lastRow, 3).Value = "是"
                    wsCatalog.Cells(lastRow, 4).Value = "从季度汇总模板自动补充"
                    dict(taxName) = True
                End If
            End If
        Next r
    End If

    Call CleanTaxCatalogSheet(wsCatalog)
    Call SetTaxCatalogValidation(wsCatalog)
    Call RefreshTaxCatalogNames(wsCatalog)
    Set EnsureTaxCatalogSheet = wsCatalog
End Function

Sub CleanTaxCatalogSheet(wsCatalog As Worksheet)
    Dim r As Long
    Dim lastRow As Long
    Dim taxName As String

    lastRow = wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row
    For r = lastRow To 2 Step -1
        taxName = Trim(CStr(wsCatalog.Cells(r, 1).Value))
        If Len(taxName) > 0 Then
            If IsInvalidTaxCatalogName(taxName) Then
                wsCatalog.Rows(r).Delete
            End If
        End If
    Next r
End Sub

Sub SetTaxCatalogValidation(wsCatalog As Worksheet)
    With wsCatalog.Range("B2:B" & TAX_CATALOG_MAX_ROWS + 1).Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="100"
        .IgnoreBlank = True
        .InputTitle = "税率(‰)"
        .InputMessage = "请输入千分比税率，例如0.3"
    End With

    With wsCatalog.Range("C2:C" & TAX_CATALOG_MAX_ROWS + 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="是" & Application.International(xlListSeparator) & "否"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "是否启用"
        .InputMessage = "请选择：是 或 否"
    End With
End Sub

Function ShouldAddTaxCatalogItem(taxName As String, taxRate As Variant) As Boolean
    taxName = Trim(CStr(taxName))
    If Len(taxName) = 0 Then Exit Function
    If IsInvalidTaxCatalogName(taxName) Then Exit Function
    If Not IsNumeric(taxRate) Then Exit Function

    ShouldAddTaxCatalogItem = True
End Function

Function IsInvalidTaxCatalogName(taxName As String) As Boolean
    taxName = Trim(CStr(taxName))
    If Len(taxName) = 0 Then Exit Function
    If taxName = "未匹配" Then IsInvalidTaxCatalogName = True: Exit Function
    If taxName = "税目" Then IsInvalidTaxCatalogName = True: Exit Function
    If taxName = "不征税项目" Then IsInvalidTaxCatalogName = True: Exit Function
    If IsNonTaxableText(taxName) Then IsInvalidTaxCatalogName = True: Exit Function
    If InStr(1, taxName, "合计", vbTextCompare) > 0 Then IsInvalidTaxCatalogName = True: Exit Function
    If InStr(1, taxName, "人工确认", vbTextCompare) > 0 Then IsInvalidTaxCatalogName = True: Exit Function
    If InStr(1, taxName, "待确认", vbTextCompare) > 0 Then IsInvalidTaxCatalogName = True: Exit Function
    If InStr(1, taxName, "黄色高亮", vbTextCompare) > 0 Then IsInvalidTaxCatalogName = True: Exit Function
    If InStr(1, taxName, "处理建议", vbTextCompare) > 0 Then IsInvalidTaxCatalogName = True: Exit Function
    If Left(taxName, 1) = "⚠" Or Left(taxName, 1) = "?" Then IsInvalidTaxCatalogName = True
End Function

Sub RefreshTaxCatalogNames(wsCatalog As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim taxName As String
    Dim enabledFlag As String

    lastRow = wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    If lastRow > TAX_CATALOG_MAX_ROWS + 1 Then lastRow = TAX_CATALOG_MAX_ROWS + 1

    wsCatalog.Range("E2:F" & TAX_CATALOG_MAX_ROWS + 2).ClearContents
    wsCatalog.Range("E2").Value = "不征收"

    For r = 2 To TAX_CATALOG_MAX_ROWS + 1
        wsCatalog.Cells(r + 1, 5).FormulaR1C1 = "=IF(AND(R[-1]C[-4]<>"""",R[-1]C[-2]<>""否""),R[-1]C[-4],"""")"
        wsCatalog.Cells(r, 6).FormulaR1C1 = "=IF(AND(RC[-5]<>"""",RC[-3]<>""否""),RC[-5],"""")"
    Next r

    wsCatalog.Columns("A:F").AutoFit

    Call SetWorkbookName("ManualConfirmList", "='" & wsCatalog.Name & "'!$E$2:INDEX('" & wsCatalog.Name & "'!$E:$E,LOOKUP(2,1/('" & wsCatalog.Name & "'!$E$2:$E$" & TAX_CATALOG_MAX_ROWS + 2 & "<>""""),ROW('" & wsCatalog.Name & "'!$E$2:$E$" & TAX_CATALOG_MAX_ROWS + 2 & ")))")
    Call SetWorkbookName("TaxConfirmList", "='" & wsCatalog.Name & "'!$F$2:INDEX('" & wsCatalog.Name & "'!$F:$F,LOOKUP(2,1/('" & wsCatalog.Name & "'!$F$2:$F$" & TAX_CATALOG_MAX_ROWS + 1 & "<>""""),ROW('" & wsCatalog.Name & "'!$F$2:$F$" & TAX_CATALOG_MAX_ROWS + 1 & ")))")
    Call SetWorkbookName("RuleTaxCategoryList", "='" & wsCatalog.Name & "'!$E$2:INDEX('" & wsCatalog.Name & "'!$E:$E,LOOKUP(2,1/('" & wsCatalog.Name & "'!$E$2:$E$" & TAX_CATALOG_MAX_ROWS + 1 & "<>""""),ROW('" & wsCatalog.Name & "'!$E$2:$E$" & TAX_CATALOG_MAX_ROWS + 1 & ")))")
    Call SetWorkbookName("TaxCatalogItems", "='" & wsCatalog.Name & "'!$A$2:$A$" & TAX_CATALOG_MAX_ROWS + 1)
    Call SetWorkbookName("TaxCatalogRates", "='" & wsCatalog.Name & "'!$B$2:$B$" & TAX_CATALOG_MAX_ROWS + 1)
End Sub

Sub SetWorkbookName(nameText As String, refersToText As String)
    On Error Resume Next
    ThisWorkbook.Names(nameText).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nameText, RefersTo:=refersToText
End Sub

Function WorkbookNameExists(nameText As String) As Boolean
    Dim nm As Name
    On Error Resume Next
    Set nm = ThisWorkbook.Names(nameText)
    WorkbookNameExists = Not nm Is Nothing
    On Error GoTo 0
End Function

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Function SetupQuarterSummaryTemplate(wsSummary As Worksheet, wsCatalog As Worksheet, Optional lastInvoiceRow As Long = 3) As Long
    Dim r As Long
    Dim outRow As Long
    Dim totalRow As Long
    Dim titleRow As Long
    Dim headerRow As Long
    Dim firstTaxRow As Long
    Dim lastTaxRow As Long
    Dim lastCatalogRow As Long
    Dim taxName As String
    Dim enabledFlag As String
    Dim invoiceEndRow As Long
    Dim amountRange As String
    Dim taxCategoryRange As String
    Dim excludeRange As String
    Dim invoiceKeyRange As String
    Dim pendingTypeRange As String
    Dim manualConfirmRange As String

    invoiceEndRow = Application.Max(3, lastInvoiceRow)
    amountRange = "'发票明细'!$T$3:$T$" & invoiceEndRow
    taxCategoryRange = "'发票明细'!$AB$3:$AB$" & invoiceEndRow
    excludeRange = "'发票明细'!$AG$3:$AG$" & invoiceEndRow
    invoiceKeyRange = "'发票明细'!$A$3:$A$" & invoiceEndRow
    pendingTypeRange = "'发票明细'!$AN$3:$AN$" & invoiceEndRow
    manualConfirmRange = "'发票明细'!$AI$3:$AI$" & invoiceEndRow

    ' 清理旧版本整行染色遗留，避免汇总页右侧空白区域残留粉/绿色。
    wsSummary.Rows("5:300").FormatConditions.Delete
    wsSummary.Rows("5:300").Interior.Pattern = xlNone
    wsSummary.Range("A5:J300").Clear

    wsSummary.Range("A4:E4").Value = Array("税目", "税率(‰)", "计税依据(价税合计)", "应纳税额", "发票数量")
    wsSummary.Range("A4:E4").Interior.Color = RGB(68, 114, 196)
    wsSummary.Range("A4:E4").Font.Color = RGB(255, 255, 255)
    wsSummary.Range("A4:E4").Font.Bold = True
    wsSummary.Range("A4:E4").HorizontalAlignment = xlCenter

    firstTaxRow = 5
    outRow = firstTaxRow
    lastCatalogRow = wsCatalog.Cells(wsCatalog.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastCatalogRow
        taxName = Trim(CStr(wsCatalog.Cells(r, 1).Value))
        enabledFlag = Trim(CStr(wsCatalog.Cells(r, 3).Value))
        If ShouldAddTaxCatalogItem(taxName, wsCatalog.Cells(r, 2).Value) And enabledFlag <> "否" Then
            wsSummary.Cells(outRow, 1).Formula = "='税目维护'!A" & r
            wsSummary.Cells(outRow, 2).Formula = "='税目维护'!B" & r
            wsSummary.Cells(outRow, 3).Formula = "=SUMIFS(" & amountRange & "," & taxCategoryRange & ",A" & outRow & "," & excludeRange & ",""否"")"
            wsSummary.Cells(outRow, 4).Formula = "=C" & outRow & "*B" & outRow & "/1000"
            wsSummary.Cells(outRow, 5).Formula = "=COUNTIFS(" & taxCategoryRange & ",A" & outRow & "," & excludeRange & ",""否"")"
            outRow = outRow + 1
        End If
    Next r

    If outRow = firstTaxRow Then
        wsSummary.Cells(outRow, 1).Value = "请在【税目维护】新增税目"
        outRow = outRow + 1
    End If

    lastTaxRow = outRow - 1

    totalRow = lastTaxRow + 3
    wsSummary.Cells(totalRow, 1).Value = "合计"
    wsSummary.Cells(totalRow, 4).Formula = "=SUM(D" & firstTaxRow & ":D" & lastTaxRow & ")"
    wsSummary.Cells(totalRow, 5).Formula = "=SUM(E" & firstTaxRow & ":E" & lastTaxRow & ")"
    wsSummary.Rows(totalRow).Font.Bold = True

    wsSummary.Cells(totalRow + 1, 1).Value = "口径核对"
    wsSummary.Cells(totalRow + 1, 1).Font.Bold = True
    wsSummary.Cells(totalRow + 2, 1).Value = "发票总数"
    wsSummary.Cells(totalRow + 2, 2).Formula = "=COUNTA(" & invoiceKeyRange & ")"
    wsSummary.Cells(totalRow + 2, 3).Value = "应税入表"
    wsSummary.Cells(totalRow + 2, 4).Formula = "=COUNTIFS(" & excludeRange & ",""否"")"
    wsSummary.Cells(totalRow + 3, 1).Value = "不征收/排除完成"
    wsSummary.Cells(totalRow + 3, 2).Formula = "=COUNTIFS(" & excludeRange & ",""是"")-D" & totalRow + 3
    wsSummary.Cells(totalRow + 3, 3).Value = "待处理"
    wsSummary.Cells(totalRow + 3, 4).Formula = "=COUNTIFS(" & pendingTypeRange & ",""*""," & manualConfirmRange & ","""")" & _
                                                "+COUNTIFS(" & pendingTypeRange & ",""未匹配""," & manualConfirmRange & ",""应税"")" & _
                                                "+COUNTIFS(" & pendingTypeRange & ",""未匹配""," & manualConfirmRange & ",""征税"")" & _
                                                "+COUNTIFS(" & pendingTypeRange & ",""未匹配""," & manualConfirmRange & ",""征收"")"
    wsSummary.Cells(totalRow + 4, 1).Value = "核对差异"
    wsSummary.Cells(totalRow + 4, 2).Formula = "=B" & totalRow + 2 & "-D" & totalRow + 2 & "-B" & totalRow + 3 & "-D" & totalRow + 3

    titleRow = totalRow + 6
    headerRow = totalRow + 7
    wsSummary.Cells(titleRow, 1).Value = "⚠  需人工确认项目（黄色高亮行）"
    wsSummary.Range(wsSummary.Cells(titleRow, 1), wsSummary.Cells(titleRow, 5)).Font.Color = RGB(255, 0, 0)
    wsSummary.Range(wsSummary.Cells(titleRow, 1), wsSummary.Cells(titleRow, 5)).Font.Bold = True

    wsSummary.Range(wsSummary.Cells(headerRow, 1), wsSummary.Cells(headerRow, 5)).Value = Array("待确认类型", "发票数量", "价税合计", "预估税额", "处理建议")
    wsSummary.Range(wsSummary.Cells(headerRow, 1), wsSummary.Cells(headerRow, 5)).Interior.Color = RGB(68, 114, 196)
    wsSummary.Range(wsSummary.Cells(headerRow, 1), wsSummary.Cells(headerRow, 5)).Font.Color = RGB(255, 255, 255)
    wsSummary.Range(wsSummary.Cells(headerRow, 1), wsSummary.Cells(headerRow, 5)).Font.Bold = True
    wsSummary.Range(wsSummary.Cells(headerRow, 1), wsSummary.Cells(headerRow, 5)).HorizontalAlignment = xlCenter

    wsSummary.Range("A4:E" & headerRow).Borders.LineStyle = xlContinuous
    wsSummary.Columns("A:E").AutoFit

    SetupQuarterSummaryTemplate = headerRow + 1
End Function

'========================================
' 人工确认联动公式：不依赖事件，导入标准模块即可使用
'========================================
Sub InstallManualConfirmationFormulas(wsInvoice As Worksheet, totalRows As Long)
    If totalRows <= 0 Then Exit Sub

    Dim firstDataRow As Long
    Dim lastDataRow As Long

    firstDataRow = 3
    lastDataRow = 2 + totalRows

    wsInvoice.Cells(2, 36).Value = "自动是否排除"
    wsInvoice.Cells(2, 37).Value = "自动排除原因"
    wsInvoice.Cells(2, 38).Value = "自动应纳税额"
    wsInvoice.Cells(2, 39).Value = "自动匹配结果"
    wsInvoice.Cells(2, 40).Value = "待确认类型"
    wsInvoice.Cells(2, 41).Value = "自动匹配税目"
    wsInvoice.Cells(2, 42).Value = "自动税率"

    ' AB 匹配税目：AI选择具体税目时，人工税目优先进入上方汇总。
    wsInvoice.Range(wsInvoice.Cells(firstDataRow, 28), wsInvoice.Cells(lastDataRow, 28)).FormulaR1C1 = _
        "=IF(OR(RC[7]=""不征收"",RC[7]=""不征税"",RC[7]=""免征""),RC[13],IF(ISNUMBER(MATCH(RC[7],TaxCatalogItems,0)),RC[7],RC[13]))"

    ' AC 税率：人工选择具体税目后，从【税目维护】取税率。
    wsInvoice.Range(wsInvoice.Cells(firstDataRow, 29), wsInvoice.Cells(lastDataRow, 29)).FormulaR1C1 = _
        "=IF(ISNUMBER(MATCH(RC[6],TaxCatalogItems,0)),INDEX(TaxCatalogRates,MATCH(RC[6],TaxCatalogItems,0)),RC[13])"

    ' AD 应纳税额：AI填不征收则为0；AI选具体税目则按对应税率计税；只填"应税"但仍未匹配时，必须先补税目。
    wsInvoice.Range(wsInvoice.Cells(firstDataRow, 30), wsInvoice.Cells(lastDataRow, 30)).FormulaR1C1 = _
        "=IF(OR(RC[5]=""不征收"",RC[5]=""不征税"",RC[5]=""免征""),0,IF(OR(ISNUMBER(MATCH(RC[5],TaxCatalogItems,0)),AND(OR(RC[5]=""应税"",RC[5]=""征税"",RC[5]=""征收""),RC[-2]<>""未匹配"")),RC[-10]*RC[-1]/1000,RC[8]))"

    ' AE 匹配结果：AI选具体税目视为应税；只填"应税"但仍未匹配时提示补税目。
    wsInvoice.Range(wsInvoice.Cells(firstDataRow, 31), wsInvoice.Cells(lastDataRow, 31)).FormulaR1C1 = _
        "=IF(OR(RC[4]=""不征收"",RC[4]=""不征税"",RC[4]=""免征""),""不征收"",IF(ISNUMBER(MATCH(RC[4],TaxCatalogItems,0)),""应税"",IF(OR(RC[4]=""应税"",RC[4]=""征税"",RC[4]=""征收""),IF(RC[-3]=""未匹配"",""待补税目"",""应税""),RC[8])))"

    ' AG 是否排除：AI选具体税目后进入上方汇总；未匹配只填"应税"不进上方，避免无税目计税。
    wsInvoice.Range(wsInvoice.Cells(firstDataRow, 33), wsInvoice.Cells(lastDataRow, 33)).FormulaR1C1 = _
        "=IF(OR(RC[2]=""不征收"",RC[2]=""不征税"",RC[2]=""免征""),""是"",IF(ISNUMBER(MATCH(RC[2],TaxCatalogItems,0)),""否"",IF(OR(RC[2]=""应税"",RC[2]=""征税"",RC[2]=""征收""),IF(RC[-5]=""未匹配"",""是"",""否""),RC[3])))"

    ' AH 排除原因：人工不征收保留审计痕迹；只填"应税"时提示选择具体税目。
    wsInvoice.Range(wsInvoice.Cells(firstDataRow, 34), wsInvoice.Cells(lastDataRow, 34)).FormulaR1C1 = _
        "=IF(OR(RC[1]=""不征收"",RC[1]=""不征税"",RC[1]=""免征""),""人工确认：不征收"",IF(ISNUMBER(MATCH(RC[1],TaxCatalogItems,0)),""人工确认税目"",IF(OR(RC[1]=""应税"",RC[1]=""征税"",RC[1]=""征收""),IF(RC[-6]=""未匹配"",""请选择具体税目，不要只填应税"",""""),RC[3])))"

    wsInvoice.Columns("AJ:AP").Hidden = True
    wsInvoice.Calculate
End Sub

Sub SetManualConfirmValidation(wsInvoice As Worksheet, totalRows As Long)
    If totalRows <= 0 Then Exit Sub

    ' 人工确认列直接选择具体税目，未匹配发票才能进入上方对应税目汇总。
    With wsInvoice.Range(wsInvoice.Cells(3, 35), wsInvoice.Cells(2 + totalRows, 35)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ManualConfirmList"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "人工确认"
        .InputMessage = "请选择：不征收，或【税目维护】中的具体税目"
    End With
End Sub

Sub SetMappingRuleValidation(wsMapping As Worksheet)
    Dim lastRow As Long
    lastRow = Application.Max(1000, wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row + 200)

    ' F列：对应税目/处理结论。允许选择应税税目，也允许选择“不征收”作为规则结论。
    With wsMapping.Range("F2:F" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=RuleTaxCategoryList"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "对应税目"
        .InputMessage = "请选择应税税目；不征收规则请选择“不征收”"
    End With

    ' H列：是否征税，统一口径。
    With wsMapping.Range("H2:H" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="征税" & Application.International(xlListSeparator) & "不征收" & Application.International(xlListSeparator) & "待确认" & Application.International(xlListSeparator) & "征税/待确认"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "是否征税"
        .InputMessage = "请选择：征税、不征收、待确认、征税/待确认"
    End With

    ' C列和I列用整数校验，减少位数/优先级录入错误。
    With wsMapping.Range("C2:C" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="19"
        .IgnoreBlank = True
        .InputTitle = "编码位数"
        .InputMessage = "请输入1到19之间的整数"
    End With

    With wsMapping.Range("I2:I" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="999"
        .IgnoreBlank = True
        .InputTitle = "优先级"
        .InputMessage = "数字越小优先级越高"
    End With
End Sub

Sub SetSummaryInputValidation(wsSummary As Worksheet)
    With wsSummary.Range("E2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="是" & Application.International(xlListSeparator) & "否"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "是否符合六税两费"
        .InputMessage = "请选择：是 或 否"
    End With
End Sub

Sub SetInvoiceConditionalFormatting(ws As Worksheet, totalRows As Long)
    If totalRows <= 0 Then Exit Sub

    Dim visibleRange As Range
    Dim helperRange As Range
    Dim manualRange As Range
    Dim lastRow As Long

    lastRow = 2 + totalRows
    Set visibleRange = ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 35))  'A:AI，用户可见业务区
    Set helperRange = ws.Range(ws.Cells(3, 36), ws.Cells(lastRow, 42))  'AJ:AP，隐藏辅助区
    Set manualRange = ws.Range(ws.Cells(3, 35), ws.Cells(lastRow, 35))

    ' 清理旧版本整行染色遗留。旧逻辑曾对整行上色，必须按整行清掉，否则右侧会残留大片粉色。
    ws.Rows("3:" & lastRow).FormatConditions.Delete
    ws.Rows("3:" & lastRow).Interior.Pattern = xlNone
    helperRange.Interior.Pattern = xlNone
    ws.Columns("AJ:AP").Hidden = True

    ' 红色：未匹配且尚未人工处理，是当前最需要处理的风险项。
    With visibleRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($A3<>"""",$AB3=""未匹配"",$AI3="""")")
        .Interior.Color = RGB(255, 199, 206)
        .StopIfTrue = True
    End With

    ' 黄色：系统已匹配，但税目/规则本身提示需要复核。
    With visibleRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($A3<>"""",$AI3="""",OR(ISNUMBER(SEARCH(""争议"",$AB3)),ISNUMBER(SEARCH(""待确认"",$AB3)),ISNUMBER(SEARCH(""需确认"",$AB3)),$AE3=""待补税目""))")
        .Interior.Color = RGB(255, 242, 204)
        .StopIfTrue = True
    End With

    ' 绿色：已人工确认且结果有效，表示已处理完成。
    With visibleRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($AI3<>"""",OR($AE3=""应税"",$AE3=""不征收""))")
        .Interior.Color = RGB(226, 239, 218)
        .StopIfTrue = True
    End With

    ' 橙色：AI列填了非标准值，只标AI单元格，不整行染色。
    With manualRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($AI3<>"""",$AE3<>""应税"",$AE3<>""不征收"")")
        .Interior.Color = RGB(244, 176, 132)
    End With
End Sub

'========================================
' 匹配函数：纯逻辑处理
'========================================
Function MatchInvoice(taxCode As String, goodsName As String, amount As Double, sellerTaxId As String, rules As Variant, ruleCount As Long) As Variant
    Dim result(6) As Variant
    Dim matched As Boolean
    Dim i As Long

    matched = False

    ' 遍历所有规则（已按优先级排序）
    For i = 1 To ruleCount
        If IsEmpty(rules(i, 1)) Then Exit For

        ' 1. 编码前缀匹配
        Dim codePrefix As String
        Dim prefixPos As Long

        codePrefix = rules(i, 2)
        prefixPos = 0
        If IsNumeric(rules(i, 3)) Then prefixPos = CLng(rules(i, 3))

        If Len(taxCode) < prefixPos Or Len(codePrefix) = 0 Or prefixPos = 0 Then
            GoTo NextRule
        End If

        If Left(taxCode, prefixPos) <> codePrefix Then
            GoTo NextRule
        End If

        ' 2. 排除关键词检查（优先）
        Dim excludeKeywords As String
        excludeKeywords = rules(i, 10)

        If Len(excludeKeywords) > 0 Then
            If ContainsAnyKeyword(goodsName, excludeKeywords) Then
                GoTo NextRule
            End If
        End If

        ' 3. 关键词判断
        Dim keywordCheck As String
        keywordCheck = rules(i, 5)

        If Len(keywordCheck) > 0 Then
            If Not ContainsAnyKeyword(goodsName, keywordCheck) Then
                GoTo NextRule
            End If
        End If

        ' 匹配成功
        matched = True

        result(0) = rules(i, 6)  '对应税目
        result(1) = rules(i, 7)  '税率
        result(4) = rules(i, 1)  '匹配规则

        ' 判断是否个体户（税号92开头）
        Dim isIndividual As Boolean
        isIndividual = (Len(sellerTaxId) >= 2 And Left(sellerTaxId, 2) = "92")

        ' 判断是否征税
        Dim isTaxable As String
        isTaxable = Trim(CStr(rules(i, 8)))

        ' 计算税额（正负数都直接计算）
        If IsNonTaxableText(isTaxable) Or IsNonTaxableText(result(0)) Or isIndividual Then
            ' 规则标记为不征收、税目列直接选择不征收，或销方为个体户。
            result(2) = 0
            result(3) = "不征收"
            result(5) = "是"
            If isIndividual Then
                result(6) = "个体户"
            ElseIf IsNonTaxableText(result(0)) Then
                result(6) = "不征收"
            Else
                result(6) = NormalizeTaxDecisionLabel(isTaxable)
            End If
        Else
            ' 征税：直接使用原始金额（正负）
            Dim taxRate As Double
            taxRate = 0
            If IsNumeric(rules(i, 7)) Then taxRate = CDbl(rules(i, 7))

            result(2) = amount * taxRate / 1000
            result(3) = "应税"
            result(5) = "否"
            result(6) = ""
        End If

        Exit For

NextRule:
    Next i

    ' 未匹配
    If Not matched Then
        result(0) = "未匹配"
        result(1) = 0
        result(2) = 0
        result(3) = "不征收"
        result(4) = ""
        result(5) = "是"
        result(6) = "无匹配规则"
    End If

    MatchInvoice = result
End Function

'========================================
' 辅助函数：检查关键词
'========================================
Function ContainsAnyKeyword(text As String, keywords As String) As Boolean
    Dim keywordArray() As String
    Dim i As Long

    keywordArray = Split(keywords, "/")

    For i = LBound(keywordArray) To UBound(keywordArray)
        Dim kw As String
        kw = Trim(keywordArray(i))

        If Len(kw) > 0 Then
            If InStr(1, text, kw, vbTextCompare) > 0 Then
                ContainsAnyKeyword = True
                Exit Function
            End If
        End If
    Next i

    ContainsAnyKeyword = False
End Function

'========================================
' 辅助函数：生成季度汇总
'========================================
Sub GenerateSummaryConfirmation(wsInvoice As Worksheet, wsSummary As Worksheet)
    If wsSummary Is Nothing Then Exit Sub

    Dim lastRow As Long
    Dim i As Long
    Dim summaryRow As Long
    Dim dict As Object
    Dim wsMapping As Worksheet
    Dim wsCatalog As Worksheet
    Dim ruleStatusDict As Object
    Dim ruleRow As Long
    Dim invoiceEndRow As Long
    Dim amountRange As String
    Dim pendingTypeRange As String
    Dim manualConfirmRange As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set wsMapping = ThisWorkbook.Worksheets("税目映射规则")
    Set wsCatalog = EnsureTaxCatalogSheet(wsMapping, wsSummary)
    Set ruleStatusDict = CreateObject("Scripting.Dictionary")

    For ruleRow = 2 To wsMapping.Cells(wsMapping.Rows.Count, 1).End(xlUp).Row
        If Len(Trim(CStr(wsMapping.Cells(ruleRow, 1).Value))) > 0 Then
            ruleStatusDict(Trim(CStr(wsMapping.Cells(ruleRow, 1).Value))) = Trim(CStr(wsMapping.Cells(ruleRow, 8).Value))
        End If
    Next ruleRow

    lastRow = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 3 Then
        wsInvoice.Range(wsInvoice.Cells(3, 40), wsInvoice.Cells(lastRow, 40)).ClearContents
    End If

    invoiceEndRow = Application.Max(3, lastRow)
    amountRange = "'发票明细'!$T$3:$T$" & invoiceEndRow
    pendingTypeRange = "'发票明细'!$AN$3:$AN$" & invoiceEndRow
    manualConfirmRange = "'发票明细'!$AI$3:$AI$" & invoiceEndRow

    summaryRow = SetupQuarterSummaryTemplate(wsSummary, wsCatalog, lastRow)

    For i = 3 To lastRow
        Dim taxCategory As String
        Dim amount As Double
        Dim isExcluded As String
        Dim matchedRule As String
        Dim manualConfirm As String

        taxCategory = Trim(CStr(wsInvoice.Cells(i, 28).Value))
        isExcluded = Trim(CStr(wsInvoice.Cells(i, 33).Value))
        matchedRule = Trim(CStr(wsInvoice.Cells(i, 32).Value))
        manualConfirm = NormalizeTaxDecisionLabel(wsInvoice.Cells(i, 35).Value)
        amount = 0
        If IsNumeric(wsInvoice.Cells(i, 20).Value) Then
            amount = CDbl(wsInvoice.Cells(i, 20).Value)
        End If

        ' 人工确认优先于自动匹配，避免已判定"不征收"的项目继续留在待处理汇总中。
        If IsNonTaxableText(manualConfirm) Then GoTo NextInvoice
        If IsTaxableText(manualConfirm) Then isExcluded = "否"

        ' 跳过排除的项目，但未匹配除外（未匹配需要统计）
        If isExcluded <> "否" And taxCategory <> "未匹配" Then GoTo NextInvoice

        Dim isYellow As Boolean, isRed As Boolean
        isRed = (taxCategory = "未匹配")
        isYellow = (InStr(1, taxCategory, "争议", vbTextCompare) > 0 Or _
                    InStr(1, taxCategory, "待确认", vbTextCompare) > 0 Or _
                    InStr(1, taxCategory, "需确认", vbTextCompare) > 0)

        ' 检查匹配规则的状态（是否为"待确认"或"征税/待确认"）
        If matchedRule <> "" And Not isYellow Then
            Dim ruleStatus As String
            If ruleStatusDict.Exists(matchedRule) Then
                ruleStatus = ruleStatusDict(matchedRule)
                If InStr(1, ruleStatus, "待确认", vbTextCompare) > 0 Then
                    isYellow = True
                End If
            End If
        End If

        ' 统计未匹配和需要人工处理的项目
        If (isYellow Or isRed) And taxCategory <> "" Then
            wsInvoice.Cells(i, 40).Value = taxCategory

            If Not dict.Exists(taxCategory) Then
                dict.Add taxCategory, Array(0, 0)
            End If

            Dim arr As Variant
            arr = dict(taxCategory)
            arr(0) = arr(0) + 1
            arr(1) = arr(1) + amount
            dict(taxCategory) = arr
        End If

NextInvoice:
    Next i

    If dict.Count = 0 Then Exit Sub

    Dim key As Variant
    For Each key In dict.Keys
        Dim data As Variant
        data = dict(key)

        Dim suggestion As String
        Dim rowColor As Long

        If key = "未匹配" Then
            suggestion = "AI填不征收会移入下方备查行；AI填应税仍需补税目规则"
            rowColor = RGB(255, 199, 206)
        Else
            suggestion = "请人工核实税目分类"
            rowColor = RGB(255, 242, 204)
        End If

        ' 待处理：AI为空；未匹配即使填了"应税"，仍需补规则，所以继续留在待处理。
        wsSummary.Cells(summaryRow, 10).Value = key
        wsSummary.Cells(summaryRow, 1).Formula = "=IF(B" & summaryRow & "=0,"""",$J" & summaryRow & ")"
        wsSummary.Cells(summaryRow, 2).Formula = "=COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ","""")" & _
                                                  "+IF($J" & summaryRow & "=""未匹配"",COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""应税"")+COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""征税"")+COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""征收""),0)"
        wsSummary.Cells(summaryRow, 3).Formula = "=SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ","""")" & _
                                                  "+IF($J" & summaryRow & "=""未匹配"",SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""应税"")+SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""征税"")+SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""征收""),0)"
        wsSummary.Cells(summaryRow, 4).Formula = "=IF(B" & summaryRow & "=0,"""",C" & summaryRow & "*0.3/1000)"
        wsSummary.Cells(summaryRow, 5).Formula = "=IF(B" & summaryRow & "=0,"""",""" & suggestion & """)"

        wsSummary.Range(wsSummary.Cells(summaryRow, 1), wsSummary.Cells(summaryRow, 5)).Interior.Color = rowColor

        summaryRow = summaryRow + 1

        ' 已确认不征收：从待处理移出，但在汇总表保留备查痕迹，便于复核。
        wsSummary.Cells(summaryRow, 10).Value = key
        wsSummary.Cells(summaryRow, 1).Formula = "=IF(B" & summaryRow & "=0,"""",$J" & summaryRow & "&"" - 已确认不征收"")"
        wsSummary.Cells(summaryRow, 2).Formula = "=COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""不征收"")" & _
                                                  "+COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""不征税"")" & _
                                                  "+COUNTIFS(" & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""免征"")"
        wsSummary.Cells(summaryRow, 3).Formula = "=SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""不征收"")" & _
                                                  "+SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""不征税"")" & _
                                                  "+SUMIFS(" & amountRange & "," & pendingTypeRange & ",$J" & summaryRow & "," & manualConfirmRange & ",""免征"")"
        wsSummary.Cells(summaryRow, 4).Formula = "=IF(B" & summaryRow & "=0,"""",0)"
        wsSummary.Cells(summaryRow, 5).Formula = "=IF(B" & summaryRow & "=0,"""",""已人工确认不征收，留痕备查"")"
        wsSummary.Range(wsSummary.Cells(summaryRow, 1), wsSummary.Cells(summaryRow, 5)).Interior.Color = RGB(226, 239, 218)

        summaryRow = summaryRow + 1
    Next key

    wsSummary.Columns(10).Hidden = True
End Sub

'========================================
' 手工确认兼容入口：恢复联动公式，避免旧事件代码覆盖公式
'========================================
Sub ApplyManualConfirmationRow(wsInvoice As Worksheet, rowNum As Long)
    Dim lastRow As Long

    lastRow = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 3 Then
        Call InstallManualConfirmationFormulas(wsInvoice, lastRow - 2)
    End If
End Sub

Sub RefreshManualConfirmation()
    Call RefreshAfterManualConfirmation(True)
End Sub

Sub RefreshAfterManualConfirmation(Optional showMessage As Boolean = True)
    Dim wsInvoice As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim handledCount As Long

    On Error GoTo ErrorHandler

    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")

    Call EnsureTaxCatalogSheet(ThisWorkbook.Worksheets("税目映射规则"), wsSummary)
    Call SetMappingRuleValidation(ThisWorkbook.Worksheets("税目映射规则"))
    Call SetSummaryInputValidation(wsSummary)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    lastRow = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row
    For i = 3 To lastRow
        If Len(Trim(CStr(wsInvoice.Cells(i, 35).Value))) > 0 Then
            handledCount = handledCount + 1
        End If
    Next i

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

    If showMessage Then
        MsgBox "人工确认刷新完成。" & vbCrLf & _
               "已处理人工确认行数：" & handledCount & vbCrLf & _
               "公式联动已恢复，请检查【季度汇总】。", vbInformation, "刷新完成"
    End If
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "人工确认刷新失败：" & Err.Description, vbCritical, "错误"
End Sub

Sub HandleWorkbookSheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo CleanUp

    If Sh Is Nothing Then Exit Sub

    If Sh.Name = "税目维护" Then
        If Intersect(Target, Sh.Range("A:C")) Is Nothing Then Exit Sub
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim wsSummary As Worksheet
        Dim wsInvoice As Worksheet
        Dim wsMapping As Worksheet

        Set wsSummary = ThisWorkbook.Worksheets("季度汇总")
        Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
        Set wsMapping = ThisWorkbook.Worksheets("税目映射规则")

        Call EnsureTaxCatalogSheet(wsMapping, wsSummary)
        Call SetMappingRuleValidation(wsMapping)
        Call SetSummaryInputValidation(wsSummary)

        If wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row >= 3 Then
            Call SetManualConfirmValidation(wsInvoice, wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row - 2)
            Call InstallManualConfirmationFormulas(wsInvoice, wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row - 2)
            Call SetInvoiceConditionalFormatting(wsInvoice, wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row - 2)
        End If

        Call GenerateSummaryConfirmation(wsInvoice, wsSummary)
        Application.Calculate
    ElseIf Sh.Name = "发票明细" Then
        If Intersect(Target, Sh.Columns(35)) Is Nothing Then Exit Sub
        Application.EnableEvents = False
        Sh.Calculate
        ThisWorkbook.Worksheets("季度汇总").Calculate
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub RefreshQuarterSummary()
    Dim wsInvoice As Worksheet
    Dim wsSummary As Worksheet

    On Error Resume Next
    Set wsInvoice = ThisWorkbook.Worksheets("发票明细")
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")
    On Error GoTo 0

    If wsInvoice Is Nothing Or wsSummary Is Nothing Then Exit Sub

    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Call GenerateSummaryConfirmation(wsInvoice, wsSummary)

CleanUp:
    Application.ScreenUpdating = True
End Sub

Function NormalizeTaxDecisionLabel(value As Variant) As String
    Dim text As String
    text = Trim(CStr(value))

    If text = "不征税" Or text = "不征收" Or text = "免征" Then
        NormalizeTaxDecisionLabel = "不征收"
    ElseIf text = "征税" Or text = "应税" Or text = "征收" Then
        NormalizeTaxDecisionLabel = "应税"
    Else
        NormalizeTaxDecisionLabel = text
    End If
End Function

Function IsNonTaxableText(value As Variant) As Boolean
    IsNonTaxableText = (NormalizeTaxDecisionLabel(value) = "不征收")
End Function

Function IsTaxableText(value As Variant) As Boolean
    IsTaxableText = (NormalizeTaxDecisionLabel(value) = "应税")
End Function

'========================================
' 辅助函数：规则排序
'========================================
Sub SortRulesByPriority(ByRef rules As Variant, ruleCount As Long)
    Dim i As Long, j As Long, k As Long
    Dim temp As Variant

    For i = 1 To ruleCount - 1
        For j = i + 1 To ruleCount
            Dim priority1 As Long, priority2 As Long

            priority1 = 999
            priority2 = 999

            If IsNumeric(rules(i, 9)) Then priority1 = CLng(rules(i, 9))
            If IsNumeric(rules(j, 9)) Then priority2 = CLng(rules(j, 9))

            If priority1 > priority2 Then
                For k = 1 To 11
                    temp = rules(i, k)
                    rules(i, k) = rules(j, k)
                    rules(j, k) = temp
                Next k
            End If
        Next j
    Next i
End Sub

'========================================
' 辅助函数：检查格式
'========================================
Sub CheckTaxCodeFormat(wsRaw As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim hasError As Boolean
    Dim errorCount As Long
    Dim errorColumns As String

    hasError = False
    errorCount = 0
    errorColumns = ""
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row

    ' 需要检查的列：2=发票代码, 3=发票号码, 4=数电发票号码, 5=销方识别号, 7=购方识别号, 10=税收分类编码
    Dim checkCols As Variant
    Dim colNames As Variant
    checkCols = Array(2, 3, 4, 5, 7, 10)
    colNames = Array("发票代码", "发票号码", "数电发票号码", "销方识别号", "购方识别号", "税收分类编码")

    For i = 3 To lastRow
        Dim colIdx As Long
        For colIdx = LBound(checkCols) To UBound(checkCols)
            Dim col As Long
            col = checkCols(colIdx)

            If Not IsEmpty(wsRaw.Cells(i, col).Value) Then
                Dim cellText As String
                cellText = CStr(wsRaw.Cells(i, col).text)

                If InStr(cellText, "E+") > 0 Or InStr(cellText, "e+") > 0 Then
                    hasError = True
                    errorCount = errorCount + 1
                    wsRaw.Cells(i, col).Interior.Color = RGB(255, 199, 206)

                    ' 记录出错的列名
                    If InStr(errorColumns, colNames(colIdx)) = 0 Then
                        If errorColumns <> "" Then errorColumns = errorColumns & "、"
                        errorColumns = errorColumns & colNames(colIdx)
                    End If
                End If
            End If
        Next colIdx
    Next i

    If hasError Then
        Dim msg As String
        msg = "?? 发现长数字格式错误！" & vbCrLf & vbCrLf & _
              "共 " & errorCount & " 处数据显示为科学计数法" & vbCrLf & _
              "涉及列：" & errorColumns & vbCrLf & vbCrLf & _
              "解决方法：" & vbCrLf & _
              "1. 粘贴前先将第2、3、4、5、7、10列设置为文本格式" & vbCrLf & _
              "2. 选中这些列 → 右键 → 设置单元格格式 → 文本" & vbCrLf & _
              "3. 然后再粘贴数据" & vbCrLf & vbCrLf & _
              "是否继续运行？（结果可能不准确）"

        Dim response As VbMsgBoxResult
        response = MsgBox(msg, vbExclamation + vbYesNo, "格式错误")

        If response = vbNo Then
            Err.Raise vbObjectError + 1, "CheckTaxCodeFormat", "用户取消"
        End If
    End If
End Sub


