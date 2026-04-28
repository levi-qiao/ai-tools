Sub RunMatch()
    '========================================
    ' 印花税自动匹配 - V4优化版
    ' 优化：解决卡死问题，提升性能
    '========================================

    Dim wsRaw As Worksheet
    Dim wsInvoice As Worksheet
    Dim wsMapping As Worksheet
    Dim lastRowRaw As Long
    Dim lastRowMapping As Long
    Dim i As Long
    Dim outputRow As Long
    Dim startTime As Double

    On Error GoTo ErrorHandler
    startTime = Timer

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
    On Error GoTo ErrorHandler

    If wsRaw Is Nothing Or wsInvoice Is Nothing Or wsMapping Is Nothing Then
        MsgBox "找不到必需的工作表！", vbCritical, "错误"
        GoTo CleanUp
    End If

    ' 1.5 检查原始数据格式
    Call CheckTaxCodeFormat(wsRaw)

    ' 2. 清空发票明细旧数据（优化：只清除使用的区域）
    Application.StatusBar = "正在清空旧数据..."
    Dim lastRowInvoice As Long
    lastRowInvoice = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row
    If lastRowInvoice >= 3 Then
        ' 只清除实际使用的列（1-40列），避免清除整行
        wsInvoice.Range(wsInvoice.Cells(3, 1), wsInvoice.Cells(lastRowInvoice, 40)).Clear
    End If

    ' 3. 清空季度汇总表
    Dim wsSummary As Worksheet
    On Error Resume Next
    Set wsSummary = ThisWorkbook.Worksheets("季度汇总")
    On Error GoTo ErrorHandler

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
        ReDim rules(1 To ruleCount, 1 To 12)

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
            rules(j - 1, 12) = Trim(CStr(wsMapping.Cells(j, 12).Value)) '负数发票处理
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
            textData(rowIdx, colIdx) = wsRaw.Cells(rowIdx + 2, textColumns(colIdx)).Text
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

    ReDim outputData(1 To validRows, 1 To 40)

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
        Dim sellerName As String

        taxCode = Trim(textData(i, 6))  '税收分类编码（使用Text）
        goodsName = Trim(CStr(rawData(i, 12)))
        sellerName = Trim(CStr(rawData(i, 6)))

        amount = 0
        If IsNumeric(rawData(i, 20)) Then
            amount = CDbl(rawData(i, 20))
        End If

        ' 执行匹配
        Dim matchResult As Variant
        matchResult = MatchInvoice(taxCode, goodsName, amount, sellerName, rules, ruleCount)

        ' 写入匹配结果（列28-35）
        outputData(outputRow, 28) = matchResult(0)  '匹配税目
        outputData(outputRow, 29) = matchResult(1)  '税率
        outputData(outputRow, 30) = matchResult(2)  '应纳税额
        outputData(outputRow, 31) = matchResult(3)  '匹配结果
        outputData(outputRow, 32) = matchResult(4)  '匹配规则
        outputData(outputRow, 33) = matchResult(5)  '是否排除
        outputData(outputRow, 34) = matchResult(6)  '排除原因
        outputData(outputRow, 35) = ""              '人工确认

NextRow:
    Next i

    ' 9. 批量写入结果（性能优化：一次性写入）
    Application.StatusBar = "正在写入结果..."
    If outputRow > 0 Then
        wsInvoice.Range(wsInvoice.Cells(3, 1), wsInvoice.Cells(2 + outputRow, 40)).Value = outputData

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

        ' 设置行颜色（批量处理）
        Call SetRowColors(wsInvoice, outputRow)
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
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Application.Calculate
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "发生错误：" & Err.Description & vbCrLf & _
           "错误代码：" & Err.Number, vbCritical, "错误"
End Sub

'========================================
' 匹配函数：纯逻辑处理
'========================================
Function MatchInvoice(taxCode As String, goodsName As String, amount As Double, sellerName As String, rules As Variant, ruleCount As Long) As Variant
    Dim result(6) As Variant
    Dim matched As Boolean
    Dim i As Long

    matched = False

    ' 特殊判断：个体工商户不征税
    If InStr(1, sellerName, "个体工商户", vbTextCompare) > 0 Then
        result(0) = "不征税"
        result(1) = "0"
        result(2) = "0"
        result(3) = "不征税"
        result(4) = "个体工商户"
        result(5) = "否"
        result(6) = ""
        MatchInvoice = result
        Exit Function
    End If

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

        ' 3. 名称包含检查
        Dim nameContains As String
        nameContains = rules(i, 4)

        If Len(nameContains) > 0 Then
            If InStr(1, goodsName, nameContains, vbTextCompare) = 0 Then
                GoTo NextRule
            End If
        End If

        ' 4. 关键词判断
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

        ' 判断是否排除
        Dim isTaxable As String
        Dim negativeHandling As String
        isTaxable = Trim(CStr(rules(i, 8)))
        negativeHandling = Trim(CStr(rules(i, 12)))

        ' 优先判断负数发票
        If amount < 0 Then
            ' 负数发票：检查"负数发票处理"字段
            If negativeHandling = "征税" Then
                ' 负数发票也征税
                Dim taxRate As Double
                taxRate = 0
                If IsNumeric(rules(i, 7)) Then taxRate = CDbl(rules(i, 7))

                result(2) = Abs(amount) * taxRate / 1000  '使用绝对值
                result(3) = "应税"
                result(5) = "否"
                result(6) = ""
            Else
                ' 负数发票排除
                result(2) = 0
                result(3) = "不征收"
                result(5) = "是"
                result(6) = "红字发票"
            End If
        ElseIf isTaxable = "不征税" Or isTaxable = "不征收" Then
            ' 正数发票但规则标记为不征收
            result(2) = 0
            result(3) = "不征收"
            result(5) = "是"
            result(6) = isTaxable
        Else
            ' 正数发票且征税
            Dim taxRate2 As Double
            taxRate2 = 0
            If IsNumeric(rules(i, 7)) Then taxRate2 = CDbl(rules(i, 7))

            result(2) = amount * taxRate2 / 1000
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
' 辅助函数：批量设置行颜色（性能优化）
'========================================
Sub SetRowColors(ws As Worksheet, totalRows As Long)
    Dim i As Long
    Dim taxCategory As String
    Dim matchStatus As String
    Dim matchedRule As String

    For i = 3 To 2 + totalRows
        matchStatus = CStr(ws.Cells(i, 31).Value)
        taxCategory = CStr(ws.Cells(i, 28).Value)
        matchedRule = CStr(ws.Cells(i, 32).Value)

        ' 未匹配 - 红色
        If taxCategory = "未匹配" And Len(Trim(matchedRule)) = 0 Then
            ws.Rows(i).Interior.Color = RGB(255, 199, 206)
        ' 需确认 - 黄色
        ElseIf matchStatus = "应税" Then
            If InStr(1, taxCategory, "争议", vbTextCompare) > 0 Or _
               InStr(1, taxCategory, "待确认", vbTextCompare) > 0 Or _
               InStr(1, taxCategory, "需确认", vbTextCompare) > 0 Then
                ws.Rows(i).Interior.Color = RGB(255, 242, 204)
            End If
        End If
    Next i
End Sub

'========================================
' 辅助函数：生成季度汇总
'========================================
Sub GenerateSummaryConfirmation(wsInvoice As Worksheet, wsSummary As Worksheet)
    If wsSummary Is Nothing Then Exit Sub

    Dim lastRow As Long
    Dim i As Long
    Dim summaryRow As Long
    Dim dict As Object

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsInvoice.Cells(wsInvoice.Rows.Count, 1).End(xlUp).Row

    For i = 3 To lastRow
        Dim taxCategory As String
        Dim amount As Double
        Dim isExcluded As String

        taxCategory = Trim(CStr(wsInvoice.Cells(i, 28).Value))
        isExcluded = Trim(CStr(wsInvoice.Cells(i, 33).Value))
        amount = 0
        If IsNumeric(wsInvoice.Cells(i, 20).Value) Then
            amount = CDbl(wsInvoice.Cells(i, 20).Value)
        End If

        If isExcluded <> "否" Then GoTo NextInvoice

        Dim isYellow As Boolean, isRed As Boolean
        isRed = (taxCategory = "未匹配")
        isYellow = (InStr(1, taxCategory, "争议", vbTextCompare) > 0 Or _
                    InStr(1, taxCategory, "待确认", vbTextCompare) > 0 Or _
                    InStr(1, taxCategory, "需确认", vbTextCompare) > 0)

        If (isYellow Or isRed) And taxCategory <> "" Then
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

    summaryRow = 21

    Dim key As Variant
    For Each key In dict.Keys
        Dim data As Variant
        data = dict(key)

        Dim suggestion As String
        Dim rowColor As Long

        If key = "未匹配" Then
            suggestion = "请添加匹配规则或检查税收分类编码"
            rowColor = RGB(255, 199, 206)
        Else
            suggestion = "请人工核实税目分类"
            rowColor = RGB(255, 242, 204)
        End If

        wsSummary.Cells(summaryRow, 1).Value = key
        wsSummary.Cells(summaryRow, 2).Value = data(0)
        wsSummary.Cells(summaryRow, 3).Value = data(1)
        wsSummary.Cells(summaryRow, 4).Value = data(1) * 0.3 / 1000
        wsSummary.Cells(summaryRow, 5).Value = suggestion

        wsSummary.Rows(summaryRow).Interior.Color = rowColor

        summaryRow = summaryRow + 1
    Next key
End Sub

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
                For k = 1 To 12
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
                cellText = CStr(wsRaw.Cells(i, col).Text)

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
        msg = "⚠️ 发现长数字格式错误！" & vbCrLf & vbCrLf & _
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
