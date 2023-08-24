Attribute VB_Name = "拆分标题"
Sub split(ByVal sheetName As Variant, ByVal xlBook_3 As Variant, ByRef JiaKuan As Variant, ByRef ZhuanXian As Variant, ByRef xlsheet As Variant)
    
    Dim colData As Variant
    Dim Keyword_1 As String, Keyword_2 As String, Keyword_3 As String, Keyword_4 As String, Keyword_5 As String, Keyword_6 As String, Keyword_7 As String, Keyword_8 As String
    Dim lastRow As Integer
    
    Keyword_1 = "【"
    Keyword_2 = "】"
    Keyword_3 = "关于"
    Keyword_4 = "分公司"
    Keyword_5 = "分公司("
    Keyword_6 = ")的"
    Keyword_7 = "分公司（"
    Keyword_8 = "）的"
    Keyword_9 = "的设计"
    Keyword_10 = "_"
    Keyword_11 = "-20"
    
    ' 获取已办信息工作表
    For Each i In xlBook_3.Worksheets
        If i.Name Like sheetName Then
            Set xlsheet = xlBook_3.Worksheets(i.Name)
            
            ' 插入缺少的列
            If xlsheet.range("A1").Value <> "项目名称" Then
                xlsheet.range("A1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("A1") = "项目名称"
            End If
            If xlsheet.range("B1").Value <> "专业名称" Then
                xlsheet.range("B1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("B1") = "专业名称"
            End If
            If xlsheet.range("C1").Value <> "单项名称" Then
                xlsheet.range("C1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("C1") = "单项名称"
            End If
            If xlsheet.range("D1").Value <> "片区" Then
                xlsheet.range("D1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("D1") = "片区"
            End If
            If xlsheet.range("E1").Value <> "分公司" Then
                xlsheet.range("E1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("E1") = "分公司"
            End If
            If xlsheet.range("F1").Value <> "设计阶段" Then
                xlsheet.range("F1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("F1") = "设计阶段"
            End If
            If xlsheet.range("G1").Value <> "项目编号" Then
                xlsheet.range("G1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("G1") = "项目编号"
            End If
            If xlsheet.range("H1").Value <> "任务名称" Then
                xlsheet.range("H1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("H1") = "任务名称"
            End If
            If xlsheet.range("I1").Value <> "日期" Then
                xlsheet.range("I1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("I1") = "日期"
            End If
            
            ' 找到最后一行
            lastRow = xlsheet.Application.WorksheetFunction.CountA(xlsheet.range("J:J"))
            
            ' 获取标题列数据
            colData = xlsheet.range("J1:J" & lastRow).Value2
            Dim regex As Object
            ' 遍历标题列的所有单元格，并处理每个单元格的数据
            For Row = 2 To UBound(colData, 1)
                'If Len(xlSheet.Range("A" & Row).Value) = 0 Then
                    ProgressValue = Row
                    ' 使用正则表达式提取起始关键词和结束关键词之间的数据
                    Set regex = CreateObject("VBScript.RegExp")
                    
                    ' 填写项目名称
                    regex.Pattern = Keyword_3 & "(.*?)" & Keyword_4
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("A" & Row) = Match.SubMatches(0) & "分公司"
                    End If

                    '填写单项名称
                    a = xlsheet.range("A" & Row)
                    If a Like "*集团专线*" Then
                        xlsheet.range("C" & Row) = "集团专线"
                    ElseIf a Like "*家庭宽带*" Then
                        xlsheet.range("C" & Row) = "家庭宽带"
                    ElseIf a Like "*商业宽带*" Then
                        xlsheet.range("C" & Row) = "商业宽带"
                    ElseIf a Like "*商宽重要客户预覆盖*" Then
                        xlsheet.range("C" & Row) = "预覆盖"
                    End If
                    
                    '填写专业名称
                    a = xlsheet.range("C" & Row)
                    If a = "集团专线" Or a = "预覆盖" Then
                        xlsheet.range("B" & Row) = "专线"
                    Else
                        xlsheet.range("B" & Row) = "家宽"
                    End If
                    
                    ' 填写分公司
                    regex.Pattern = Keyword_10 & "(.*?)" & Keyword_4
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("E" & Row) = Match.SubMatches(0)
                    End If
                    
                    ' 填写地市
                    a = xlsheet.range("E" & Row)
                    If a = "北碚" Or a = "合川" Or a = "铜梁" Or a = "潼南" Then
                        xlsheet.range("D" & Row) = "北碚片区"
                    Else
                        xlsheet.range("D" & Row) = "永川片区"
                    End If
                    
                    ' 填写设计阶段
                    regex.Pattern = Keyword_1 & "(.*?)" & Keyword_2
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("F" & Row) = Match.SubMatches(0)
                    End If

                    ' 填写项目编号
                    If InStr(colData(Row, 1), "分公司(") > 0 Then
                        regex.Pattern = Keyword_5 & "(.*?)" & Keyword_6
                        Set matches = regex.Execute(colData(Row, 1))
                        If matches.Count > 0 Then
                            Set Match = matches(0)
                            trimmedStr = Match.SubMatches(0)
                            ' 使用 Replace 函数去掉字符串 trimmedStr 左右两边的括号
                            trimmedStr = Replace(trimmedStr, "(", "")
                            trimmedStr = Replace(trimmedStr, ")", "")
                            xlsheet.range("G" & Row) = trimmedStr
                        End If
                    ElseIf InStr(colData(Row, 1), "分公司（") > 0 Then
                        regex.Pattern = Keyword_7 & "(.*?)" & Keyword_8
                        Set matches = regex.Execute(colData(Row, 1))
                        If matches.Count > 0 Then
                            Set Match = matches(0)
                            trimmedStr = Match.SubMatches(0)
                            xlsheet.range("G" & Row) = trimmedStr
                        End If
                    End If
                    
                    ' 填写任务名称
                    regex.Pattern = "的" & "(.*?)" & Keyword_9
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("H" & Row) = Match.SubMatches(0)
                    End If
                    
                    If xlsheet.range("B" & Row).Value = "家宽" Then
                        ' 填写日期
                        regex.Pattern = Keyword_11 & "(.*?)" & Keyword_9
                        Set matches = regex.Execute(colData(Row, 1))
                        If matches.Count > 0 Then
                            Set Match = matches(0)
                            xlsheet.range("I" & Row) = "20" & Match.SubMatches(0)
                        End If
                    End If
                'End If
            Next Row
        End If
    Next i
    
    ' 选取所有列，并设定自动调整列宽
    xlsheet.Columns("A:I").AutoFit
    
    Design_distributions = xlsheet.range("F1:F" & lastRow).Value
    For i = lastRow To 2 Step -1
        If Design_distributions(i, 1) = "设计勘察" Then
            xlsheet.Rows(i).Delete
        End If
    Next i
    
    If sheetName = "*待办信息*" Then
        ' 遍历每个单元格，并将不同的元素添加到字典中
        For i = lastRow To 2 Step -1
            a = xlsheet.cells(i, 8).Value
            If a Like "*SCM领用" Or a Like "*SCM领料" Or a Like "*辅材分摊" Or a Like "*跨项目调拨" Then
                xlsheet.Rows(i).Delete
            End If
        Next i
    End If
            

    ' 统计第B列家宽已办数量，并更新字典值
    rangeString = "A1:A" & lastRow

    For Each k In JiaKuan.keys
        If Len(k) = 0 Then
            Exit For
        Else
            formulaString = "=COUNTIF(" & rangeString & "," & Chr(34) & k & Chr(34) & ")"
            Count = xlsheet.Evaluate(formulaString)
            JiaKuan(k) = Count
        End If
    Next k
    ' 统计第B列专线已办数量，并更新字典值
    For Each k In ZhuanXian.keys
        If Len(k) = 0 Then
            Exit For
        Else
            formulaString = "=COUNTIF(" & rangeString & "," & Chr(34) & k & Chr(34) & ")"
            Count = xlsheet.Evaluate(formulaString)
            ZhuanXian(k) = Count
        End If
    Next k
End Sub
