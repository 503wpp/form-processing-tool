Attribute VB_Name = "开始拆分"
Public xlApp As Variant
Public xlBook As Variant
Public xlsheet As Variant
Public lastRow As Variant

Sub 开始拆分()
    Dim colData As Variant
    Dim Keyword_1 As String, Keyword_2 As String, Keyword_3 As String, Keyword_4 As String, Keyword_5 As String, Keyword_6 As String, Keyword_7 As String, Keyword_8 As String
    
    Keyword_1 = "关于"
    Keyword_2 = "_"
    Keyword_3 = "分公司("
    Keyword_4 = ")的"
    Keyword_5 = "分公司（"
    Keyword_6 = "）的"
    Keyword_7 = "分公司"
    Keyword_8 = "的"
    Keyword_9 = "的设计"
    Keyword_11 = "-20"
    
    ' 对待办数据进行操作
    Set xlApp = CreateObject("ket.Application")
    If sourceFilePath_1 = "1" Then
        MsgBox ("请导入待办数据")
        Exit Sub
    ElseIf sourceFilePath_1 <> "" Then
        Set xlBook = xlApp.Workbooks.Open(sourceFilePath_1)  '打开指定路径指定名称文件
    End If
    ' 不显示WPS界面
    xlApp.Visible = False
    
    Form1.ListView1.ListItems.Clear               '清空列表
    Form1.ListView1.ColumnHeaders.Clear           '清空列表头
    Form1.ListView1.View = lvwReport              '设置列表显示方式
    
    ' 初始化进度条
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.Label3.Caption = "运行进度：正在处理待办数据，请稍等......" & "0%"
    
    ' 指定要查找的字段名及其所在的工作表名字的一部分
    fieldName = "待办信息"
    sheetName = "*" & fieldName & "*"

    ' 在 Workbook 中查找名字包含指定字段名的工作表
    For Each i In xlBook.Worksheets
        If i.Name Like sheetName Then
            Set xlsheet = xlBook.Worksheets(i.Name)
            ' 找到了名字包含指定字段名的工作表
            ' 找到最后一行
            lastRow = xlsheet.Application.WorksheetFunction.CountA(xlsheet.range("A:A"))
                
            ' 在工作表的 A 列前插入 5 列空白列
            xlsheet.range("A1").EntireColumn.Resize(, 8).Insert Shift:=xlToRight
            xlsheet.range("A1") = "地市"
            xlsheet.range("B1") = "分公司"
            xlsheet.range("C1") = "设计阶段"
            xlsheet.range("D1") = "项目名称"
            xlsheet.range("E1") = "项目编号"
            xlsheet.range("F1") = "任务名称"
            xlsheet.range("G1") = "日期"
            xlsheet.range("H1") = "台账名称"
 
            colData = xlsheet.range("I1:I" & lastRow).Value2
            ' 遍历第 6 列的所有单元格，并处理每个单元格的数据
            For Row = 2 To UBound(colData, 1)
                ProgressValue = Row
                ' 使用正则表达式提取起始关键词和结束关键词之间的数据
                Set regex = CreateObject("VBScript.RegExp")
                
                ' 填写分公司
                regex.Pattern = Keyword_2 & "(.*?)" & Keyword_7
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    xlsheet.range("B" & Row) = Match.SubMatches(0)
                End If
                    
                ' 填写地市
                a = xlsheet.range("B" & Row)
                If a = "北碚" Or a = "合川" Or a = "铜梁" Or a = "潼南" Then
                    xlsheet.range("A" & Row) = "北碚片区"
                Else
                    xlsheet.range("A" & Row) = "永川片区"
                End If
                
                ' 填写设计阶段
                regex.Pattern = "【" & "(.*?)" & "】"
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    xlsheet.range("C" & Row) = Match.SubMatches(0)
                End If
                    
                ' 填写项目名称
                regex.Pattern = Keyword_1 & "(.*?)" & Keyword_7
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    xlsheet.range("D" & Row) = Match.SubMatches(0) & "分公司"
                End If
                    
                ' 填写项目编号
                If InStr(colData(Row, 1), "分公司(") > 0 Then
                    regex.Pattern = Keyword_3 & "(.*?)" & Keyword_4
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        trimmedStr = Match.SubMatches(0)
                        ' 使用 Replace 函数去掉字符串 trimmedStr 左右两边的括号
                        trimmedStr = Replace(trimmedStr, "(", "")
                        trimmedStr = Replace(trimmedStr, ")", "")
                        xlsheet.range("E" & Row) = trimmedStr
                    End If
                ElseIf InStr(colData(Row, 1), "分公司（") > 0 Then
                    regex.Pattern = Keyword_5 & "(.*?)" & Keyword_6
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("E" & Row) = Match.SubMatches(0)
                    End If
                End If
                    
                ' 填写任务名称
                regex.Pattern = Keyword_8 & "(.*?)" & Keyword_9
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    num = Match.SubMatches(0)
                    xlsheet.range("F" & Row) = num
                End If
                
                If xlsheet.range("D" & Row).Value Like "*家庭宽带*" Or xlsheet.range("D" & Row).Value Like "*商业宽带*" Then
                    ' 填写日期
                    regex.Pattern = Keyword_11 & "(.*?)" & Keyword_9
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("G" & Row) = "20" & Match.SubMatches(0)
                    End If
                End If
                
                ' 更新进度条
                Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 100 / lastRow
                Form1.Label3.Caption = "运行进度：正在处理待办数据，请稍等......" & Form1.ProgressBar1.Value & "%"
                
            Next Row
        End If
        
        Form1.ListView1.ColumnHeaders.Add , , "拆分完成！", 2000 '给列表中添加列名
        
    Next i
    Form1.Label3.Caption = "运行进度：拆分完成！ 100%"
    MsgBox ("拆分完成")
End Sub
