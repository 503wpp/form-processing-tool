Attribute VB_Name = "开始对比"
Sub 开始对比()
    
    ' 初始化进度条
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "运行进度：正在读取台账数据，请稍等......" & "0%"
    
    Set xlApp_2 = CreateObject("ket.Application") '网上查询说et.Application及ket.Application也能
    If sourceFilePath_2 = "1" Then
        MsgBox ("请导入台账数据")
        Exit Sub
    ElseIf sourceFilePath_2 <> "" Then
        Set xlBook_2 = xlApp_2.Workbooks.Open(sourceFilePath_2)  '打开指定路径指定名称文件
    End If
    ' 不显示WPS界面
    xlApp_2.Visible = False
    
    
    
    ' 初始化列表
    Form1.ListView1.ListItems.Clear               '清空列表
    Form1.ListView1.ColumnHeaders.Clear           '清空列表头
    Form1.ListView1.View = lvwReport              '设置列表显示方式
    Form1.ListView1.GridLines = True              '显示网络线
     
    ' 获取指定名字的工作表
    Set xlSheet_2 = xlBook_2.Worksheets("23年台账明细")
    ' 找到最后一行
    lastRow_2 = xlSheet_2.Application.WorksheetFunction.CountA(xlSheet_2.range("A:A"))
    '遍历第2行
    For i = 1 To xlSheet_2.Columns.Count
        '查找值为小区名称的列号
        If xlSheet_2.cells(2, i).Value = "小区名称" Then
            ColNum = i
            Exit For
        End If
    Next i
    Col = Chr(64 + ColNum)
    colData_2 = xlSheet_2.range(Col & "3:" & Col & lastRow_2).Value2
        
    If IsEmpty(colData_2) Then
        MsgBox ("未读取到小区名称或台账数据为空")
        Exit Sub
    End If
        
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10

    Form1.ListView1.ColumnHeaders.Add , , "", 300 '给列表中添加列名
    Form1.ListView1.ColumnHeaders.Add , , "行号", 600
    Form1.ListView1.ColumnHeaders.Add , , "地市", 800
    Form1.ListView1.ColumnHeaders.Add , , "分公司", 800
    Form1.ListView1.ColumnHeaders.Add , , "设计阶段", 1000
    Form1.ListView1.ColumnHeaders.Add , , "项目名称", 1800
    Form1.ListView1.ColumnHeaders.Add , , "项目编号", 1600
    Form1.ListView1.ColumnHeaders.Add , , "任务名称", 3000
    Form1.ListView1.ColumnHeaders.Add , , "日期", 800
    Form1.ListView1.ColumnHeaders.Add , , "台账名称", 900
                
    colData = xlsheet.range("F1:F" & lastRow).Value2
    ' 遍历第 6 列的所有单元格，并处理每个单元格的数据
    For Row = 2 To UBound(colData, 1)
        ProgressValue = Row

        ' 对比台账数据
        ' 判断元素 num 是否能找到
        Dim matchIndex As Variant
        num = xlsheet.range("F" & Row).Value
        matchIndex = xlApp_2.Application.Match(num, colData_2, 0)
        If Not IsError(matchIndex) Then
            xlsheet.range("H" & Row) = "相同"
        Else
            xlsheet.range("H" & Row) = "不相同"
            X = Form1.ListView1.ListItems.Count + 1
            Form1.ListView1.ListItems.Add , , X
            Form1.ListView1.ListItems(X).SubItems(1) = Row
            Form1.ListView1.ListItems(X).SubItems(2) = xlsheet.range("A" & Row)
            Form1.ListView1.ListItems(X).SubItems(3) = xlsheet.range("B" & Row)
            Form1.ListView1.ListItems(X).SubItems(4) = xlsheet.range("C" & Row)
            Form1.ListView1.ListItems(X).SubItems(5) = xlsheet.range("D" & Row)
            Form1.ListView1.ListItems(X).SubItems(6) = xlsheet.range("E" & Row)
            Form1.ListView1.ListItems(X).SubItems(7) = xlsheet.range("F" & Row)
            Form1.ListView1.ListItems(X).SubItems(8) = xlsheet.range("G" & Row)
            Form1.ListView1.ListItems(X).SubItems(9) = xlsheet.range("H" & Row)
        End If
                    
        ' 更新进度条
        Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 90 / lastRow
        Form1.Label3.Caption = "运行进度：正在拆分并对比数据，请稍等......" & Form1.ProgressBar1.Value & "%"
                           
    Next Row
        
    Form1.Label3.Caption = "运行进度：对比完成 100%"
    MsgBox ("对比完成")
        
    ' 关闭工作簿和Excel应用程序对象
    xlBook_2.Close
    xlApp_2.Quit

    ' 释放对Excel对象的引用
    Set xlBook_2 = Nothing
    Set xlApp_2 = Nothing
End Sub


