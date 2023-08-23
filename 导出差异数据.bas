Attribute VB_Name = "导出差异数据"
Sub 导出差异数据()
    ' 初始化进度条
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "运行进度：正在导出对比后的不相同部分待办数据，请稍等......" & "0%"
    ' 创建一个新的表格
    Set New_xlBook = xlApp.Workbooks.Add
    
    Set New_xlSheet = New_xlBook.Worksheets("Sheet1")
    
    Set objRange = xlsheet.Rows(1)
    objRange.Copy New_xlSheet.range("A1")
    
    ' 设置背景颜色
    Set objRange = New_xlSheet.range("A1:H1")
    objRange.Interior.Color = RGB(192, 192, 192)
    objRange.Font.Name = "宋体"
    objRange.Font.Size = 14
    objRange.Font.Bold = True
    
    ColNum = 2
    For i = 2 To lastRow Step 1
        If xlsheet.cells(i, 8).Value = "不相同" Then
            Set objRange = xlsheet.Rows(i)
            Col = Chr(64 + ColNum)
            objRange.Copy New_xlSheet.Rows(ColNum)
            ColNum = ColNum + 1
        End If
    Next i
    
    ' 选取所有列，并设定自动调整列宽
    New_xlSheet.Columns.Select
    New_xlSheet.Columns.AutoFit
    
    ' 保存工作簿为当前文件路径下的新文件名
    Dim filePath As String
    filePath = xlBook.FullName ' 获取当前文件路径
    Dim newFilePath As String
    newFilePath = Left(filePath, InStrRev(filePath, "\")) & "差异结果.xlsx" ' 新文件名
    'New_xlBook.SaveCopyAs newFilePath
    New_xlBook.SaveAs newFilePath
    
    Form1.ProgressBar1.Value = 100
    Form1.Label3.Caption = "运行进度：导出成功！ 100%"
    MsgBox ("导出成功")

    ' 关闭工作簿和Excel应用程序对象
    New_xlBook.Close
    'xlApp.Quit

    ' 释放对Excel对象的引用
    Set New_xlBook = Nothing
    'Set xlApp = Nothing
    
End Sub
