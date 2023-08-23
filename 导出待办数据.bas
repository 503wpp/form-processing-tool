Attribute VB_Name = "导出待办数据"
Sub 导出待办数据()
    ' 初始化进度条
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "运行进度：正在导出拆分后的待办数据，请稍等......" & "0%"
    
    ' 设置背景颜色
    Set objRange = xlsheet.range("A1:H1")
    objRange.Interior.Color = RGB(192, 192, 192)
    objRange.Font.Name = "宋体"
    objRange.Font.Size = 14
    objRange.Font.Bold = True
    
    ' 选取所有列，并设定自动调整列宽
    xlsheet.Columns.Select
    xlsheet.Columns.AutoFit

    ' 保存工作簿为当前文件路径下的新文件名
    Dim filePath As String
    filePath = xlBook.FullName ' 获取当前文件路径
    Dim newFilePath As String
    newFilePath = Left(filePath, InStrRev(filePath, "\")) & "拆分后待办数据.xlsx" ' 新文件名
    xlBook.SaveCopyAs newFilePath
    
    Form1.ProgressBar1.Value = 100
    Form1.Label3.Caption = "运行进度：导出成功！ 100%"
    MsgBox ("导出成功")

    ' 关闭工作簿和Excel应用程序对象
    'xlBook.Close False
    'xlApp.Quit

    ' 释放对Excel对象的引用
    'Set xlBook = Nothing
    'Set xlApp = Nothing
End Sub
