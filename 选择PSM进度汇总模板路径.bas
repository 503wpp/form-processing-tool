Attribute VB_Name = "选择PSM进度汇总模板路径"
Public sourceFilePath_3 As String

Sub 选择PSM进度汇总模板路径()
    sourceFilePath_3 = vbNull
    Form1.CommonDialog1.Filter = "Excel文件 (*.xls; *.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*"
    Form1.CommonDialog1.DialogTitle = "选择文件"
    Form1.CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    
    ' 显示文件对话框，并获取用户选择的文件名
    Form1.CommonDialog1.ShowOpen
    If Form1.CommonDialog1.FileName <> "" Then
        sourceFilePath_3 = Form1.CommonDialog1.FileName
    End If
End Sub
