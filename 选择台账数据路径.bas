Attribute VB_Name = "选择台账数据路径"
Public sourceFilePath_2 As String
Sub 选择台账数据路径()
    sourceFilePath_2 = vbNull
    Form1.CommonDialog1.Filter = "Excel文件 (*.xls; *.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*"
    Form1.CommonDialog1.DialogTitle = "选择文件"
    Form1.CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    
    ' 显示文件对话框，并获取用户选择的文件名
    Form1.CommonDialog1.ShowOpen
    If Form1.CommonDialog1.FileName <> "" Then
        sourceFilePath_2 = Form1.CommonDialog1.FileName
    End If
End Sub
