Attribute VB_Name = "ѡ��̨������·��"
Public sourceFilePath_2 As String
Sub ѡ��̨������·��()
    sourceFilePath_2 = vbNull
    Form1.CommonDialog1.Filter = "Excel�ļ� (*.xls; *.xlsx)|*.xls;*.xlsx|�����ļ� (*.*)|*.*"
    Form1.CommonDialog1.DialogTitle = "ѡ���ļ�"
    Form1.CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    
    ' ��ʾ�ļ��Ի��򣬲���ȡ�û�ѡ����ļ���
    Form1.CommonDialog1.ShowOpen
    If Form1.CommonDialog1.FileName <> "" Then
        sourceFilePath_2 = Form1.CommonDialog1.FileName
    End If
End Sub
