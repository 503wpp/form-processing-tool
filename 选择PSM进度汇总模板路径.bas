Attribute VB_Name = "ѡ��PSM���Ȼ���ģ��·��"
Public sourceFilePath_3 As String

Sub ѡ��PSM���Ȼ���ģ��·��()
    sourceFilePath_3 = vbNull
    Form1.CommonDialog1.Filter = "Excel�ļ� (*.xls; *.xlsx)|*.xls;*.xlsx|�����ļ� (*.*)|*.*"
    Form1.CommonDialog1.DialogTitle = "ѡ���ļ�"
    Form1.CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNPathMustExist
    
    ' ��ʾ�ļ��Ի��򣬲���ȡ�û�ѡ����ļ���
    Form1.CommonDialog1.ShowOpen
    If Form1.CommonDialog1.FileName <> "" Then
        sourceFilePath_3 = Form1.CommonDialog1.FileName
    End If
End Sub
