Attribute VB_Name = "����ȫ������"
Sub ����ȫ������()
    ' ��ʼ��������
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "���н��ȣ����ڵ����ԱȺ��ȫ�����ݣ����Ե�......" & "0%"
    
    ' ���ñ�����ɫ
    Set objRange = xlsheet.range("A1:H1")
    objRange.Interior.Color = RGB(192, 192, 192)
    objRange.Font.Name = "����"
    objRange.Font.Size = 14
    objRange.Font.Bold = True
    
    ' ѡȡ�����У����趨�Զ������п�
    xlsheet.Columns.Select
    xlsheet.Columns.AutoFit
    
    ' ���湤����Ϊ��ǰ�ļ�·���µ����ļ���
    Dim filePath As String
    filePath = xlBook.FullName ' ��ȡ��ǰ�ļ�·��
    Dim newFilePath As String
    newFilePath = Left(filePath, InStrRev(filePath, "\")) & "ȫ���ԱȽ��.xlsx" ' ���ļ���
    xlBook.SaveCopyAs newFilePath
    
    Form1.ProgressBar1.Value = 100
    Form1.Label3.Caption = "���н��ȣ������ɹ��� 100%"
    MsgBox ("�����ɹ�")

    ' �رչ�������ExcelӦ�ó������
    'xlBook.Close False
    'xlApp.Quit

    ' �ͷŶ�Excel���������
    'Set xlBook = Nothing
    'Set xlApp = Nothing
    
End Sub
