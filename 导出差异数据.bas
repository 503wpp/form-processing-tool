Attribute VB_Name = "������������"
Sub ������������()
    ' ��ʼ��������
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "���н��ȣ����ڵ����ԱȺ�Ĳ���ͬ���ִ������ݣ����Ե�......" & "0%"
    ' ����һ���µı��
    Set New_xlBook = xlApp.Workbooks.Add
    
    Set New_xlSheet = New_xlBook.Worksheets("Sheet1")
    
    Set objRange = xlsheet.Rows(1)
    objRange.Copy New_xlSheet.range("A1")
    
    ' ���ñ�����ɫ
    Set objRange = New_xlSheet.range("A1:H1")
    objRange.Interior.Color = RGB(192, 192, 192)
    objRange.Font.Name = "����"
    objRange.Font.Size = 14
    objRange.Font.Bold = True
    
    ColNum = 2
    For i = 2 To lastRow Step 1
        If xlsheet.cells(i, 8).Value = "����ͬ" Then
            Set objRange = xlsheet.Rows(i)
            Col = Chr(64 + ColNum)
            objRange.Copy New_xlSheet.Rows(ColNum)
            ColNum = ColNum + 1
        End If
    Next i
    
    ' ѡȡ�����У����趨�Զ������п�
    New_xlSheet.Columns.Select
    New_xlSheet.Columns.AutoFit
    
    ' ���湤����Ϊ��ǰ�ļ�·���µ����ļ���
    Dim filePath As String
    filePath = xlBook.FullName ' ��ȡ��ǰ�ļ�·��
    Dim newFilePath As String
    newFilePath = Left(filePath, InStrRev(filePath, "\")) & "������.xlsx" ' ���ļ���
    'New_xlBook.SaveCopyAs newFilePath
    New_xlBook.SaveAs newFilePath
    
    Form1.ProgressBar1.Value = 100
    Form1.Label3.Caption = "���н��ȣ������ɹ��� 100%"
    MsgBox ("�����ɹ�")

    ' �رչ�������ExcelӦ�ó������
    New_xlBook.Close
    'xlApp.Quit

    ' �ͷŶ�Excel���������
    Set New_xlBook = Nothing
    'Set xlApp = Nothing
    
End Sub
