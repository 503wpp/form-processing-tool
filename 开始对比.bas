Attribute VB_Name = "��ʼ�Ա�"
Sub ��ʼ�Ա�()
    
    ' ��ʼ��������
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "���н��ȣ����ڶ�ȡ̨�����ݣ����Ե�......" & "0%"
    
    Set xlApp_2 = CreateObject("ket.Application") '���ϲ�ѯ˵et.Application��ket.ApplicationҲ��
    If sourceFilePath_2 = "1" Then
        MsgBox ("�뵼��̨������")
        Exit Sub
    ElseIf sourceFilePath_2 <> "" Then
        Set xlBook_2 = xlApp_2.Workbooks.Open(sourceFilePath_2)  '��ָ��·��ָ�������ļ�
    End If
    ' ����ʾWPS����
    xlApp_2.Visible = False
    
    
    
    ' ��ʼ���б�
    Form1.ListView1.ListItems.Clear               '����б�
    Form1.ListView1.ColumnHeaders.Clear           '����б�ͷ
    Form1.ListView1.View = lvwReport              '�����б���ʾ��ʽ
    Form1.ListView1.GridLines = True              '��ʾ������
     
    ' ��ȡָ�����ֵĹ�����
    Set xlSheet_2 = xlBook_2.Worksheets("23��̨����ϸ")
    ' �ҵ����һ��
    lastRow_2 = xlSheet_2.Application.WorksheetFunction.CountA(xlSheet_2.range("A:A"))
    '������2��
    For i = 1 To xlSheet_2.Columns.Count
        '����ֵΪС�����Ƶ��к�
        If xlSheet_2.cells(2, i).Value = "С������" Then
            ColNum = i
            Exit For
        End If
    Next i
    Col = Chr(64 + ColNum)
    colData_2 = xlSheet_2.range(Col & "3:" & Col & lastRow_2).Value2
        
    If IsEmpty(colData_2) Then
        MsgBox ("δ��ȡ��С�����ƻ�̨������Ϊ��")
        Exit Sub
    End If
        
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10

    Form1.ListView1.ColumnHeaders.Add , , "", 300 '���б����������
    Form1.ListView1.ColumnHeaders.Add , , "�к�", 600
    Form1.ListView1.ColumnHeaders.Add , , "����", 800
    Form1.ListView1.ColumnHeaders.Add , , "�ֹ�˾", 800
    Form1.ListView1.ColumnHeaders.Add , , "��ƽ׶�", 1000
    Form1.ListView1.ColumnHeaders.Add , , "��Ŀ����", 1800
    Form1.ListView1.ColumnHeaders.Add , , "��Ŀ���", 1600
    Form1.ListView1.ColumnHeaders.Add , , "��������", 3000
    Form1.ListView1.ColumnHeaders.Add , , "����", 800
    Form1.ListView1.ColumnHeaders.Add , , "̨������", 900
                
    colData = xlsheet.range("F1:F" & lastRow).Value2
    ' ������ 6 �е����е�Ԫ�񣬲�����ÿ����Ԫ�������
    For Row = 2 To UBound(colData, 1)
        ProgressValue = Row

        ' �Ա�̨������
        ' �ж�Ԫ�� num �Ƿ����ҵ�
        Dim matchIndex As Variant
        num = xlsheet.range("F" & Row).Value
        matchIndex = xlApp_2.Application.Match(num, colData_2, 0)
        If Not IsError(matchIndex) Then
            xlsheet.range("H" & Row) = "��ͬ"
        Else
            xlsheet.range("H" & Row) = "����ͬ"
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
                    
        ' ���½�����
        Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 90 / lastRow
        Form1.Label3.Caption = "���н��ȣ����ڲ�ֲ��Ա����ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"
                           
    Next Row
        
    Form1.Label3.Caption = "���н��ȣ��Ա���� 100%"
    MsgBox ("�Ա����")
        
    ' �رչ�������ExcelӦ�ó������
    xlBook_2.Close
    xlApp_2.Quit

    ' �ͷŶ�Excel���������
    Set xlBook_2 = Nothing
    Set xlApp_2 = Nothing
End Sub


