Attribute VB_Name = "��ʼ���"
Public xlApp As Variant
Public xlBook As Variant
Public xlsheet As Variant
Public lastRow As Variant

Sub ��ʼ���()
    Dim colData As Variant
    Dim Keyword_1 As String, Keyword_2 As String, Keyword_3 As String, Keyword_4 As String, Keyword_5 As String, Keyword_6 As String, Keyword_7 As String, Keyword_8 As String
    
    Keyword_1 = "����"
    Keyword_2 = "_"
    Keyword_3 = "�ֹ�˾("
    Keyword_4 = ")��"
    Keyword_5 = "�ֹ�˾��"
    Keyword_6 = "����"
    Keyword_7 = "�ֹ�˾"
    Keyword_8 = "��"
    Keyword_9 = "�����"
    Keyword_11 = "-20"
    
    ' �Դ������ݽ��в���
    Set xlApp = CreateObject("ket.Application")
    If sourceFilePath_1 = "1" Then
        MsgBox ("�뵼���������")
        Exit Sub
    ElseIf sourceFilePath_1 <> "" Then
        Set xlBook = xlApp.Workbooks.Open(sourceFilePath_1)  '��ָ��·��ָ�������ļ�
    End If
    ' ����ʾWPS����
    xlApp.Visible = False
    
    Form1.ListView1.ListItems.Clear               '����б�
    Form1.ListView1.ColumnHeaders.Clear           '����б�ͷ
    Form1.ListView1.View = lvwReport              '�����б���ʾ��ʽ
    
    ' ��ʼ��������
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.Label3.Caption = "���н��ȣ����ڴ���������ݣ����Ե�......" & "0%"
    
    ' ָ��Ҫ���ҵ��ֶ����������ڵĹ��������ֵ�һ����
    fieldName = "������Ϣ"
    sheetName = "*" & fieldName & "*"

    ' �� Workbook �в������ְ���ָ���ֶ����Ĺ�����
    For Each i In xlBook.Worksheets
        If i.Name Like sheetName Then
            Set xlsheet = xlBook.Worksheets(i.Name)
            ' �ҵ������ְ���ָ���ֶ����Ĺ�����
            ' �ҵ����һ��
            lastRow = xlsheet.Application.WorksheetFunction.CountA(xlsheet.range("A:A"))
                
            ' �ڹ������ A ��ǰ���� 5 �пհ���
            xlsheet.range("A1").EntireColumn.Resize(, 8).Insert Shift:=xlToRight
            xlsheet.range("A1") = "����"
            xlsheet.range("B1") = "�ֹ�˾"
            xlsheet.range("C1") = "��ƽ׶�"
            xlsheet.range("D1") = "��Ŀ����"
            xlsheet.range("E1") = "��Ŀ���"
            xlsheet.range("F1") = "��������"
            xlsheet.range("G1") = "����"
            xlsheet.range("H1") = "̨������"
 
            colData = xlsheet.range("I1:I" & lastRow).Value2
            ' ������ 6 �е����е�Ԫ�񣬲�����ÿ����Ԫ�������
            For Row = 2 To UBound(colData, 1)
                ProgressValue = Row
                ' ʹ��������ʽ��ȡ��ʼ�ؼ��ʺͽ����ؼ���֮�������
                Set regex = CreateObject("VBScript.RegExp")
                
                ' ��д�ֹ�˾
                regex.Pattern = Keyword_2 & "(.*?)" & Keyword_7
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    xlsheet.range("B" & Row) = Match.SubMatches(0)
                End If
                    
                ' ��д����
                a = xlsheet.range("B" & Row)
                If a = "����" Or a = "�ϴ�" Or a = "ͭ��" Or a = "����" Then
                    xlsheet.range("A" & Row) = "����Ƭ��"
                Else
                    xlsheet.range("A" & Row) = "����Ƭ��"
                End If
                
                ' ��д��ƽ׶�
                regex.Pattern = "��" & "(.*?)" & "��"
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    xlsheet.range("C" & Row) = Match.SubMatches(0)
                End If
                    
                ' ��д��Ŀ����
                regex.Pattern = Keyword_1 & "(.*?)" & Keyword_7
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    xlsheet.range("D" & Row) = Match.SubMatches(0) & "�ֹ�˾"
                End If
                    
                ' ��д��Ŀ���
                If InStr(colData(Row, 1), "�ֹ�˾(") > 0 Then
                    regex.Pattern = Keyword_3 & "(.*?)" & Keyword_4
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        trimmedStr = Match.SubMatches(0)
                        ' ʹ�� Replace ����ȥ���ַ��� trimmedStr �������ߵ�����
                        trimmedStr = Replace(trimmedStr, "(", "")
                        trimmedStr = Replace(trimmedStr, ")", "")
                        xlsheet.range("E" & Row) = trimmedStr
                    End If
                ElseIf InStr(colData(Row, 1), "�ֹ�˾��") > 0 Then
                    regex.Pattern = Keyword_5 & "(.*?)" & Keyword_6
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("E" & Row) = Match.SubMatches(0)
                    End If
                End If
                    
                ' ��д��������
                regex.Pattern = Keyword_8 & "(.*?)" & Keyword_9
                Set matches = regex.Execute(colData(Row, 1))
                If matches.Count > 0 Then
                    Set Match = matches(0)
                    num = Match.SubMatches(0)
                    xlsheet.range("F" & Row) = num
                End If
                
                If xlsheet.range("D" & Row).Value Like "*��ͥ���*" Or xlsheet.range("D" & Row).Value Like "*��ҵ���*" Then
                    ' ��д����
                    regex.Pattern = Keyword_11 & "(.*?)" & Keyword_9
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("G" & Row) = "20" & Match.SubMatches(0)
                    End If
                End If
                
                ' ���½�����
                Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 100 / lastRow
                Form1.Label3.Caption = "���н��ȣ����ڴ���������ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"
                
            Next Row
        End If
        
        Form1.ListView1.ColumnHeaders.Add , , "�����ɣ�", 2000 '���б����������
        
    Next i
    Form1.Label3.Caption = "���н��ȣ������ɣ� 100%"
    MsgBox ("������")
End Sub
