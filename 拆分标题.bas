Attribute VB_Name = "��ֱ���"
Sub split(ByVal sheetName As Variant, ByVal xlBook_3 As Variant, ByRef JiaKuan As Variant, ByRef ZhuanXian As Variant, ByRef xlsheet As Variant)
    
    Dim colData As Variant
    Dim Keyword_1 As String, Keyword_2 As String, Keyword_3 As String, Keyword_4 As String, Keyword_5 As String, Keyword_6 As String, Keyword_7 As String, Keyword_8 As String
    Dim lastRow As Integer
    
    Keyword_1 = "��"
    Keyword_2 = "��"
    Keyword_3 = "����"
    Keyword_4 = "�ֹ�˾"
    Keyword_5 = "�ֹ�˾("
    Keyword_6 = ")��"
    Keyword_7 = "�ֹ�˾��"
    Keyword_8 = "����"
    Keyword_9 = "�����"
    Keyword_10 = "_"
    Keyword_11 = "-20"
    
    ' ��ȡ�Ѱ���Ϣ������
    For Each i In xlBook_3.Worksheets
        If i.Name Like sheetName Then
            Set xlsheet = xlBook_3.Worksheets(i.Name)
            
            ' ����ȱ�ٵ���
            If xlsheet.range("A1").Value <> "����" Then
                xlsheet.range("A1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("A1") = "����"
            End If
            If xlsheet.range("B1").Value <> "�ֹ�˾" Then
                xlsheet.range("B1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("B1") = "�ֹ�˾"
            End If
            If xlsheet.range("C1").Value <> "��ƽ׶�" Then
                xlsheet.range("C1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("C1") = "��ƽ׶�"
            End If
            If xlsheet.range("D1").Value <> "��Ŀ����" Then
                xlsheet.range("D1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("D1") = "��Ŀ����"
            End If
            If xlsheet.range("E1").Value <> "��Ŀ���" Then
                xlsheet.range("E1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("E1") = "��Ŀ���"
            End If
            If xlsheet.range("F1").Value <> "��������" Then
                xlsheet.range("F1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("F1") = "��������"
            End If
            If xlsheet.range("G1").Value <> "����" Then
                xlsheet.range("G1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet.range("G1") = "����"
            End If
            
            ' �ҵ����һ��
            lastRow = xlsheet.Application.WorksheetFunction.CountA(xlsheet.range("H:H"))
            
            ' ��ȡ����������
            colData = xlsheet.range("H1:H" & lastRow).Value2
            Dim regex As Object
            ' ���������е����е�Ԫ�񣬲�����ÿ����Ԫ�������
            For Row = 2 To UBound(colData, 1)
                'If Len(xlSheet.Range("A" & Row).Value) = 0 Then
                    ProgressValue = Row
                    ' ʹ��������ʽ��ȡ��ʼ�ؼ��ʺͽ����ؼ���֮�������
                    Set regex = CreateObject("VBScript.RegExp")
                    
                    ' ��д�ֹ�˾
                    regex.Pattern = Keyword_10 & "(.*?)" & Keyword_4
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
                    regex.Pattern = Keyword_1 & "(.*?)" & Keyword_2
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("C" & Row) = Match.SubMatches(0)
                    End If
                
                    ' ��д��Ŀ����
                    regex.Pattern = Keyword_3 & "(.*?)" & Keyword_4
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("D" & Row) = Match.SubMatches(0) & "�ֹ�˾"
                    End If
                    
                    ' ��д��Ŀ���
                    If InStr(colData(Row, 1), "�ֹ�˾(") > 0 Then
                        regex.Pattern = Keyword_5 & "(.*?)" & Keyword_6
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
                        regex.Pattern = Keyword_7 & "(.*?)" & Keyword_8
                        Set matches = regex.Execute(colData(Row, 1))
                        If matches.Count > 0 Then
                            Set Match = matches(0)
                            trimmedStr = Match.SubMatches(0)
                            xlsheet.range("E" & Row) = trimmedStr
                        End If
                    End If
                    
                    ' ��д��������
                    regex.Pattern = "��" & "(.*?)" & Keyword_9
                    Set matches = regex.Execute(colData(Row, 1))
                    If matches.Count > 0 Then
                        Set Match = matches(0)
                        xlsheet.range("F" & Row) = Match.SubMatches(0)
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
                'End If
            Next Row
        End If
    Next i
    
    ' ѡȡ�����У����趨�Զ������п�
    xlsheet.Columns("A:G").AutoFit
    
    Design_distributions = xlsheet.range("C1:C" & lastRow).Value
    For i = lastRow To 2 Step -1
        If Design_distributions(i, 1) = "��ƿ���" Then
            xlsheet.Rows(i).Delete
        End If
    Next i
    
    If sheetName = "*������Ϣ*" Then
        ' ����ÿ����Ԫ�񣬲�����ͬ��Ԫ����ӵ��ֵ���
        For i = lastRow To 2 Step -1
            a = xlsheet.cells(i, 6).Value
            If a Like "*SCM����" Or a Like "*SCM����" Or a Like "*���ķ�̯" Or a Like "*����Ŀ����" Then
                xlsheet.Rows(i).Delete
            End If
        Next i
    End If
            

    ' ͳ�Ƶ�B�мҿ��Ѱ��������������ֵ�ֵ
    rangeString = "D1:D" & lastRow

    For Each k In JiaKuan.keys
        If Len(k) = 0 Then
            Exit For
        Else
            formulaString = "=COUNTIF(" & rangeString & "," & Chr(34) & k & Chr(34) & ")"
            Count = xlsheet.Evaluate(formulaString)
            JiaKuan(k) = Count
        End If
    Next k
    ' ͳ�Ƶ�B��ר���Ѱ��������������ֵ�ֵ
    For Each k In ZhuanXian.keys
        If Len(k) = 0 Then
            Exit For
        Else
            formulaString = "=COUNTIF(" & rangeString & "," & Chr(34) & k & Chr(34) & ")"
            Count = xlsheet.Evaluate(formulaString)
            ZhuanXian(k) = Count
        End If
    Next k
End Sub
