Attribute VB_Name = "��ʼ����"
Sub ��ʼ����()
    Dim sheetName_JiaKuan As String, sheetName_ZhuanXian As String, sheetName_GongJian As String
    Dim sheetName_YiBan As String, sheetName_DaiBan As String, sheetName_LiXiang As String, sheetName_SheJi As String
    Dim sheetName_JinDu As String, sheetName As String, sheetName_1 As String
    sheetName_JiaKuan = "*�ҿ�̨��*"
    sheetName_ZhuanXian = "*ר��̨��20*"
    sheetName_GongJian = "������Ϣ��ϸ-����*"
    sheetName_YiBan = "*�Ѱ���Ϣ*"
    sheetName_DaiBan = "*������Ϣ*"
    sheetName_LiXiang = "*��������*"
    sheetName_SheJi = "��ƹ�ģ��ϸ*"
    sheetName = "PSM���̽���"
    sheetName_JinDu = "������Ϣ״̬"
    sheetName_1 = "*�ֹ�˾"
    Dim lastRow_3_1 As Integer, lastRow_3_2 As Integer, lastRow_3_3 As Integer, lastRow_3_4 As Integer
    Dim lastRow_3_6 As Integer, lastRow_3_7 As Integer, lastRow_3_8 As Integer
    
    

    ' ��ʼ��������
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "���н��ȣ����ڶ�ȡPSMϵͳ�������ݣ����Ե�......" & "0%"
    
    Set xlApp_3 = CreateObject("ket.Application") '���ϲ�ѯ˵et.Application��ket.ApplicationҲ��
    If sourceFilePath_3 = "1" Then
        MsgBox ("�뵼��PSMϵͳ��������")
        Exit Sub
    ElseIf sourceFilePath_3 <> "" Then
        Set xlBook_3 = xlApp_3.Workbooks.Open(sourceFilePath_3)  '��ָ��·��ָ�������ļ�
    End If
    ' ����ʾWPS����
    xlApp_3.Visible = False
    
    Form1.ListView1.ListItems.Clear               '����б�
    Form1.ListView1.ColumnHeaders.Clear           '����б�ͷ
    Form1.ListView1.View = lvwReport              '�����б���ʾ��ʽ
    
    ' ��ȡPSM���̽��ȹ�����
    Set xlsheet_3_1 = xlBook_3.Worksheets(sheetName)
    
    ' ��ȡ��Ŀ����
    ' �ҵ����һ��
    lastRow_3_1 = xlsheet_3_1.Application.WorksheetFunction.CountA(xlsheet_3_1.range("A:A"))
    Dim lastColumn_3_1 As Integer
    Dim FindColumnNumber_1_1 As Integer, FindColumnNumber_1_2 As Integer, FindColumnNumber_1_3 As Integer, FindColumnNumber_1_4 As Integer, FindColumnNumber_1_5 As Integer, FindColumnNumber_1_6 As Integer
    
    lastColumn_3_1 = xlsheet_3_1.UsedRange.Columns.Count
    For Column = 1 To lastColumn_3_1
        k = xlsheet_3_1.cells(3, Column).Value
        If k = "��Ŀ����" Then
            FindColumnNumber_1_1 = Column
        End If
        If k = "רҵ����" Then
            FindColumnNumber_1_2 = Column
        End If
        If k = "���Ͷ�ʻ���      ����Ԫ��" Then
            FindColumnNumber_1_3 = Column
        End If
        If k = "��������������ܣ�����" Then
            FindColumnNumber_1_4 = Column
        End If
        If k = "��Ʊ�����ɻ��ܣ�����" Then
            FindColumnNumber_1_5 = Column
        End If
        If k = "��Ʊ���δ��ɻ��ܣ�����" Then
            FindColumnNumber_1_6 = Column
        End If
    Next Column
    
    ' ��ʼ�������ֵ䣬����ҿ��ר����Ŀ����
    Set JiaKuan_TaiZhang = CreateObject("Scripting.Dictionary")
    Set JiaKuan_YiBan = CreateObject("Scripting.Dictionary")
    Set JiaKuan_DaiBan = CreateObject("Scripting.Dictionary")
    Set JiaKuan_ZiJin = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_TaiZhang = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_YiBan = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_DaiBan = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_ZiJin = CreateObject("Scripting.Dictionary")
    
    Project_names = xlsheet_3_1.range(Chr(FindColumnNumber_1_1 + 64) & "1:" & Chr(FindColumnNumber_1_1 + 64) & lastRow_3_1).Value ' ��������ݴ洢��������
    Subsidiary_companys = xlsheet_3_1.range(Chr(FindColumnNumber_1_2 + 64) & "1:" & Chr(FindColumnNumber_1_2 + 64) & lastRow_3_1).Value
    For i = 2 To lastRow_3_1
        Project_name = Project_names(i, 1)
        Subsidiary_company = Subsidiary_companys(i, 1)
        If Subsidiary_company = "�ҿ�" And Project_name Like sheetName_1 Then
            JiaKuan_TaiZhang.Add Project_name, 0
            JiaKuan_YiBan.Add Project_name, 0
            JiaKuan_DaiBan.Add Project_name, 0
            JiaKuan_ZiJin.Add Project_name, 0
        ElseIf Subsidiary_company = "ר��" And Project_name Like sheetName_1 Then
            ZhuanXian_TaiZhang.Add Project_name, 0
            ZhuanXian_YiBan.Add Project_name, 0
            ZhuanXian_DaiBan.Add Project_name, 0
            ZhuanXian_ZiJin.Add Project_name, 0
        End If
    Next i
    
    ' ����������������������������������������������������ҿ�̨�ˡ�������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ����ڴ���ҿ�̨�����ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"

    ' ��ȡ�ҿ�̨�˹�����
    For Each i In xlBook_3.Worksheets
        If i.Name Like sheetName_JiaKuan Then
            Dim rangeString As String
            
            Set xlSheet_3_2 = xlBook_3.Worksheets(i.Name)
            ' �ҵ����һ��
            lastRow_3_2 = xlSheet_3_2.Application.WorksheetFunction.CountA(xlSheet_3_2.range("B:B"))

            ' ͳ�Ƶ�JX��ÿ����Ŀ�����������������ֵ�ֵ
            rangeString = "JX1:JX" & lastRow_3_2
            For Each k In JiaKuan_TaiZhang.keys
                If Len(k) = 0 Then
                    Exit For
                Else
                    JiaKuan_TaiZhang(k) = xlSheet_3_2.Evaluate("=COUNTIF(" & rangeString & "," & Chr(34) & k & Chr(34) & ")")
                End If
            Next k
        End If
    Next i
    
    ' ������������������������������������������������������ר��̨�ˡ�������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ����ڴ���ר��̨�����ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"

    ' ��ȡר��̨�˹�����
    For Each i In xlBook_3.Worksheets
        If i.Name Like sheetName_ZhuanXian Then
            Set xlSheet_3_3 = xlBook_3.Worksheets(i.Name)
            ' �ҵ����һ��
            lastRow_3_3 = xlSheet_3_3.Application.WorksheetFunction.CountA(xlSheet_3_3.range("A:A"))

            ' ͳ�Ƶ�A��ÿ����Ŀ�����������������ֵ�ֵ
            rangeString = "A1:A" & lastRow_3_3

            For Each k In ZhuanXian_TaiZhang.keys
                If Len(k) = 0 Then
                    Exit For
                Else
                    ZhuanXian_TaiZhang(k) = xlSheet_3_3.Evaluate("=COUNTIF(" & rangeString & "," & Chr(34) & k & Chr(34) & ")")
                End If
            Next k
        End If
    Next i
    
    ' �������������������������������������������������������Ѱ���Ϣ��������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ����ڴ����Ѱ���Ϣ���ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"
    
    Dim xlsheet_3_4 As Object

    split sheetName_YiBan, xlBook_3, JiaKuan_YiBan, ZhuanXian_YiBan, xlsheet_3_4
           
    ' �����������������������������������������������������������Ϣ��������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ����ڴ��������Ϣ���ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"
    
    Dim xlsheet_3_5 As Object

    split sheetName_DaiBan, xlBook_3, JiaKuan_DaiBan, ZhuanXian_DaiBan, xlsheet_3_5
    
    ' ������������������������������������������������������������Ϣ��ϸ������������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ����ڴ���������Ϣ��ϸ���ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"

    ' ��ȡ������
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_GongJian Then
            Set xlsheet_3_6 = xlBook_3.Worksheets(SH.Name)
            
            ' �ҵ����һ��
            lastRow_3_6 = xlsheet_3_6.Application.WorksheetFunction.CountA(xlsheet_3_6.range("G:G"))
            
            ' �鿴������� B ���Ƿ���Ƭ����һ�У����û������B��ǰ���� 1 �пհ���
            If xlsheet_3_6.range("A1").Value <> "״̬" Then
                xlsheet_3_6.range("A1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("A1") = "״̬"
            End If
            If xlsheet_3_6.range("B1").Value <> "רҵ����" Then
                xlsheet_3_6.range("B1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("B1") = "רҵ����"
            End If
            If xlsheet_3_6.range("C1").Value <> "��������" Then
                xlsheet_3_6.range("C1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("C1") = "��������"
            End If
            If xlsheet_3_6.range("D1").Value <> "Ƭ��" Then
                xlsheet_3_6.range("D1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("D1") = "Ƭ��"
            End If
            If xlsheet_3_6.range("E1").Value <> "�ֹ�˾" Then
                xlsheet_3_6.range("E1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("E1") = "�ֹ�˾"
            End If
            
            Dim lastColumn_3_6 As Integer
            Dim FindColumnNumber_1 As Integer, FindColumnNumber_2 As Integer, FindColumnNumber_3 As Integer, FindColumnNumber_4 As Integer
    
            lastColumn_3_6 = xlsheet_3_6.UsedRange.Columns.Count
    
            For Column = 1 To lastColumn_3_6
                k = xlsheet_3_6.cells(1, Column).Value
                If k = "��������" Then
                    FindColumnNumber_1 = Column
                End If
                If k = "���񴴽�ʱ��" Then
                    FindColumnNumber_2 = Column
                End If
                If k = "��Ŀ����" Then
                    FindColumnNumber_3 = Column
                End If
                If k = "��Ʊ������ʱ��" Then
                    FindColumnNumber_4 = Column
                End If
            Next Column
            
            Task_names_6 = xlsheet_3_6.range(Chr(FindColumnNumber_1 + 64) & "1:" & Chr(FindColumnNumber_1 + 64) & lastRow_3_6).Value ' ��������ݴ洢��������
            Creation_times_6 = xlsheet_3_6.range(Chr(FindColumnNumber_2 + 64) & "1:" & Chr(FindColumnNumber_2 + 64) & lastRow_3_6).Value
            Project_names_6 = xlsheet_3_6.range(Chr(FindColumnNumber_3 + 64) & "1:" & Chr(FindColumnNumber_3 + 64) & lastRow_3_6).Value
            'Completion_times_6 = xlsheet_3_6.range(Chr(FindColumnNumber_4 + 64) & "1:" & Chr(FindColumnNumber_4 + 64) & lastRow_3_6).Value
            
            ' ����ÿ����Ԫ�񣬲�����ͬ��Ԫ����ӵ��ֵ���
            For i = 2 To lastRow_3_6
                Project_name_6 = Project_names_6(i, 1)
                '��д��������
                If Project_name_6 Like "*����ר��*" Then
                    xlsheet_3_6.range("C" & i) = "����ר��"
                ElseIf Project_name_6 Like "*��ͥ���*" Then
                    xlsheet_3_6.range("C" & i) = "��ͥ���"
                ElseIf Project_name_6 Like "*��ҵ���*" Then
                    xlsheet_3_6.range("C" & i) = "��ҵ���"
                ElseIf Project_name_6 Like "*�̿���Ҫ�ͻ�Ԥ����*" Then
                    xlsheet_3_6.range("C" & i) = "Ԥ����"
                End If
                    
                '��дרҵ����
                a = xlsheet_3_6.range("C" & i)
                If a = "����ר��" Or a = "Ԥ����" Then
                    xlsheet_3_6.range("B" & i) = "ר��"
                Else
                    xlsheet_3_6.range("B" & i) = "�ҿ�"
                End If
                
                '��дƬ��
                If Project_name_6 Like "*����*" Or Project_name_6 Like "*�ϴ�*" Or Project_name_6 Like "*ͭ��*" Or Project_name_6 Like "*����*" Then
                    xlsheet_3_6.cells(i, 4).Value = "����Ƭ��"
                    If Project_name_6 Like "*����*" Then
                        xlsheet_3_6.cells(i, 5).Value = "����"
                    ElseIf Project_name_6 Like "*�ϴ�*" Then
                        xlsheet_3_6.cells(i, 5).Value = "�ϴ�"
                    ElseIf Project_name_6 Like "*ͭ��*" Then
                        xlsheet_3_6.cells(i, 5).Value = "ͭ��"
                    ElseIf Project_name_6 Like "*����*" Then
                        xlsheet_3_6.cells(i, 5).Value = "����"
                    End If
                Else
                    xlsheet_3_6.cells(i, 4).Value = "����Ƭ��"
                    If Project_name_6 Like "*�ɽ*" Then
                        xlsheet_3_6.cells(i, 5).Value = "�ɽ"
                    ElseIf Project_name_6 Like "*����*" Then
                        xlsheet_3_6.cells(i, 5).Value = "����"
                    ElseIf Project_name_6 Like "*����*" Then
                        xlsheet_3_6.cells(i, 5).Value = "����"
                    ElseIf Project_name_6 Like "*�ٲ�*" Then
                        xlsheet_3_6.cells(i, 5).Value = "�ٲ�"
                    ElseIf Project_name_6 Like "*����*" Then
                        xlsheet_3_6.cells(i, 5).Value = "����"
                    End If
                End If
            Next i
    
            ' �����ֵ����洢��ͬ��Ԫ��
            Set distinctElements = CreateObject("Scripting.Dictionary")
            Set Days = CreateObject("Scripting.Dictionary")

            ' ����ÿ����Ԫ�񣬲�����ͬ��Ԫ����ӵ��ֵ���
            For i = lastRow_3_6 To 2 Step -1
                distinctElements(Creation_times_6(i, 1)) = Empty
            Next i
            
            For Each element In distinctElements.keys
                endIdx = InStr(Mid$(element, 6, 5), "/")
                Days(Mid$(element, 1, endIdx + 4)) = Empty
            Next element
            
            ' ���ֵ��еļ�ֵ�Ը��Ƶ�һ��������
            Dim keys() As Variant
            Dim values() As Variant
    
            ReDim keys(0 To Days.Count - 1)
            ReDim values(0 To Days.Count - 1)
    
            i = 0
            For Each element In Days.keys
                keys(i) = element
                values(i) = CInt(Mid$(element, 3, 2)) * 100 + CInt(Mid$(element, 6, 2))
                i = i + 1
            Next element
            
            For i = LBound(values) To UBound(values) - 1
                For j = i + 1 To UBound(values)
                    If values(i) > values(j) Then
                        tempValue = values(i)
                        values(i) = values(j)
                        values(j) = tempValue
                
                        tempKey = keys(i)
                        keys(i) = keys(j)
                        keys(j) = tempKey
                    End If
                Next j
            Next i
            
            Dim GongJian() As Variant
            Dim GongJian_Uncompleted() As Variant
            Dim JiYao() As Variant
            Dim ZhuanZi() As Variant
            Dim ShenHe() As Variant
            Dim Processing() As Variant
            Dim Undistributed() As Variant
            Dim Completed() As Variant
            ReDim GongJian(0 To Days.Count - 1)
            ReDim GongJian_Uncompleted(0 To Days.Count - 1)
            ReDim JiYao(0 To Days.Count - 1)
            ReDim ShenHe(0 To Days.Count - 1)
            ReDim ZhuanZi(0 To Days.Count - 1)
            ReDim Processing(0 To Days.Count - 1)
            ReDim Undistributed(0 To Days.Count - 1)
            ReDim Completed(0 To Days.Count - 1)
            For i = 1 To UBound(values) + 1
                Set GongJian(i - 1) = CreateObject("Scripting.Dictionary")
                Set GongJian_Uncompleted(i - 1) = CreateObject("Scripting.Dictionary")
                Set JiYao(i - 1) = CreateObject("Scripting.Dictionary")
                Set ShenHe(i - 1) = CreateObject("Scripting.Dictionary")
                Set ZhuanZi(i - 1) = CreateObject("Scripting.Dictionary")
                Set Processing(i - 1) = CreateObject("Scripting.Dictionary")
                Set Undistributed(i - 1) = CreateObject("Scripting.Dictionary")
                Set Completed(i - 1) = CreateObject("Scripting.Dictionary")
                For j = 2 To lastRow_3_1
                    Project_name = Project_names(j, 1)
                    If Project_name Like sheetName_1 Then
                        GongJian(i - 1).Add Project_name, 0
                        GongJian_Uncompleted(i - 1).Add Project_name, 0
                        JiYao(i - 1).Add Project_name, 0
                        ZhuanZi(i - 1).Add Project_name, 0
                        ShenHe(i - 1).Add Project_name, 0
                        Processing(i - 1).Add Project_name, 0
                        Undistributed(i - 1).Add Project_name, 0
                        Completed(i - 1).Add Project_name, 0
                    End If
                Next j
            Next i
            
            
            ' �ҵ����һ��
            lastRow_3_5 = xlsheet_3_5.Application.WorksheetFunction.CountA(xlsheet_3_5.range("H:H"))
            Dim data_DaiBan() As Variant
            data_DaiBan = xlsheet_3_5.range("H1:H" & lastRow_3_5).Value  ' ��������ݴ洢��������
            
            ' �ҵ����һ��
            lastRow_3_4 = xlsheet_3_4.Application.WorksheetFunction.CountA(xlsheet_3_4.range("H:H"))
            
            Dim data_YiBan() As Variant
            data_YiBan = xlsheet_3_4.range("H1:H" & lastRow_3_4).Value  ' ��������ݴ洢��������

            Dim foundIndex As Variant
            
            
            ' ͳ�ƹ������������������������
            For i = 2 To lastRow_3_6
                For j = 1 To UBound(values) + 1
                    Task_name_6 = Task_names_6(i, 1)
                    Creation_time_6 = Creation_times_6(i, 1)
                    'Completion_time_6 = Completion_times_6(i, 1)
                    Project_name_6 = Project_names_6(i, 1)
                    If Creation_time_6 Like keys(j - 1) & "*" Then
                        If Task_name_6 Like "*ת��*" Then
                            ZhuanZi(j - 1)(Project_name_6) = ZhuanZi(j - 1)(Project_name_6) + 1
                            xlsheet_3_6.cells(i, 1).Value = "ת��"
                        Else
                        GongJian(j - 1)(Project_name_6) = GongJian(j - 1)(Project_name_6) + 1
                        foundIndex_DaiBan = xlsheet_3_6.Application.Match(Task_name_6, xlsheet_3_6.Application.Index(data_DaiBan, 0, 1), 0)
                        If Len(xlsheet_3_6.cells(i, FindColumnNumber_4).Value) = 0 Then
                            GongJian_Uncompleted(j - 1)(Project_name_6) = GongJian_Uncompleted(j - 1)(Project_name_6) + 1
                            foundIndex_YiBan = xlsheet_3_6.Application.Match(Task_name_6, xlsheet_3_6.Application.Index(data_YiBan, 0, 1), 0)
                            If IsNumeric(foundIndex_DaiBan) Then
                                Processing(j - 1)(Project_name_6) = Processing(j - 1)(Project_name_6) + 1
                                xlsheet_3_6.cells(i, 1).Value = "��Ʊ�����"
                            Else
                                If IsNumeric(foundIndex_YiBan) Then
                                    ShenHe(j - 1)(Project_name_6) = ShenHe(j - 1)(Project_name_6) + 1
                                    xlsheet_3_6.cells(i, 1).Value = "��Ŀ���������"
                                Else
                                    Undistributed(j - 1)(Project_name_6) = Undistributed(j - 1)(Project_name_6) + 1
                                    xlsheet_3_6.cells(i, 1).Value = "δ�ɷ���ƹ���"
                                End If
                            End If
                        Else
                            If IsNumeric(foundIndex_DaiBan) Then
                                Processing(j - 1)(Project_name_6) = Processing(j - 1)(Project_name_6) + 1
                                xlsheet_3_6.cells(i, 1).Value = "��Ʊ�����"
                            Else
                                Completed(j - 1)(Project_name_6) = Completed(j - 1)(Project_name_6) + 1
                                xlsheet_3_6.cells(i, 1).Value = "���̽���"
                            End If
                        End If
                        End If
                    End If
                Next j
            Next i
        End If
    Next SH
    
    ' ��������������������������������������������������������������������������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 20
    Form1.Label3.Caption = "���н��ȣ����ڴ��������������ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"

    ' ��ȡ������
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_LiXiang Then

            Set xlsheet_3_7 = xlBook_3.Worksheets(SH.Name)
            ' �ҵ����һ��
            lastRow_3_7 = xlsheet_3_7.Application.WorksheetFunction.CountA(xlsheet_3_7.range("A:A"))
            
            Dim lastColumn_3_7 As Integer, FindColumnNumber_7_1 As Integer, FindColumnNumber_7_2 As Integer
            'lastColumn_3_7 = xlsheet_3_7.UsedRange.Columns.Count
            lastColumn_3_7 = 20
            Biao_tou_7 = xlsheet_3_7.range("A2:" & Chr(lastColumn_3_7 + 64) & "2").Value
            For Column = 1 To lastColumn_3_7
                If Biao_tou_7(1, Column) = "��Ŀ����" Then
                    FindColumnNumber_7_1 = Column
                End If
                If Biao_tou_7(1, Column) = "�ϼ������������" Then
                    FindColumnNumber_7_2 = Column
                End If
            Next Column
            Project_names_7 = xlsheet_3_7.range(Chr(FindColumnNumber_7_1 + 64) & "1:" & Chr(FindColumnNumber_7_1 + 64) & lastRow_3_7).Value ' ��������ݴ洢��������
            ZiJins = xlsheet_3_7.range(Chr(FindColumnNumber_7_2 + 64) & "1:" & Chr(FindColumnNumber_7_2 + 64) & lastRow_3_7).Value

            ' ͳ�Ƶ�D��ÿ����Ŀ�����������������ֵ�ֵ
            For i = 4 To lastRow_3_7
                If JiaKuan_ZiJin.Exists(Project_names_7(i, 1)) Then
                    JiaKuan_ZiJin(Project_names_7(i, 1)) = ZiJins(i, 1)
                ElseIf ZhuanXian_ZiJin.Exists(Project_names_7(i, 1)) Then
                    ZhuanXian_ZiJin(Project_names_7(i, 1)) = ZiJins(i, 1)
                End If
            Next i
        End If
    Next SH
    
    ' ��������������������������������������������������������ƹ�ģ��ϸ������������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ����ڴ�����ƹ�ģ��ϸ���ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"
    
    ' ��ȡ������
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_SheJi Then
            Set xlsheet_3_8 = xlBook_3.Worksheets(SH.Name)
            
            ' �ҵ����һ��
            lastRow_3_8 = xlsheet_3_8.Application.WorksheetFunction.CountA(xlsheet_3_8.range("A:A"))
            
            Dim lastColumn_3_8 As Integer, FindColumnNumber_8_1 As Integer, FindColumnNumber_8_2 As Integer
            lastColumn_3_8 = xlsheet_3_8.UsedRange.Columns.Count
            Biao_tou_8 = xlsheet_3_8.range("A2:" & Chr(lastColumn_3_8 + 64) & "2").Value
            For Column = 1 To lastColumn_3_8
                If Biao_tou_8(1, Column) = "��Ŀ����" Then
                    FindColumnNumber_8_3 = Column
                End If
                If Biao_tou_8(1, Column) = "�·�" Then
                    FindColumnNumber_8_1 = Column
                End If
                If Biao_tou_8(1, Column) = "�ۺ��˰����Ͷ��(Ԫ��" Then
                    FindColumnNumber_8_2 = Column
                End If
            Next Column
            
            Project_names_8 = xlsheet_3_8.range(Chr(FindColumnNumber_8_3 + 64) & "1:" & Chr(FindColumnNumber_8_3 + 64) & lastRow_3_8).Value
            months = xlsheet_3_8.range(Chr(FindColumnNumber_8_1 + 64) & "1:" & Chr(FindColumnNumber_8_1 + 64) & lastRow_3_8).Value ' ��������ݴ洢��������
            Total_investments = xlsheet_3_8.range(Chr(FindColumnNumber_8_2 + 64) & "1:" & Chr(FindColumnNumber_8_2 + 64) & lastRow_3_8).Value
            For i = 1 To lastRow_3_8
                For j = 1 To UBound(values) + 1
                    If months(i, 1) Like keys(j - 1) & "*" Then
                        JiYao(j - 1)(Project_names_8(i, 1)) = JiYao(j - 1)(Project_names_8(i, 1)) + Total_investments(i, 1)
                    End If
                Next j
            Next i
        End If
    Next SH
    
    ' ��������������������������������������������������������д������������������������������������������������������������
    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ�������д���ݣ����Ե�......" & Form1.ProgressBar1.Value & "%"
    
    Dim copyRange As Object
    Dim pasteRange As Object
    Dim flag As Long
    
    For j = 4 To lastRow_3_1
        If Project_names(j, 1) Like sheetName_1 Then
            xlsheet_3_1.cells(j, FindColumnNumber_1_3).Value = 0
            xlsheet_3_1.cells(j, FindColumnNumber_1_4).Value = 0
            xlsheet_3_1.cells(j, FindColumnNumber_1_6).Value = 0
            xlsheet_3_1.cells(j, FindColumnNumber_1_5).Value = 0
        End If
    Next j
    
    FindColumnNumber_1 = 11
    For i = 0 To UBound(values)
        If keys(i) = "2023/6" Then
            For j = 2 To lastRow_3_1
                Project_name = Project_names(j, 1)
                GongJian(i + 1)(Project_name) = GongJian(i + 1)(Project_name) + GongJian(i)(Project_name)
                JiYao(i + 1)(Project_name) = JiYao(i + 1)(Project_name) + JiYao(i)(Project_name)
            Next j
        Else
            lastColumn_3_1 = xlsheet_3_1.UsedRange.Columns.Count
            For Column = lastColumn_3_1 To 1 Step -1
                a = "*" & Format(keys(i), "yyyy��m��") & "*"
                If xlsheet_3_1.cells(3, Column).Value Like a Then
                    flag = 1
                    Exit For
                End If
            Next Column
            If flag = 1 Then
                FindColumnNumber_1 = Column
                flag = 0
                For j = 4 To lastRow_3_1
                    Project_name = Project_names(j, 1)
                    If Project_name Like sheetName_1 Then
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 4).Value = GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1).Value = JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 2).Value = GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 3).Value = GongJian(i)(Project_name) - GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 1).Value = ShenHe(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1_3).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_3).Value + JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, FindColumnNumber_1_4).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_4).Value + GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1_6).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_6).Value + GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1_5).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_5).Value + GongJian(i)(Project_name) - GongJian_Uncompleted(i)(Project_name)
                    End If
                Next j
        
            ElseIf flag = 0 Then
                xlsheet_3_1.range(Chr(FindColumnNumber_1 + 64) & "3").EntireColumn.Resize(, 5).Insert Shift:=xlToRight
                    
                Set copyRange = xlsheet_3_1.range(xlsheet_3_1.cells(3, FindColumnNumber_1 + 5), xlsheet_3_1.cells(lastRow_3_1, FindColumnNumber_1 + 5)) ' Ҫ���к͸��Ƶ�����Χ
                Set pasteRange = xlsheet_3_1.range(xlsheet_3_1.cells(3, FindColumnNumber_1), xlsheet_3_1.cells(lastRow_3_1, FindColumnNumber_1)) ' Ҫճ����λ�÷�Χ����11����ͬλ�ã�
                ' ���С�ճ����ɾ������
                copyRange.Cut
                pasteRange.PasteSpecial xlPasteValues
                copyRange.ClearContents
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 1).Value = Format(keys(i), "yyyy��m��") & "�������񣨸���"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 2).Value = Format(keys(i), "yyyy��m��") & "��Ʊ�����ɣ�����"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 3).Value = Format(keys(i), "yyyy��m��") & "��Ʊ���δ��ɣ�����"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 4).Value = Format(keys(i), "yyyy��m��") & "��Ŀ��������У�����"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 5).Value = Format(keys(i), "yyyy��m��") & "���Ͷ�ʣ���Ԫ��"
                For j = FindColumnNumber_1_ To FindColumnNumber_1 + 5
                    xlsheet_3_1.cells(8, j).Formula = "=SUM(" & Chr(j + 64) & "4:" & Chr(j + 64) & "7)"
                    xlsheet_3_1.cells(13, j).Formula = "=SUM(" & Chr(j + 64) & "9:" & Chr(j + 64) & "12)"
                    xlsheet_3_1.cells(14, j).Formula = "=" & Chr(j + 64) & "8+" & Chr(j + 64) & "13"
                    xlsheet_3_1.cells(20, j).Formula = "=SUM(" & Chr(j + 64) & "16:" & Chr(j + 64) & "19)"
                    xlsheet_3_1.cells(25, j).Formula = "=SUM(" & Chr(j + 64) & "21:" & Chr(j + 64) & "24)"
                    xlsheet_3_1.cells(26, j).Formula = "=" & Chr(j + 64) & "20+" & Chr(j + 64) & "25"
                    xlsheet_3_1.cells(27, j).Formula = "=" & Chr(j + 64) & "14+" & Chr(j + 64) & "26"
                    xlsheet_3_1.cells(33, j).Formula = "=SUM(" & Chr(j + 64) & "28:" & Chr(j + 64) & "32)"
                    xlsheet_3_1.cells(39, j).Formula = "=SUM(" & Chr(j + 64) & "34:" & Chr(j + 64) & "38)"
                    xlsheet_3_1.cells(40, j).Formula = "=" & Chr(j + 64) & "33+" & Chr(j + 64) & "39"
                    xlsheet_3_1.cells(47, j).Formula = "=SUM(" & Chr(j + 64) & "42:" & Chr(j + 64) & "46)"
                    xlsheet_3_1.cells(53, j).Formula = "=SUM(" & Chr(j + 64) & "48:" & Chr(j + 64) & "52)"
                    xlsheet_3_1.cells(54, j).Formula = "=" & Chr(j + 64) & "47+" & Chr(j + 64) & "53"
                    xlsheet_3_1.cells(55, j).Formula = "=" & Chr(j + 64) & "40+" & Chr(j + 64) & "54"
                    xlsheet_3_1.cells(56, j).Formula = "=" & Chr(j + 64) & "27+" & Chr(j + 64) & "55"
                Next j
                For j = 4 To lastRow_3_1
                    Project_name = Project_names(j, 1)
                    If xlsheet_3_1.cells(j, 3).Value Like sheetName_1 Then
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 1).Value = GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 5).Value = JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 3).Value = GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 2).Value = GongJian(i)(Project_name) - GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 4).Value = ShenHe(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1_3).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_3).Value + JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, FindColumnNumber_1_4).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_4).Value + GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1_6).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_6).Value + GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1_5).Value = xlsheet_3_1.cells(j, FindColumnNumber_1_5).Value + GongJian(i)(Project_name) - GongJian_Uncompleted(i)(Project_name)
                    End If
                Next j
            End If
        End If
    Next i
    
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_JinDu Then
            Set xlsheet_3_9 = xlBook_3.Worksheets(SH.Name)
            
            ' �ҵ����һ��
            lastRow_3_9 = xlsheet_3_9.Application.WorksheetFunction.CountA(xlsheet_3_9.range("A:A"))
            
            Dim lastColumn_3_9 As Integer
            Dim FindColumnNumber_9_1 As Integer
    
            lastColumn_3_9 = xlsheet_3_9.UsedRange.Columns.Count
            For Column = 1 To lastColumn_3_9
                k = xlsheet_3_9.cells(2, Column).Value
                If k = "��Ŀ����" Then
                    FindColumnNumber_9_1 = Column
                End If
            Next Column
            
            xlsheet_3_9.range("F3:J6").Value = 0
            xlsheet_3_9.range("F8:J11").Value = 0
            xlsheet_3_9.range("F14:J17").Value = 0
            xlsheet_3_9.range("F19:J22").Value = 0
            xlsheet_3_9.range("F26:J30").Value = 0
            xlsheet_3_9.range("F32:J36").Value = 0
            xlsheet_3_9.range("F39:J43").Value = 0
            xlsheet_3_9.range("F45:J49").Value = 0
            
            Project_names_9 = xlsheet_3_9.range(Chr(FindColumnNumber_9_1 + 64) & "1:" & Chr(FindColumnNumber_9_1 + 64) & lastRow_3_9).Value ' ��������ݴ洢��������
            For i = 3 To lastRow_3_9
                Project_name_9 = Project_names_9(i, 1)
                If Project_name_9 Like sheetName_1 Then
                    For j = 0 To UBound(values)
                        xlsheet_3_9.cells(i, 6).Value = xlsheet_3_9.cells(i, 6).Value + ZhuanZi(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 7).Value = xlsheet_3_9.cells(i, 7).Value + Undistributed(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 8).Value = xlsheet_3_9.cells(i, 8).Value + Processing(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 9).Value = xlsheet_3_9.cells(i, 9).Value + ShenHe(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 10).Value = xlsheet_3_9.cells(i, 10).Value + Completed(j)(Project_name_9)
                    Next j
                End If
            Next i
        End If
    Next SH
    
    Dim FindColumnNumber_1_7 As Integer, FindColumnNumber_1_8 As Integer, FindColumnNumber_1_9 As Integer, FindColumnNumber_1_10 As Integer
    lastColumn_3_1 = xlsheet_3_1.UsedRange.Columns.Count
    
    For Column = 1 To lastColumn_3_1
        If xlsheet_3_1.cells(3, Column).Value = "��������(��Ԫ��" Then
            FindColumnNumber_1_7 = Column
            Exit For
        End If
    Next Column
    
    For Column = 1 To lastColumn_3_1
        If xlsheet_3_1.cells(3, Column).Value = "����̨��С�����ƣ�����" Then
            FindColumnNumber_1_8 = Column
            Exit For
        End If
    Next Column
    
    For Column = lastColumn_3_1 To 1 Step -1
        If xlsheet_3_1.cells(3, Column).Value = "�Ѱ���Ϣ������" Then
            FindColumnNumber_1_9 = Column
            Exit For
        End If
    Next Column
    
    For Column = lastColumn_3_1 To 1 Step -1
        If xlsheet_3_1.cells(3, Column).Value = "������Ϣ������" Then
            FindColumnNumber_1_10 = Column
            Exit For
        End If
    Next Column

    For i = 4 To lastRow_3_1
        Project_name = Project_names(i, 1)
        Subsidiary_company = Subsidiary_companys(i, 1)
        If Subsidiary_company = "�ҿ�" And Project_name Like sheetName_1 Then
            xlsheet_3_1.cells(i, FindColumnNumber_1_7).Value = JiaKuan_ZiJin(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_1_8).Value = JiaKuan_TaiZhang(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_1_9).Value = JiaKuan_YiBan(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_1_10).Value = JiaKuan_DaiBan(Project_name)
        ElseIf Subsidiary_company = "ר��" And Project_name Like sheetName_1 Then
            xlsheet_3_1.cells(i, FindColumnNumber_1_7).Value = ZhuanXian_ZiJin(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_1_8).Value = ZhuanXian_TaiZhang(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_1_9).Value = ZhuanXian_YiBan(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_1_10).Value = ZhuanXian_DaiBan(Project_name)
        End If
        If Project_name Like sheetName_1 Then
            xlsheet_3_1.cells(i, FindColumnNumber_1_9 - 1).Value = xlsheet_3_1.cells(i, 10).Value - xlsheet_3_1.cells(i, FindColumnNumber_1_9 - 3).Value
            'For j = 0 To UBound(values)
                'xlsheet_3_1.cells(i, FindColumnNumber_1_9 - 1).Value = xlsheet_3_1.cells(i, FindColumnNumber_1_9 - 1).Value + ZhuanZi(j)(Project_name)
            'Next j
        End If
        If Project_name Like sheetName_1 Then
            If xlsheet_3_1.cells(i, FindColumnNumber_1_3 + 1).Value < 0 Then
                Set objRange = xlsheet_3_1.range("H" & i & ":H" & i)
                objRange.Interior.Color = RGB(255, 255, 0)
            Else
                Set objRange = xlsheet_3_1.range("H" & i & ":H" & i)
                objRange.Interior.Color = RGB(255, 255, 255)
            End If
        End If
    Next i
    
    xlBook_3.Save
    xlBook_3.Close SaveChanges:=False
    xlApp_3.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing

    ' ���½�����
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "���н��ȣ���д���" & Form1.ProgressBar1.Value & "%"
    
    MsgBox ("�������")

End Sub
