Attribute VB_Name = "开始汇总"
Sub 开始汇总()
    Dim sheetName_JiaKuan As String, sheetName_ZhuanXian As String, sheetName_GongJian As String
    Dim sheetName_YiBan As String, sheetName_DaiBan As String, sheetName_LiXiang As String, sheetName_SheJi As String
    Dim sheetName_JinDu As String, sheetName As String, sheetName_1 As String
    sheetName_JiaKuan = "*家宽台账*"
    sheetName_ZhuanXian = "*专线台账20*"
    sheetName_GongJian = "任务信息明细-渝西*"
    sheetName_YiBan = "*已办信息*"
    sheetName_DaiBan = "*待办信息*"
    sheetName_LiXiang = "*立项批复*"
    sheetName_SheJi = "设计规模明细*"
    sheetName = "PSM流程进度"
    sheetName_JinDu = "任务信息状态"
    sheetName_1 = "*分公司"
    Dim lastRow_3_1 As Integer, lastRow_3_2 As Integer, lastRow_3_3 As Integer, lastRow_3_4 As Integer
    Dim lastRow_3_6 As Integer, lastRow_3_7 As Integer, lastRow_3_8 As Integer
    
    

    ' 初始化进度条
    Form1.ProgressBar1.Min = 0
    Form1.ProgressBar1.Max = 100
    Form1.ProgressBar1.Value = 0
    Form1.Label3.Caption = "运行进度：正在读取PSM系统进度数据，请稍等......" & "0%"
    
    Set xlApp_3 = CreateObject("ket.Application") '网上查询说et.Application及ket.Application也能
    If sourceFilePath_3 = "1" Then
        MsgBox ("请导入PSM系统进度数据")
        Exit Sub
    ElseIf sourceFilePath_3 <> "" Then
        Set xlBook_3 = xlApp_3.Workbooks.Open(sourceFilePath_3)  '打开指定路径指定名称文件
    End If
    ' 不显示WPS界面
    xlApp_3.Visible = True
    
    Form1.ListView1.ListItems.Clear               '清空列表
    Form1.ListView1.ColumnHeaders.Clear           '清空列表头
    Form1.ListView1.View = lvwReport              '设置列表显示方式
    
    ' 获取PSM流程进度工作表
    Set xlsheet_3_1 = xlBook_3.Worksheets(sheetName)
    
    ' 获取项目名称
    ' 找到最后一行
    lastRow_3_1 = xlsheet_3_1.Application.WorksheetFunction.CountA(xlsheet_3_1.range("A:A"))
    
    ' 初始化两个字典，保存家宽和专线项目名称
    Set JiaKuan_TaiZhang = CreateObject("Scripting.Dictionary")
    Set JiaKuan_YiBan = CreateObject("Scripting.Dictionary")
    Set JiaKuan_DaiBan = CreateObject("Scripting.Dictionary")
    Set JiaKuan_ZiJin = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_TaiZhang = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_YiBan = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_DaiBan = CreateObject("Scripting.Dictionary")
    Set ZhuanXian_ZiJin = CreateObject("Scripting.Dictionary")
    
    Project_names = xlsheet_3_1.range("D1:D" & lastRow_3_1).Value ' 将表格数据存储到数组中
    Subsidiary_companys = xlsheet_3_1.range("C1:C" & lastRow_3_1).Value
    For i = 2 To lastRow_3_1
        Project_name = Project_names(i, 1)
        Subsidiary_company = Subsidiary_companys(i, 1)
        If Subsidiary_company = "家宽" And Project_name Like sheetName_1 Then
            JiaKuan_TaiZhang.Add Project_name, 0
            JiaKuan_YiBan.Add Project_name, 0
            JiaKuan_DaiBan.Add Project_name, 0
            JiaKuan_ZiJin.Add Project_name, 0
        ElseIf Subsidiary_company = "专线" And Project_name Like sheetName_1 Then
            ZhuanXian_TaiZhang.Add Project_name, 0
            ZhuanXian_YiBan.Add Project_name, 0
            ZhuanXian_DaiBan.Add Project_name, 0
            ZhuanXian_ZiJin.Add Project_name, 0
        End If
    Next i
    
    ' ————————————————————————处理家宽台账————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在处理家宽台账数据，请稍等......" & Form1.ProgressBar1.Value & "%"

    ' 获取家宽台账工作表
    For Each i In xlBook_3.Worksheets
        If i.Name Like sheetName_JiaKuan Then
            Dim rangeString As String
            
            Set xlSheet_3_2 = xlBook_3.Worksheets(i.Name)
            ' 找到最后一行
            lastRow_3_2 = xlSheet_3_2.Application.WorksheetFunction.CountA(xlSheet_3_2.range("B:B"))

            ' 统计第JX列每个项目名称数量，并更新字典值
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
    
    ' —————————————————————————处理专线台账————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在处理专线台账数据，请稍等......" & Form1.ProgressBar1.Value & "%"

    ' 获取专线台账工作表
    For Each i In xlBook_3.Worksheets
        If i.Name Like sheetName_ZhuanXian Then
            Set xlSheet_3_3 = xlBook_3.Worksheets(i.Name)
            ' 找到最后一行
            lastRow_3_3 = xlSheet_3_3.Application.WorksheetFunction.CountA(xlSheet_3_3.range("A:A"))

            ' 统计第A列每个项目名称数量，并更新字典值
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
    
    ' —————————————————————————处理已办信息————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在处理已办信息数据，请稍等......" & Form1.ProgressBar1.Value & "%"
    
    Dim xlsheet_3_4 As Object

    split sheetName_YiBan, xlBook_3, JiaKuan_YiBan, ZhuanXian_YiBan, xlsheet_3_4
           
    ' —————————————————————————处理待办信息————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在处理待办信息数据，请稍等......" & Form1.ProgressBar1.Value & "%"
    
    Dim xlsheet_3_5 As Object

    split sheetName_DaiBan, xlBook_3, JiaKuan_DaiBan, ZhuanXian_DaiBan, xlsheet_3_5
    
    ' —————————————————————————处理任务信息明细——————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在处理任务信息明细数据，请稍等......" & Form1.ProgressBar1.Value & "%"

    ' 获取工作表
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_GongJian Then
            Set xlsheet_3_6 = xlBook_3.Worksheets(SH.Name)
            
            ' 找到最后一行
            lastRow_3_6 = xlsheet_3_6.Application.WorksheetFunction.CountA(xlsheet_3_6.range("D:D"))
            
            ' 查看工作表的 B 列是否有片区这一列，如果没有则在B列前插入 1 列空白列
            If xlsheet_3_6.range("A1").Value <> "状态" Then
                xlsheet_3_6.range("A1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("A1") = "状态"
            End If
            If xlsheet_3_6.range("B1").Value <> "片区" Then
                xlsheet_3_6.range("B1").EntireColumn.Resize(, 1).Insert Shift:=xlToRight
                xlsheet_3_6.range("B1") = "片区"
            End If
            
            ' 遍历每个单元格，并将不同的元素添加到字典中
            For i = 2 To lastRow_3_6
                If xlsheet_3_6.cells(i, 4).Value Like "*北碚*" Or xlsheet_3_6.cells(i, 4).Value Like "*合川*" Or xlsheet_3_6.cells(i, 4).Value Like "*铜梁*" Or xlsheet_3_6.cells(i, 4).Value Like "*潼南*" Then
                    xlsheet_3_6.cells(i, 2).Value = "北碚片区"
                Else
                    xlsheet_3_6.cells(i, 2).Value = "永川片区"
                End If
            Next i
    
            ' 创建字典来存储不同的元素
            Set distinctElements = CreateObject("Scripting.Dictionary")
            Set Days = CreateObject("Scripting.Dictionary")

            Dim column As Integer, lastColumn_3_6 As Integer
            Dim FindColumnNumber_1 As Integer, FindColumnNumber_2 As Integer, FindColumnNumber_3 As Integer, FindColumnNumber_4 As Integer
    
            lastColumn_3_6 = xlsheet_3_6.UsedRange.Columns.Count
    
            For column = 1 To lastColumn_3_6
                k = xlsheet_3_6.cells(1, column).Value
                If k = "任务名称" Then
                    FindColumnNumber_1 = column
                End If
                If k = "任务创建时间" Then
                    FindColumnNumber_2 = column
                End If
                If k = "项目名称" Then
                    FindColumnNumber_3 = column
                End If
                If k = "设计编制完成时间" Then
                    FindColumnNumber_4 = column
                End If
            Next column
            
            
            Task_names = xlsheet_3_6.range(Chr(FindColumnNumber_1 + 64) & "1:" & Chr(FindColumnNumber_1 + 64) & lastRow_3_6).Value ' 将表格数据存储到数组中
            Creation_times = xlsheet_3_6.range(Chr(FindColumnNumber_2 + 64) & "1:" & Chr(FindColumnNumber_2 + 64) & lastRow_3_6).Value
            Completion_times = xlsheet_3_6.range(Chr(FindColumnNumber_4 + 64) & "1:" & Chr(FindColumnNumber_4 + 64) & lastRow_3_6).Value
            
            ' 遍历每个单元格，并将不同的元素添加到字典中
            For i = lastRow_3_6 To 2 Step -1
                distinctElements(Creation_times(i, 1)) = Empty
            Next i
            
            For Each element In distinctElements.keys
                endIdx = InStr(Mid$(element, 6, 5), "/")
                Days(Mid$(element, 1, endIdx + 4)) = Empty
            Next element
            
            ' 将字典中的键值对复制到一个数组中
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
                    Project_name = Project_names(i, 1)
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
            
            
            ' 找到最后一行
            lastRow_3_5 = xlsheet_3_5.Application.WorksheetFunction.CountA(xlsheet_3_5.range("F:F"))
            Dim data_DaiBan() As Variant
            data_DaiBan = xlsheet_3_5.range("F1:F" & lastRow_3_5).Value  ' 将表格数据存储到数组中
            
            ' 找到最后一行
            lastRow_3_4 = xlsheet_3_4.Application.WorksheetFunction.CountA(xlsheet_3_4.range("F:F"))
            
            Dim data_YiBan() As Variant
            data_YiBan = xlsheet_3_4.range("F1:F" & lastRow_3_4).Value  ' 将表格数据存储到数组中

            Dim foundIndex As Variant
            
            
            ' 统计工建创建任务数，完成任务数
            For i = 2 To lastRow_3_6
                For j = 1 To UBound(values) + 1
                    Task_name = Task_names(i, 1)
                    Creation_time = Creation_times(i, 1)
                    Completion_time = Completion_times(i, 1)
                    If Creation_time Like keys(j - 1) & "*" Then
                        GongJian(j - 1)(Completion_time) = GongJian(j - 1)(Completion_time) + 1
                        If Task_name Like "*转资*" Then
                            ZhuanZi(j - 1)(Completion_time) = ZhuanZi(j - 1)(Completion_time) + 1
                            xlsheet_3_6.cells(i, 1).Value = "转资"
                        Else
                        foundIndex_DaiBan = xlsheet_3_6.Application.Match(Task_name, xlsheet_3_6.Application.Index(data_DaiBan, 0, 1), 0)
                        If Len(xlsheet_3_6.cells(i, FindColumnNumber_4).Value) = 0 Then
                            GongJian_Uncompleted(j - 1)(Completion_time) = GongJian_Uncompleted(j - 1)(Completion_time) + 1
                            
                            foundIndex_YiBan = xlsheet_3_6.Application.Match(Task_name, xlsheet_3_6.Application.Index(data_YiBan, 0, 1), 0)
                            If IsNumeric(foundIndex_DaiBan) Then
                                Processing(j - 1)(Completion_time) = Processing(j - 1)(Completion_time) + 1
                                xlsheet_3_6.cells(i, 1).Value = "设计编制中"
                            Else
                                If IsNumeric(foundIndex_YiBan) Then
                                    ShenHe(j - 1)(Completion_time) = ShenHe(j - 1)(Completion_time) + 1
                                    xlsheet_3_6.cells(i, 1).Value = "项目经理审核中"
                                Else
                                    Undistributed(j - 1)(Completion_time) = Undistributed(j - 1)(Completion_time) + 1
                                    xlsheet_3_6.cells(i, 1).Value = "未派发设计工单"
                                End If
                            End If
                        Else
                            If IsNumeric(foundIndex_DaiBan) Then
                                Processing(j - 1)(Completion_time) = Processing(j - 1)(Completion_time) + 1
                                xlsheet_3_6.cells(i, 1).Value = "设计编制中"
                            Else
                                Completed(j - 1)(Completion_time) = Completed(j - 1)(Completion_time) + 1
                                xlsheet_3_6.cells(i, 1).Value = "流程结束"
                            End If
                        End If
                        End If
                    End If
                Next j
            Next i
        End If
    Next SH
    
    ' —————————————————————————处理立项批复——————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 20
    Form1.Label3.Caption = "运行进度：正在处理立项批复数据，请稍等......" & Form1.ProgressBar1.Value & "%"

    ' 获取工作表
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_LiXiang Then

            Set xlsheet_3_7 = xlBook_3.Worksheets(SH.Name)
            ' 找到最后一行
            lastRow_3_7 = xlsheet_3_7.Application.WorksheetFunction.CountA(xlsheet_3_7.range("A:A"))
            Project_names_7 = xlsheet_3_7.range("A1:A" & lastRow_3_7).Value ' 将表格数据存储到数组中
            ZiJins = xlsheet_3_7.range("D1:D" & lastRow_3_7).Value

            ' 统计第D列每个项目名称数量，并更新字典值
            For i = 4 To lastRow_3_7
                If JiaKuan_ZiJin.Exists(Project_names_7(i, 1)) Then
                    JiaKuan_ZiJin(Project_names_7(i, 1)) = ZiJins(i, 1)
                ElseIf ZhuanXian_ZiJin.Exists(Project_names_7(i, 1)) Then
                    ZhuanXian_ZiJin(Project_names_7(i, 1)) = ZiJins(i, 1)
                End If
            Next i
        End If
    Next SH
    
    ' —————————————————————————处理设计规模明细——————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在处理设计规模明细数据，请稍等......" & Form1.ProgressBar1.Value & "%"
    
    ' 获取工作表
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_SheJi Then
            Set xlsheet_3_8 = xlBook_3.Worksheets(SH.Name)
            
            ' 找到最后一行
            lastRow_3_8 = xlsheet_3_8.Application.WorksheetFunction.CountA(xlsheet_3_8.range("A:A"))
            
            Dim lastColumn_3_8 As Integer
            lastColumn_3_8 = xlsheet_3_8.UsedRange.Columns.Count
            Biao_tou = xlsheet_3_8.range("A2:" & Chr(lastColumn_3_8 + 64) & "2").Value
    
            For column = 1 To lastColumn_3_8
                If Biao_tou(1, column) = "月份" Then
                    FindColumnNumber_1 = column
                    Exit For
                End If
            Next column
            
            For column = lastColumn_3_8 To 1 Step -1
                If Biao_tou(1, column) = "折后除税价总投资(元）" Then
                    FindColumnNumber_2 = column
                    Exit For
                End If
            Next column
            
            Project_names_8 = xlsheet_3_8.range("A1:A" & lastRow_3_8).Value
            months = xlsheet_3_8.range(Chr(FindColumnNumber_1 + 64) & "1:" & Chr(FindColumnNumber_1 + 64) & lastRow_3_8).Value ' 将表格数据存储到数组中
            Total_investments = xlsheet_3_8.range(Chr(FindColumnNumber_2 + 64) & "1:" & Chr(FindColumnNumber_2 + 64) & lastRow_3_6).Value
            For i = 1 To lastRow_3_8
                For j = 1 To UBound(values) + 1
                    If months(i, 1) Like keys(j - 1) & "*" Then
                        JiYao(j - 1)(Project_names_8(i, 1)) = JiYao(j - 1)(Project_names_8(i, 1)) + Total_investments(i, 1)
                    End If
                Next j
            Next i
        End If
    Next SH
    
    ' —————————————————————————正在填写——————————————————————————————
    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：正在填写数据，请稍等......" & Form1.ProgressBar1.Value & "%"
    
    Dim copyRange As Object
    Dim pasteRange As Object
    Dim flag As Long
    Dim lastColumn_3_1 As Integer
    
    For j = 4 To lastRow_3_1
        If xlsheet_3_1.cells(j, 4).Value Like sheetName_1 Then
            xlsheet_3_1.cells(j, 6).Value = 0
            xlsheet_3_1.cells(j, 9).Value = 0
            xlsheet_3_1.cells(j, 11).Value = 0
            xlsheet_3_1.cells(j, 10).Value = 0
        End If
    Next j
    
    FindColumnNumber_1 = 11
    For i = 0 To UBound(values)
        If keys(i) = "2023/6" Then
            For j = 2 To lastRow_3_1
                GongJian(i + 1)(xlsheet_3_1.cells(j, 4).Value) = GongJian(i + 1)(xlsheet_3_1.cells(j, 4).Value) + GongJian(i)(xlsheet_3_1.cells(j, 4).Value)
                JiYao(i + 1)(xlsheet_3_1.cells(j, 4).Value) = JiYao(i + 1)(xlsheet_3_1.cells(j, 4).Value) + JiYao(i)(xlsheet_3_1.cells(j, 4).Value)
            Next j
        Else
            lastColumn_3_1 = xlsheet_3_1.UsedRange.Columns.Count
            For column = lastColumn_3_1 To 1 Step -1
                a = "*" & Format(keys(i), "yyyy年m月") & "*"
                If xlsheet_3_1.cells(3, column).Value Like a Then
                    flag = 1
                    Exit For
                End If
            Next column
            If flag = 1 Then
                FindColumnNumber_1 = column
                flag = 0
                For j = 4 To lastRow_3_1
                    Project_name = Project_names(j, 1)
                    If Project_name Like sheetName_1 Then
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 4).Value = GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1).Value = JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 2).Value = GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 3).Value = xlsheet_3_1.cells(j, FindColumnNumber_1 - 4).Value - xlsheet_3_1.cells(j, FindColumnNumber_1 - 2).Value
                        xlsheet_3_1.cells(j, FindColumnNumber_1 - 1).Value = ShenHe(i)(Project_name)
                        xlsheet_3_1.cells(j, 6).Value = xlsheet_3_1.cells(j, 6).Value + JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, 9).Value = xlsheet_3_1.cells(j, 9).Value + GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, 11).Value = xlsheet_3_1.cells(j, 11).Value + GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, 10).Value = xlsheet_3_1.cells(j, 10).Value + xlsheet_3_1.cells(j, FindColumnNumber_1 - 3).Value
                    End If
                Next j
        
            ElseIf flag = 0 Then
                xlsheet_3_1.range(Chr(FindColumnNumber_1 + 64) & "3").EntireColumn.Resize(, 5).Insert Shift:=xlToRight
                    
                Set copyRange = xlsheet_3_1.range(xlsheet_3_1.cells(3, FindColumnNumber_1 + 5), xlsheet_3_1.cells(lastRow_3_1, FindColumnNumber_1 + 5)) ' 要剪切和复制的区域范围
                Set pasteRange = xlsheet_3_1.range(xlsheet_3_1.cells(3, FindColumnNumber_1), xlsheet_3_1.cells(lastRow_3_1, FindColumnNumber_1)) ' 要粘贴的位置范围（第11列相同位置）
                ' 剪切、粘贴和删除内容
                copyRange.Cut
                pasteRange.PasteSpecial xlPasteValues
                copyRange.ClearContents
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 1).Value = Format(keys(i), "yyyy年m月") & "创建任务（个）"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 2).Value = Format(keys(i), "yyyy年m月") & "设计编制完成（个）"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 3).Value = Format(keys(i), "yyyy年m月") & "设计编制未完成（个）"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 4).Value = Format(keys(i), "yyyy年m月") & "项目经理审核中（个）"
                xlsheet_3_1.cells(3, FindColumnNumber_1 + 5).Value = Format(keys(i), "yyyy年m月") & "设计投资（万元）"
                For j = FindColumnNumber_1 To FindColumnNumber_1 + 5
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
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 2).Value = xlsheet_3_1.cells(j, FindColumnNumber_1 + 1).Value - xlsheet_3_1.cells(j, FindColumnNumber_1 + 3).Value
                        xlsheet_3_1.cells(j, FindColumnNumber_1 + 4).Value = ShenHe(i)(Project_name)
                        xlsheet_3_1.cells(j, 6).Value = xlsheet_3_1.cells(j, 6).Value + JiYao(i)(Project_name) / 10000
                        xlsheet_3_1.cells(j, 9).Value = xlsheet_3_1.cells(j, 9).Value + GongJian(i)(Project_name)
                        xlsheet_3_1.cells(j, 11).Value = xlsheet_3_1.cells(j, 11).Value + GongJian_Uncompleted(i)(Project_name)
                        xlsheet_3_1.cells(j, 10).Value = xlsheet_3_1.cells(j, 10).Value + xlsheet_3_1.cells(j, FindColumnNumber_1 + 2).Value
                    End If
                Next j
            End If
        End If
    Next i
    
    For Each SH In xlBook_3.Worksheets
        If SH.Name Like sheetName_JinDu Then
            Set xlsheet_3_9 = xlBook_3.Worksheets(SH.Name)
            
            ' 找到最后一行
            lastRow_3_9 = xlsheet_3_9.Application.WorksheetFunction.CountA(xlsheet_3_9.range("A:A"))
            
            xlsheet_3_9.range("E3:I6").Value = 0
            xlsheet_3_9.range("E8:I11").Value = 0
            xlsheet_3_9.range("E14:I17").Value = 0
            xlsheet_3_9.range("E19:I22").Value = 0
            xlsheet_3_9.range("E26:I30").Value = 0
            xlsheet_3_9.range("E32:I36").Value = 0
            xlsheet_3_9.range("E39:I43").Value = 0
            xlsheet_3_9.range("E45:I49").Value = 0
            
            Project_names_9 = xlsheet_3_9.range("D1:D" & lastRow_3_9).Value ' 将表格数据存储到数组中
            For i = 3 To lastRow_3_9
                Project_name_9 = Project_names_9(i, 1)
                If Project_name_9 Like sheetName_1 Then
                    For j = 0 To UBound(values)
                        xlsheet_3_9.cells(i, 5).Value = xlsheet_3_9.cells(i, 5).Value + ZhuanZi(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 6).Value = xlsheet_3_9.cells(i, 6).Value + Undistributed(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 7).Value = xlsheet_3_9.cells(i, 7).Value + Processing(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 8).Value = xlsheet_3_9.cells(i, 8).Value + ShenHe(j)(Project_name_9)
                        xlsheet_3_9.cells(i, 9).Value = xlsheet_3_9.cells(i, 9).Value + Completed(j)(Project_name_9)
                        
                    Next j
                End If
            Next i
        End If
    Next SH
    
    lastColumn_3_1 = xlsheet_3_1.UsedRange.Columns.Count
    
    For column = lastColumn_3_1 To 1 Step -1
        If xlsheet_3_1.cells(3, column).Value = "已办信息（个）" Then
            FindColumnNumber_2 = column
            Exit For
        End If
    Next column
    
    For column = lastColumn_3_1 To 1 Step -1
        If xlsheet_3_1.cells(3, column).Value = "待办信息（个）" Then
            FindColumnNumber_3 = column
            Exit For
        End If
    Next column

    For i = 4 To lastRow_3_1
        Project_name = Project_names(i, 1)
        Subsidiary_company = Subsidiary_companys(i, 1)
        If Subsidiary_company = "家宽" And Project_name Like sheetName_1 Then
            xlsheet_3_1.cells(i, 5).Value = JiaKuan_ZiJin(Project_name)
            xlsheet_3_1.cells(i, 8).Value = JiaKuan_TaiZhang(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_2).Value = JiaKuan_YiBan(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_3).Value = JiaKuan_DaiBan(Project_name)
        ElseIf Subsidiary_company = "专线" And Project_name Like sheetName_1 Then
            xlsheet_3_1.cells(i, 5).Value = ZhuanXian_ZiJin(Project_name)
            xlsheet_3_1.cells(i, 8).Value = ZhuanXian_TaiZhang(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_2).Value = ZhuanXian_YiBan(Project_name)
            xlsheet_3_1.cells(i, FindColumnNumber_3).Value = ZhuanXian_DaiBan(Project_name)
        End If
        If Project_name Like sheetName_1 Then
            xlsheet_3_1.cells(i, FindColumnNumber_2 - 1).Value = xlsheet_3_1.cells(i, 9).Value - xlsheet_3_1.cells(i, FindColumnNumber_2 - 3).Value
            For j = 0 To UBound(values)
                xlsheet_3_1.cells(i, FindColumnNumber_2 - 1).Value = xlsheet_3_1.cells(i, FindColumnNumber_2 - 1).Value + ZhuanZi(j)(Project_name)
            Next j
        End If
        If xlsheet_3_1.cells(i, 7).Value < 0 Then
            ' 设置背景颜色
            Set objRange = xlsheet_3_1.range("G" & i & ":G" & i)
            objRange.Interior.Color = RGB(255, 255, 0)
        Else
            ' 设置背景颜色
            Set objRange = xlsheet_3_1.range("G" & i & ":G" & i)
            objRange.Interior.Color = RGB(255, 255, 255)
        End If
    Next i
    
    xlBook_3.Save
    xlBook_3.Close SaveChanges:=False
    xlApp_3.Quit

    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing

    ' 更新进度条
    Form1.ProgressBar1.Value = Form1.ProgressBar1.Value + 10
    Form1.Label3.Caption = "运行进度：填写完成" & Form1.ProgressBar1.Value & "%"
    
    MsgBox ("汇总完成")

End Sub
