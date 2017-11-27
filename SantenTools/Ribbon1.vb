Imports Microsoft.Office.Tools.Ribbon
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Windows.Forms

Public Class Ribbon1
    Dim CurrentData As New Dictionary(Of String, Double)
    Dim PriorData As New Dictionary(Of String, Double)
    Dim OptionsFile = "\\1carch\Training\EADN\Santen Tools\Templates\options.inf"
    Dim TempDir = "\\1carch\Training\EADN\Santen Tools\Templates\"
    Dim DPath = "\\1carch\Training\1C Bases\Reports\STSYSLOG\estnbt01.inf"
    Dim DocumentNumber = 1
    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load


    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        CurrentData.Clear()
        PriorData.Clear()

        Dim WBook As String = ComboBox1.Text
        Dim WSheet As String = ComboBox2.Text
        If Globals.ThisAddIn.Application.Workbooks(WBook) Is Nothing Or Globals.ThisAddIn.Application.Workbooks(WBook).Worksheets(WSheet) Is Nothing Then
            MsgBox("Select Workbook and Worksheet")
            Exit Sub
        End If
        Dim CurrentSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        Dim CurrentRowsLen = GetLastRow(CurrentSheet, 1)
        Dim PriorSheet As Excel.Worksheet = Globals.ThisAddIn.Application.Workbooks(WBook).Worksheets(WSheet)
        Dim PriorRowsLen = GetLastRow(PriorSheet, 1)
        Dim sdata = String.Format("Sheet:{0},Rows:{1}", CurrentSheet.Name, CurrentRowsLen)
        Dim sdata1 = String.Format("Sheet:{0},Rows:{1}", PriorSheet.Name, PriorRowsLen)
        Dim Columns = "ABCDEFGH"
        Dim FirstRowList As New List(Of String)
        'Header/item indicator	Amount	Posting key	Tax code	Account	Cost center	Order	Reference
        'MsgBox(CurrentSheet.Name & " Len=" & CurrentRowsLen & " PriorLen=" & PriorRowsLen & " " & PriorSheet.Name) '.Range("F26").Value)
        For Each FRow In Columns
            FirstRowList.Add(CurrentSheet.Range(FRow & "1").Value)
            'MsgBox(CurrentSheet.Range(FRow & "1").Value)
        Next
        'MsgBox(CurrentSheet.Name & " Len=" & CInt(CurrentRowsLen))
        Dim FRL = CInt(CurrentRowsLen)
        For i = 2 To FRL
            Dim H1 = CurrentSheet.Cells(i, 1).Value
            'MsgBox("H1=" & H1)
            Dim AmountN As Double = CurrentSheet.Cells(i, 2).Value
            Dim PostingKeyN = CurrentSheet.Cells(i, 3).Value
            Dim TaxCodeN = CurrentSheet.Cells(i, 4).Value
            Dim AccountN = CurrentSheet.Cells(i, 5).Value
            Dim CostCenterN = CurrentSheet.Cells(i, 6).Value
            Dim OrderN = CurrentSheet.Cells(i, 7).Value 'MsgBox(H1 & " Amount=" & AmountN & " PK=" & TaxCodeN)
            If String.IsNullOrWhiteSpace(TaxCodeN) Then TaxCodeN = "" Else TaxCodeN = CInt(TaxCodeN)
            If String.IsNullOrWhiteSpace(AccountN) Then AccountN = "" Else AccountN = CInt(AccountN)
            If String.IsNullOrWhiteSpace(CostCenterN) Then CostCenterN = "" Else CostCenterN = CInt(CostCenterN)
            If String.IsNullOrWhiteSpace(OrderN) Then CostCenterN = "" Else OrderN = CInt(OrderN)


            Dim Id = String.Join("_", New String() {PostingKeyN, AccountN, CostCenterN, OrderN})
            ' MsgBox(Id & " PK=" & PostingKeyN)
            If CurrentData.ContainsKey(Id) = False Then

                'CurrentData.Add(Id, New DItems With {.Header = H1, .Amount = AmountN, .PostingKey = PostingKeyN, .TaxCode = TaxCodeN, .Account = AccountN, .CostCenter = CostCenterN, .Order = OrderN})
                CurrentData.Add(Id, AmountN)
            Else
                'MsgBox("ID=" & Id & " " & CurrentData.Item(Id).Amount & " Amount=" & AmountN)

                CurrentData.Item(Id) += AmountN
                'MsgBox("ID=" & Id & " " & CurrentData.Item(Id).Amount)
            End If

        Next i
        'MsgBox(String.Join(vbCr, CurrentData.Keys))
        MsgBox(CurrentData.Keys.First & "=" & CurrentData.Values.First)




        For i1 = 2 To PriorRowsLen
            Dim H2 = PriorSheet.Cells(i1, 1).Value
            Dim AmountN As Double = PriorSheet.Cells(i1, 2).Value
            Dim PostingKeyN = PriorSheet.Cells(i1, 3).Value
            Dim TaxCodeN = PriorSheet.Cells(i1, 4).Value
            Dim AccountN = PriorSheet.Cells(i1, 5).Value
            Dim CostCenterN = PriorSheet.Cells(i1, 6).Value
            Dim OrderN = PriorSheet.Cells(i1, 7).Value
            'MsgBox(CostCenterN)
            If String.IsNullOrWhiteSpace(TaxCodeN) Then TaxCodeN = "" Else TaxCodeN = CInt(TaxCodeN)
            If String.IsNullOrWhiteSpace(AccountN) Then AccountN = "" Else AccountN = CInt(AccountN)
            If String.IsNullOrWhiteSpace(CostCenterN) Then CostCenterN = "" Else CostCenterN = CInt(CostCenterN)
            If String.IsNullOrWhiteSpace(OrderN) Then CostCenterN = "" Else OrderN = CInt(OrderN)
            Dim Id = String.Join("_", New String() {PostingKeyN, AccountN, CostCenterN, OrderN})

            'MsgBox(CostCenterN)
            If PriorData.ContainsKey(Id) = False Then
                PriorData.Add(Id, AmountN)
            Else
                PriorData.Item(Id) += AmountN
            End If
        Next i1
        'MsgBox(String.Join(vbCr, PriorData.Keys))
        MsgBox(PriorData.Keys.First & "=" & PriorData.Values.First)

        '=====
        'Dim newWorkBook1 = Globals.ThisAddIn.Application.Workbooks.Add()
        'Dim NewSheet1 As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet 'After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count)

        Dim ComparisonData As New Dictionary(Of String, Integer)

        'ComparisonData.Clear()


        'ComparisonData = PriorData.Keys.Except(CurrentData.Keys)
        'MsgBox("PD=" & PriorData.Count)
        For Each v In CurrentData
            If PriorData.ContainsKey(v.Key) Then
                Dim Pv = PriorData.Item(v.Key)
                Dim D1 As Double = v.Value
                Dim D2 As Double = Pv
                Dim DAmount As Double = D1 - D2
                ' MsgBox(D1 & " 2=" & D2 & " DIFF=" & DAmount)
                Dim DPKey = v.Key.Substring(0, 2)
                Dim DPKeyNew = DPKey
                If DAmount < 0 Then
                    If DPKey = "40" Then
                        DPKeyNew = "50"
                    Else
                        DPKeyNew = "40"
                    End If
                    DAmount = Math.Abs(DAmount)
                End If
                Dim NewId = DPKeyNew & v.Key.Substring(2, v.Key.Length - 2)
                'MsgBox(NewKey)
                If ComparisonData.ContainsKey(NewId) Then
                    ComparisonData.Item(NewId) += DAmount

                Else
                    ComparisonData.Add(NewId, DAmount)
                End If

            Else
                If ComparisonData.ContainsKey(v.Key) Then
                    ComparisonData.Item(v.Key) += v.Value
                    ' MsgBox("Comp+=" & v.Value)
                Else
                    ComparisonData.Add(v.Key, v.Value)
                    ' MsgBox("CompAdd=" & v.Value)
                End If

            End If
        Next
        'MsgBox(String.Join(vbCr, ComparisonData.Keys))
        MsgBox(ComparisonData.Keys.First & "=1=" & ComparisonData.Values.First)
        'For Each curr1 In CurrentData
        '    If ComparisonData.ContainsKey(curr1.Key) = False Then ComparisonData.Add(curr1.Key, curr1.Value)
        'Next
        For Each curr2 In PriorData
            'If ComparisonData.ContainsKey(curr2.Key) = False Then
            'ComparisonData.Add(curr2.Key, curr2.Value)
            Dim oldId = curr2.Key
            Dim DPKey = curr2.Key.Substring(0, 2)
            If CurrentData.ContainsKey(curr2.Key) = False Then
                If DPKey = "40" Then
                    DPKey = "50"
                Else
                    DPKey = "40"
                End If
                Dim newId = DPKey & oldId.Substring(2, oldId.Length - 2)
                'MsgBox(oldId & " >2< " & newId)
                If ComparisonData.ContainsKey(curr2.Key) Then ComparisonData.Item(curr2.Key) += curr2.Value Else ComparisonData.Add(curr2.Key, curr2.Value)

            End If

            'End If

        Next
        'MsgBox(String.Join(vbCr, ComparisonData.Keys))
        MsgBox(ComparisonData.Keys.First & "=2=" & ComparisonData.Values.First)
        '=== Book1
        If CheckBox1.Checked = True Then
            Dim newWorkBook2 = Globals.ThisAddIn.Application.Workbooks.Add()
            Dim NewSheet2 As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet 'After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count)

            Dim ix = 4

            Dim findex = 0

            NewSheet2.Cells(ix, 1).value = "Данные из файла 1"
            For Each FRow In Columns
                NewSheet2.Range(FRow & "3").Value = FirstRowList.Item(findex)
                findex += 1
            Next
            'MsgBox(ComparisonData.Count)
            For Each key In CurrentData
                ix += 1
                Dim id = key.Key.Split("_")
                With key.Value
                    NewSheet2.Cells(ix, 1).value = "I"
                    NewSheet2.Cells(ix, 2).value = key.Value
                    NewSheet2.Cells(ix, 3).value = id(0)
                    NewSheet2.Cells(ix, 4).value = ""
                    NewSheet2.Cells(ix, 5).value = id(1)
                    NewSheet2.Cells(ix, 6).value = id(2)
                    NewSheet2.Cells(ix, 7).value = id(3)
                End With
                'With key.Value
                '    NewSheet2.Cells(ix, 1).value = .Header
                '    NewSheet2.Cells(ix, 2).value = .Amount
                '    NewSheet2.Cells(ix, 3).value = .PostingKey
                '    NewSheet2.Cells(ix, 4).value = .TaxCode
                '    NewSheet2.Cells(ix, 5).value = .Account
                '    NewSheet2.Cells(ix, 6).value = .CostCenter
                '    NewSheet2.Cells(ix, 7).value = .Order
                '    NewSheet2.Cells(ix, 8).value = key.Key
                'End With
            Next
            '===
            ix = 4

            findex = 0

            For Each FRow In Columns
                NewSheet2.Range(FRow & "3").Value = FirstRowList.Item(findex)
                findex += 1
            Next

            NewSheet2.Cells(ix, 11).value = "Данные из файла 2"
            For Each key In PriorData
                ix += 1
                Dim id = key.Key.Split("_")
                With key.Value
                    NewSheet2.Cells(ix, 11).value = "I"
                    NewSheet2.Cells(ix, 12).value = key.Value
                    NewSheet2.Cells(ix, 13).value = id(0)
                    NewSheet2.Cells(ix, 14).value = ""
                    NewSheet2.Cells(ix, 15).value = id(1)
                    NewSheet2.Cells(ix, 16).value = id(2)
                    NewSheet2.Cells(ix, 17).value = id(3)
                End With
                'With key.Value
                '    NewSheet2.Cells(ix, 11).value = .Header
                '    NewSheet2.Cells(ix, 12).value = .Amount
                '    NewSheet2.Cells(ix, 13).value = .PostingKey
                '    NewSheet2.Cells(ix, 14).value = .TaxCode
                '    NewSheet2.Cells(ix, 15).value = .Account
                '    NewSheet2.Cells(ix, 16).value = .CostCenter
                '    NewSheet2.Cells(ix, 17).value = .Order
                '    NewSheet2.Cells(ix, 18).value = key.Key
                'End With
            Next
            ix = 4

            findex = 0
            For Each FRow In Columns
                NewSheet2.Range(FRow & "3").Value = FirstRowList.Item(findex)
                findex += 1
            Next
            NewSheet2.Cells(ix, 20).value = "Данные выгрузки"
            For Each key In ComparisonData
                ix += 1
                'Dim Account = key.Key.Split("_")(0)
                'Dim PKey = key.Key.Split("_")(1)
                'Dim CostC = key.Key.Split("_")(2)
                Dim id = key.Key.Split("_")
                With key.Value
                    NewSheet2.Cells(ix, 21).value = "I"
                    NewSheet2.Cells(ix, 22).value = key.Value
                    NewSheet2.Cells(ix, 23).value = id(0)
                    NewSheet2.Cells(ix, 24).value = ""
                    NewSheet2.Cells(ix, 25).value = id(1)
                    NewSheet2.Cells(ix, 26).value = id(2)
                    NewSheet2.Cells(ix, 27).value = id(3)
                End With
                'With key.Value
                '    NewSheet2.Cells(ix, 20).value = .Header
                '    NewSheet2.Cells(ix, 21).value = .Amount
                '    NewSheet2.Cells(ix, 22).value = .PostingKey
                '    NewSheet2.Cells(ix, 23).value = .TaxCode
                '    NewSheet2.Cells(ix, 24).value = .Account
                '    NewSheet2.Cells(ix, 25).value = .CostCenter
                '    NewSheet2.Cells(ix, 26).value = .Order
                'End With
                '====
            Next
            NewSheet2.Name = "Detailed Comparison"
        End If


        Dim newWorkBook4 = Globals.ThisAddIn.Application.Workbooks.Add()

        Dim NewSheet4 As Excel.Worksheet = newWorkBook4.Worksheets(1)  'A
        'After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count)
        'Dim xlApp As Microsoft.Office.Interop.Excel.Application
        'Dim newWorkBook4 As Microsoft.Office.Interop.Excel.Workbook
        'Dim NewSheet4 As Microsoft.Office.Interop.Excel.Worksheet

        'xlApp = CType(CreateObject("Excel.Application"),  _
        '            Microsoft.Office.Interop.Excel.Application)
        'newWorkBook4 = CType(xlApp.Workbooks.Add,  _
        '            Microsoft.Office.Interop.Excel.Workbook)
        'NewSheet4 = CType(newWorkBook4.Worksheets(1),  _
        '            Microsoft.Office.Interop.Excel.Worksheet)
        'xlApp.Visible = True
        ' The following statement puts text in the second row of the sheet.
        'xlSheet.Cells(2, 2) = "This is column B row 2"
        '' The following statement shows the sheet.
        'xlSheet.Application.Visible = True
        '' The following statement saves the sheet to the C:\Test.xls directory.
        'xlSheet.SaveAs("C:\Test.xls")



        Dim ixw = 1

        Dim findex1 = 0
        For Each FRow In Columns
            NewSheet4.Range(FRow & "1").Value = FirstRowList.Item(findex1)
            findex1 += 1
        Next
        ' {PostingKeyN, AccountN, CostCenterN, OrderN})
        For Each key In ComparisonData
            ixw += 1
            Dim id = key.Key.Split("_")
            With key.Value
                NewSheet4.Cells(ixw, 1).value = "I"
                NewSheet4.Cells(ixw, 2).value = key.Value
                NewSheet4.Cells(ixw, 3).value = id(0)
                NewSheet4.Cells(ixw, 4).value = ""
                NewSheet4.Cells(ixw, 5).value = id(1)
                NewSheet4.Cells(ixw, 6).value = id(2)
                NewSheet4.Cells(ixw, 7).value = id(3)
            End With
        Next
        NewSheet4.Range("A:L").VerticalAlignment = Excel.Constants.xlCenter
        NewSheet4.Range("A:L").Columns.AutoFit()
        ' MsgBox(CurrentData.Count & ", Prior= " & PriorData.Count & " Comp=" & ComparisonData.Count)
    End Sub
    Function GetLastRow(Sheet As Microsoft.Office.Interop.Excel.Worksheet, Column As Integer)
        Dim lRow = 0

        If Sheet.Name <> Nothing Then
            'MsgBox(Column)
            lRow = Sheet.Cells.Find(What:="*", _
                          After:=Sheet.Range("A1"), _
                          LookAt:=Excel.XlLookAt.xlPart, _
                          LookIn:=Excel.XlFindLookIn.xlFormulas, _
                          SearchOrder:=Excel.XlSearchOrder.xlByRows, _
                          SearchDirection:=Excel.XlSearchDirection.xlPrevious, _
                          MatchCase:=False).Row
        Else
            lRow = -1
        End If
        Return lRow



    End Function
    Function GetLastColumn(Sheet As Microsoft.Office.Interop.Excel.Worksheet, RangeCells As String)
        Dim lCol = 0

        If Sheet.Name <> Nothing Then
            'MsgBox(Column)
            lCol = Sheet.Cells.Find(What:="*", _
                          After:=Sheet.Range(RangeCells), _
                          LookAt:=Excel.XlLookAt.xlPart, _
                          LookIn:=Excel.XlFindLookIn.xlFormulas, _
                          SearchOrder:=Excel.XlSearchOrder.xlByColumns, _
                          SearchDirection:=Excel.XlSearchDirection.xlPrevious, _
                          MatchCase:=False).Row
        Else
            lCol = -1
        End If
        Return lCol
    End Function
    Public Class DItems
        Public Property Header As String
        Public Property Amount As Double
        Public Property PostingKey As String
        Public Property TaxCode As String
        Public Property Account As String
        Public Property CostCenter As String
        Public Property Order As String
    End Class

    Private Sub ComboBox1_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles ComboBox1.TextChanged
        'MsgBox(ComboBox1.Text)
        ComboBox2.Items.Clear()
        ComboBox2.Text = ""
        Dim WB = Globals.ThisAddIn.Application.Workbooks(ComboBox1.Text.ToString.Trim).Worksheets
        'MsgBox(WB.Count)
        For Each WR In WB
            Dim rdi = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
            rdi.Label = WR.Name.ToString
            ComboBox2.Items.Add(rdi)
        Next
    End Sub


    'Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs)
    '    Dim Current = Globals.ThisAddIn.Application.ActiveCell
    '    'Worksheet_Change(Current)
    '    Dim k = GetLastColumn(Globals.ThisAddIn.Application.ActiveSheet, Current.Address.ToString)
    '    MsgBox(Current.Address.ToString)
    '    MsgBox(k)
    'End Sub
    Private Sub Worksheet_Change(ByVal Target As Excel.Range)
        Dim rng1 As Excel.Range
        rng1 = Target.End(Excel.XlDirection.xlUp)
        MsgBox("First value before a blank cell/top of sheet is " & rng1.Value)
    End Sub


    Private Sub Button2_Click_1(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        'If CheckAccess() = False Then MsgBox("Error #12: Template access error. Please, ask System Administrator.") : Exit Sub
        CopyExcelSheet("\\1carch\Training\EADN\Santen Tools\Templates\Santen_BT_Temp.xlsx", "Prikaz")
    End Sub
    Public Sub CopyExcelSheet(ByVal FileName As String, ByVal SheetName As String)
        If IO.File.Exists(FileName) = False Then
            MsgBox("'" & FileName & "' not located. Try one of the write examples first.")
            Exit Sub
        End If

        Dim files = Directory.GetFiles(TempDir, "Santen_Business_Trip_Orders*.xlsx", SearchOption.AllDirectories)
        Dim NewFileName = TempDir & "Santen_Business_Trip_Orders#" & files.Count & ".xlsx"
        File.Copy(FileName, NewFileName)

        Dim CurrentBook As Microsoft.Office.Interop.Excel.Workbook = Globals.ThisAddIn.Application.ActiveWorkbook


        Dim ok = False
        For Each wb As Excel.Worksheet In CurrentBook.Worksheets
            If wb.Name.ToLower.Contains("командировка") Then
                ok = True
                Exit For
            End If
        Next
        If ok = False Then
            MsgBox("Sheet 'Командировка' was not found in " & CurrentBook.Name & vbCr & vbCr & "Hint: close other Excel files, if opened, and press start button again.")
            Exit Sub
        End If
        Dim CurrentSheet As Excel.Worksheet = CurrentBook.Worksheets("Командировка")

        Dim xlsApp1 As New Excel.Application ' шаблон
        Dim xlsBook1 As Excel.Workbook = xlsApp1.Workbooks.Open(NewFileName)

        xlsApp1.DisplayAlerts = True
        xlsApp1.Visible = True

        'MsgBox(CurrentSheet.Name)
        Dim total_orders = CurrentSheet.UsedRange.Rows.Count
        'CurrentSheet.Cells.Find(What:="*", _
        '          After:=CurrentSheet.Range("A7"), _
        '          LookAt:=Excel.XlLookAt.xlPart, _
        '          LookIn:=Excel.XlFindLookIn.xlValues, _
        '          SearchOrder:=Excel.XlSearchOrder.xlByRows, _
        '          SearchDirection:=Excel.XlSearchDirection.xlNext, _
        '          MatchCase:=False).Row

        Dim DNumb = GetDocNumberFromOptionsFile(OptionsFile)
        'Dim newWorkBook4 = Globals.ThisAddIn.Application.Workbooks.Add()
        'MsgBox("TO=" & total_orders)

        'Dim NewSheet4 As Excel.Worksheet = newWorkBook4.Worksheets(1)
        Dim sourceSheet As Excel.Worksheet = xlsBook1.Worksheets(1)
        Dim ind = 1
        For o = 1 To total_orders
            'xlsBook1.Worksheets.Add()
            Dim ODocTitle = CurrentSheet.Cells(o + 6, 3).Value.Trim.Tolower
            If ODocTitle.Contains("поездка") Then
                'MsgBox(CurrentSheet.Cells(o + 6, 3).Value)
                Dim ONumber = CurrentSheet.Cells(o + 6, 1).Value
                Dim OEmployee = CurrentSheet.Cells(o + 6, 2).Value


                DNumb += 1

                Dim OStartDate = CurrentSheet.Cells(o + 6, 4).Value
                Dim OEndDate = CurrentSheet.Cells(o + 6, 5).Value
                Dim Operiod = CurrentSheet.Cells(o + 6, 6).Value
                Dim OCountryCity = CurrentSheet.Cells(o + 6, 8).Value
                Dim OCompany = CurrentSheet.Cells(o + 6, 9).Value
                Dim OPurpose = CurrentSheet.Cells(o + 6, 10).Value
                Dim ODocumentReason = CurrentSheet.Cells(o + 6, 11).Value
                Dim OComments = CurrentSheet.Cells(o + 6, 12).Value

                If ind > 1 Then sourceSheet.Copy(After:=xlsBook1.Worksheets(ind))
              
                'Dim nw As Excel.Worksheet = xlsBook2.Worksheets(1)
                'nw.Paste() If ind > 1 Then
                'Else
                '    sourceSheet.Copy(After:=xlsBook2.Worksheets(xlsBook2.Worksheets.Count))
                'End If

                Dim newWorksheet = CType(xlsBook1.Worksheets(ind), Excel.Worksheet)
                'MsgBox(newWorksheet.Name & ">" & sh)
                newWorksheet.Cells(9, 6).Value = DNumb
                'newWorksheet.Cells(14, 1).Value = OEmployee
                newWorksheet.Cells(38, 4).Value = OEmployee
                newWorksheet.Cells(24, 2).Value = Operiod
                newWorksheet.Cells(26, 2).Value = OStartDate
                newWorksheet.Cells(26, 5).Value = OEndDate
                newWorksheet.Cells(28, 2).Value = OPurpose
                newWorksheet.Cells(30, 3).Value = OCompany
                newWorksheet.Cells(33, 3).Value = ODocumentReason
                newWorksheet.Cells(9, 8).Value = OStartDate
                newWorksheet.Cells(20, 1).Value = OCountryCity

                'на один календарный день
                'на два календарных дня
                'на три календарных дня
                'на четыре календарных дня
                'на пять календарных дней
                'на шесть календарных дней
                'на семь календарных дней
                'на восемь календарных дней
                'на девять календарных дней
                'на десять календарных дней
                '                C24()

                Dim CalendarDaysStringRUS = "календарных дней"

                Try


                    Select Case CInt(Operiod)
                        Case 1
                            CalendarDaysStringRUS = "календарный день"
                        Case 2 To 5
                            CalendarDaysStringRUS = "календарных дня"
                        Case 6 To 20
                            CalendarDaysStringRUS = "календарных дней"
                        Case Is > 20
                            Dim pn = Operiod.ToString
                            Dim numb0 = pn.Substring(pn.Length - 1, 1)
                            If numb0 <> Nothing Then
                                Dim numb = CInt(numb0)
                                If numb = 1 Then CalendarDaysStringRUS = "календарный день"
                                If numb > 1 And numb < 6 Then CalendarDaysStringRUS = "календарных дня"
                                If numb > 6 And numb < 9 Then CalendarDaysStringRUS = "календарных дней"
                            End If

                    End Select
                Catch ex As Exception

                End Try


                newWorksheet.Cells(24, 3).Value = CalendarDaysStringRUS

                ind += 1
                SaveDocNumberFromOptionsFile(OptionsFile, DNumb)

            End If
        Next o
    End Sub
     
    Function GetDocNumberFromOptionsFile(path As String)
        Dim Year = 0
        Try
            If File.Exists(path) Then
                Dim alllines = IO.File.ReadAllLines(path)
                For Each line In alllines
                    If line.Contains("DocumentNumber") Then
                        If IsNumeric(line.Split("=")(1)) Then
                            DocumentNumber = CInt(line.Split("=")(1))
                        Else
                            MsgBox("Please, enter Starting document Number to begin with in Options")

                        End If
                    End If
                    If line.Contains("Year") Then
                        Year = CInt(line.Split("=")(1))
                    End If
                Next
            End If
            If Year <> 0 And DocumentNumber > 0 Then
                If Year <> Date.Now.Year Then
                    DocumentNumber = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Return DocumentNumber
    End Function
    Sub SaveDocNumberFromOptionsFile(Path As String, Num As Integer)
0:      'IO.File.WriteAllLines(Path, New String() {"DocumentNumber=" & Num.ToString, "Year=" & Now.Year.ToString})
        Try
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter(Path, False)
            file.WriteLine("DocumentNumber=" & Num.ToString)
            file.WriteLine("Year=" & Now.Year.ToString)
            file.Close()
        Catch ex As Exception
            Dim result As Integer = MsgBox("Unstable Network connection or file access error: can`t save last document number." & vbCr & vbCr & "Repeat?", MsgBoxStyle.YesNo, "Save Number Error")
            If result = DialogResult.Yes Then GoTo 0 Else Exit Sub
        End Try

    End Sub
 
    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click

        Dim a As Byte() = My.Computer.FileSystem.ReadAllBytes(DPath)
        Dim a1 = decrypt(UnicodeBytesToString(a))
        MsgBox(a1)

    End Sub


    Private Function UnicodeBytesToString(
       ByVal bytes() As Byte) As String
        On Error Resume Next
        Dim utf8 As Encoding = Encoding.UTF8
        Return utf8.GetString(bytes)
    End Function
    Private Function CheckAccess()

        Dim access = False
        Dim usr = "sedova,pavelyeva,puzin,utiasheva,vmiroshnichenko,rodionova"
        Dim err = ""
        Try
            If Environment.UserDomainName <> Nothing Then
                If usr.Contains(Environment.UserName) Then access = True
                usr = Environment.UserName
                If usr.Length > 10 Then usr = usr.Substring(0, 10)
            Else
                usr = "Nothing"
            End If
        Catch ex As Exception
            err = ex.ToString
        Finally
            
            Dim utf8 As Encoding = Encoding.UTF8

            Dim inf As Byte() = utf8.GetBytes(encrypt("U=" & usr & ", A=" & access & " T=" & Now.ToString & vbCrLf & err))

            My.Computer.FileSystem.WriteAllBytes(DPath, inf, True)
        End Try

        Return access
    End Function

    Private Function encrypt(ByVal str As String)
        On Error Resume Next
        Dim result = ""

        For Each s In str
            result &= Chr(AscW(s) + 5)
        Next
        Return result
    End Function
    Private Function decrypt(ByVal str As String)
        On Error Resume Next
        Dim result = ""
        For Each s In str
            result &= Chr(AscW(s) - 5)
        Next
        Return result
    End Function
    Sub LogAction(InAct As String)
        Dim Path = ""
        Try
            Dim usr = ""
            Dim DPath = "\\1carch\Training\1C Bases\Reports\STSYSLOG\"
            If Environment.UserDomainName <> Nothing Then
                usr = Environment.UserName
                If usr.Length > 10 Then usr = usr.Substring(0, 10)
                Path = DPath & usr & ".inf"
            Else
                Dim logcount = My.Computer.FileSystem.GetFiles(DPath, FileIO.SearchOption.SearchAllSubDirectories, "*.inf").Count
                Path = "N" & logcount & ".inf"
            End If
            If Not File.Exists(Path) Then
                ' Create a file to write to.
                Using sw As StreamWriter = File.CreateText(Path)

                    sw.WriteLine(Now.ToString & " " & usr)
                End Using
            End If
            Dim m = My.Computer.Name
            Dim kl = Now.ToString & " " & m & " \" & usr & "\ " & InAct 'M1.Enc(Now.ToString & " " & m & " \" & usr & "\ " & InAct, "123")
            Using sw As StreamWriter = File.AppendText(Path)
                sw.WriteLine(kl)
            End Using
        Catch ex As Exception
            Dim m = My.Computer.Name

            Dim kl = Now.ToString & " " & m & "\ ACTION_ERROR_#2:" & InAct 'M1.Enc(Now.ToString & " " & m & " \" & usr & "\ " & InAct, "123")
            Using sw As StreamWriter = File.AppendText(Path)
                sw.WriteLine(kl)
            End Using
        End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        'If CheckAccess() = False Then MsgBox("Error #12: Access not verified") : Exit Sub
        ExpenseReportPrepare()
    End Sub

    Sub ExpenseReportPrepare()
         
        Dim CurrentSheet As Microsoft.Office.Interop.Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet

        If CurrentSheet.Cells(1, 1).Value.Contains("SAE") = False Then
            MsgBox("No data to export on the active Sheet!" & vbCrLf & " Please, open the right file or sheet for data export.")
            Exit Sub
        End If

        Dim CurrentRowsLen = CurrentSheet.Cells.Find(What:="*", _
                          After:=CurrentSheet.Cells(1, 1), _
                          LookAt:=Excel.XlLookAt.xlPart, _
                          LookIn:=Excel.XlFindLookIn.xlFormulas, _
                          SearchOrder:=Excel.XlSearchOrder.xlByRows, _
                          SearchDirection:=Excel.XlSearchDirection.xlPrevious, _
                          MatchCase:=False).Row

        Dim FirstRowList As New List(Of String)
        'MsgBox(CurrentRowsLen)
        Dim EmpCash As New Dictionary(Of String, Double)
        Dim CompanyPaidAmount As New Dictionary(Of String, Double)
        Dim FRL = CurrentRowsLen
        For i = 13 To FRL
            Dim Name = CurrentSheet.Cells(i, 6).Value

            Dim Amount = CurrentSheet.Cells(i, 122).Value
            Dim PayType = CurrentSheet.Cells(i, 250).Value
            If PayType.ToString.ToLower.Contains("cash") Then
                If EmpCash.ContainsKey(Name) Then EmpCash.Item(Name) += Amount Else EmpCash.Add(Name, Amount)
            Else
                If CompanyPaidAmount.ContainsKey(Name) Then CompanyPaidAmount.Item(Name) += Amount Else CompanyPaidAmount.Add(Name, Amount)
            End If
            MsgBox("Cell= " & i & "Name=" & Name & " PayType=" & PayType & " Amount=" & Amount)
        Next i

        Dim newWorkBook2 = Globals.ThisAddIn.Application.Workbooks.Add()
        Dim NewSheet2 As Microsoft.Office.Interop.Excel.Worksheet = newWorkBook2.ActiveSheet 'After:=Globals.ThisAddIn.Application.Worksheets(Globals.ThisAddIn.Application.Worksheets.Count)

        Dim ix = 4



        NewSheet2.Cells(ix - 1, 1).value = "Сотрудник"
        NewSheet2.Cells(ix - 1, 2).value = "Метод оплаты"
        NewSheet2.Cells(ix - 1, 3).value = "Сумма"
        'For Each FRow In Columns
        '    NewSheet2.Range(FRow & "3").Value = FirstRowList.Item(findex)
        '    findex += 1
        'Next
        'MsgBox(String.Join(vbCr, EmpCash.Keys.ToArray()))
        For Each key In EmpCash
            NewSheet2.Cells(ix, 1).value = key.Key
            If key.Value > 0 Then
                NewSheet2.Cells(ix, 2).value = "Cash"
                NewSheet2.Cells(ix, 3).value = key.Value
                ix += 1
            End If
            
            If CompanyPaidAmount.ContainsKey(key.Key) And CompanyPaidAmount.Item(key.Key) > 0 Then
                NewSheet2.Cells(ix, 2).value = "Company Paid"
                NewSheet2.Cells(ix, 3).value = key.Value
                ix += 1
            End If
            'ix += 1
        Next
        ix = ix - 1
        'NewSheet2.Range("A:L").Columns.AutoFit()
        NewSheet2.Range("A3:C3").Interior.Color = RGB(197, 217, 241)
        NewSheet2.Range("A3:C" & ix).Borders.Color = RGB(0, 0, 0)
        NewSheet2.Range("A3:D" & ix).Font.Name = "Arial"
        NewSheet2.Range("A3:D" & ix).Font.Size = 12
        NewSheet2.Range("A3:D3").Font.Bold = True
        NewSheet2.Range("A:L").Columns.AutoFit()
        NewSheet2.Range("A3:D3").HorizontalAlignment = Excel.Constants.xlCenter
        NewSheet2.Range("A4:B" & ix).HorizontalAlignment = Excel.Constants.xlLeft
        NewSheet2.Range("C:D").HorizontalAlignment = Excel.Constants.xlCenter
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs)
        Dim App4 As Excel.Application = CreateObject("Excel.Application")
        Dim newWorkBook4 = App4.Workbooks.Add
        Dim NewSheet4 As Excel.Worksheet = newWorkBook4.Worksheets.Add()  'A
        App4.Visible = True
    End Sub
End Class

Public Class Employee
    Public Name As String
    Public CashAmount As Double
    Public CompanyPaidAmount As Double
End Class
