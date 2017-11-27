Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        AddHandler Globals.ThisAddIn.Application.WorkbookOpen, AddressOf MyWorkbookOpenEvent
        AddHandler Globals.ThisAddIn.Application.NewWorkbook, AddressOf MyNewWorkbookEvent
        If Globals.ThisAddIn.Application.Workbooks.Count = 1 Then MyWorkbookOpenEvent(Globals.ThisAddIn.Application.Workbooks(1))

    End Sub

    Private Sub MyWorkbookOpenEvent(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook)
        'System.Windows.Forms.MessageBox.Show("OPEN workbook event")
        With Globals.Ribbons.Ribbon1
            .ComboBox1.Items.Clear()
            .ComboBox2.Items.Clear()
            For Each w As Microsoft.Office.Interop.Excel.Workbook In Globals.ThisAddIn.Application.Workbooks
                Dim rdi = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                rdi.Label = w.Name.ToString
                .ComboBox1.Items.Add(rdi)
            Next
           
        End With

    End Sub

    Private Sub MyNewWorkbookEvent(ByVal Wb As Microsoft.Office.Interop.Excel.Workbook)
        'System.Windows.Forms.MessageBox.Show("NEW Workbook event")
        With Globals.Ribbons.Ribbon1
            .ComboBox1.Items.Clear()
            .ComboBox2.Items.Clear()
            For Each w As Microsoft.Office.Interop.Excel.Workbook In Globals.ThisAddIn.Application.Workbooks
                Dim rdi = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                rdi.Label = w.Name.ToString
                .ComboBox1.Items.Add(rdi)
            Next

        End With
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveHandler Globals.ThisAddIn.Application.WorkbookOpen, AddressOf MyWorkbookOpenEvent
        RemoveHandler Globals.ThisAddIn.Application.NewWorkbook, AddressOf MyNewWorkbookEvent
    End Sub

End Class
