Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Требуется для поддержки конструктора композиции классов Windows.Forms
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'Этот вызов установлен конструктором компонентов.
        InitializeComponent()

    End Sub

    'Компонент переопределяет метод dispose для очистки списка элементов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора компонентов
    Private components As System.ComponentModel.IContainer

    'ПРИМЕЧАНИЕ. Следующая процедура является обязательной для конструктора компонентов
    'Для ее изменения используйте конструктор компонентов.
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Ribbon1))
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.Label3 = Me.Factory.CreateRibbonLabel
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.Label4 = Me.Factory.CreateRibbonLabel
        Me.ComboBox1 = Me.Factory.CreateRibbonComboBox
        Me.ComboBox2 = Me.Factory.CreateRibbonComboBox
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.CheckBox1 = Me.Factory.CreateRibbonCheckBox
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Button4 = Me.Factory.CreateRibbonButton
        Tab1 = Me.Factory.CreateRibbonTab
        Tab1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Tab1.Groups.Add(Me.Group2)
        Tab1.Groups.Add(Me.Group1)
        Tab1.Groups.Add(Me.Group3)
        Tab1.Groups.Add(Me.Group4)
        Tab1.Label = "Santen Tools"
        Tab1.Name = "Tab1"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Label2)
        Me.Group2.Items.Add(Me.Label3)
        Me.Group2.Items.Add(Me.Label1)
        Me.Group2.Label = "Instruction for Comparison Tools"
        Me.Group2.Name = "Group2"
        Me.Group2.Visible = False
        '
        'Label2
        '
        Me.Label2.Label = "1. Open [FILE1] and [FILE2] in MS Excel;"
        Me.Label2.Name = "Label2"
        '
        'Label3
        '
        Me.Label3.Label = "2. From [FILE2] choose [FILE1] in Filename/ sheet in Sheetname;"
        Me.Label3.Name = "Label3"
        '
        'Label1
        '
        Me.Label1.Label = "3. Press [Compare]."
        Me.Label1.Name = "Label1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Label4)
        Me.Group1.Items.Add(Me.ComboBox1)
        Me.Group1.Items.Add(Me.ComboBox2)
        Me.Group1.Items.Add(Me.Separator1)
        Me.Group1.Items.Add(Me.CheckBox1)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Label = "Comparison Tools"
        Me.Group1.Name = "Group1"
        Me.Group1.Visible = False
        '
        'Label4
        '
        Me.Label4.Label = "Open FILE1 & choose FILE2 to compare"
        Me.Label4.Name = "Label4"
        '
        'ComboBox1
        '
        Me.ComboBox1.Label = "Filename:"
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Text = Nothing
        '
        'ComboBox2
        '
        Me.ComboBox2.Label = "Sheetname:"
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Text = Nothing
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'CheckBox1
        '
        Me.CheckBox1.Label = "Detailed Comparison"
        Me.CheckBox1.Name = "CheckBox1"
        '
        'Button1
        '
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Label = "Compare"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.Button2)
        Me.Group3.Items.Add(Me.Button3)
        Me.Group3.Label = "Business Trips Orders"
        Me.Group3.Name = "Group3"
        '
        'Button2
        '
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Label = "Make Orders"
        Me.Button2.Name = "Button2"
        Me.Button2.ShowImage = True
        '
        'Button3
        '
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Label = "Options"
        Me.Button3.Name = "Button3"
        Me.Button3.ShowImage = True
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Button4)
        Me.Group4.Label = "Expense Reports"
        Me.Group4.Name = "Group4"
        Me.Group4.Visible = False
        '
        'Button4
        '
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Label = "Create"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Tab1)
        Tab1.ResumeLayout(False)
        Tab1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ComboBox1 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents ComboBox2 As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label3 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label4 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Friend WithEvents CheckBox1 As Microsoft.Office.Tools.Ribbon.RibbonCheckBox
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
