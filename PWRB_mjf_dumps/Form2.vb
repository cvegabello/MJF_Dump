Public Class Form2
    Public _form1 As Form1

    Public Sub New(ByVal form1 As Form1)
        'MyBase.New()

        'El Diseñador de Windows Forms requiere esta llamada.
        'Infragistics.Win.AppStyling.StyleManager.Load(Application.StartupPath + "\Nautilus.isl")
        InitializeComponent()
        _form1 = form1

        'Agregar cualquier inicialización después de la llamada a InitializeComponent()

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        SaveSettingsHoursMinutes()
        Me.Dispose()
    End Sub

    Private Sub SaveSettingsHoursMinutes()
        Utils.SalvarSetting(appName, "HoursFrecuency", ComboBox1.Text)
        Utils.SalvarSetting(appName, "MinutesFrecuency", ComboBox2.Text)

        _form1.TextBox2.Text = ComboBox1.Text
        _form1.TextBox3.Text = ComboBox2.Text

    End Sub

    Private Sub GetSettingsHoursMinutes()
        ComboBox1.Text = Utils.GetSetting(appName, "HoursFrecuency", "").ToString()
        ComboBox2.Text = Utils.GetSetting(appName, "MinutesFrecuency", "").ToString()


    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        GetSettingsHoursMinutes()
    End Sub
End Class