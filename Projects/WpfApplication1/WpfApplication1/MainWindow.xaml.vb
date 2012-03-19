Class MainWindow 
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        TestUserControl.TextBlock1.Text = "Bloody 'ell, textblock"
        TestUserControl.Show()
    End Sub
End Class
