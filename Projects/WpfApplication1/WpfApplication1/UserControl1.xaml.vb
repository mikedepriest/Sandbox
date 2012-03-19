Partial Public Class UserControl1
    Inherits Windows.Controls.UserControl

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Visibility = Visibility.Collapsed

    End Sub

    Public Sub Show()
        Visibility = Visibility.Visible
    End Sub
    Public Sub Hide()
        Visibility = Visibility.Collapsed
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
End Class
