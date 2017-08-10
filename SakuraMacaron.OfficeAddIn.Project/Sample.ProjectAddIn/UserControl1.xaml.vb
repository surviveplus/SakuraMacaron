Public Class UserControl1
    Private Sub doButton_Click(sender As Object, e As Windows.RoutedEventArgs)
        RaiseEvent DoButtonClick(Me, EventArgs.Empty)
    End Sub

    Public Event DoButtonClick As EventHandler(Of EventArgs)

End Class
