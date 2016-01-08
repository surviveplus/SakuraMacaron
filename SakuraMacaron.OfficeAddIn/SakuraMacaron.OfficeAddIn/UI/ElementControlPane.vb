Namespace UI

    Public Class ElementControlPane(Of T As Windows.UIElement)

        Public Property Pane As Microsoft.Office.Tools.CustomTaskPane

        Public ReadOnly Property Control As InteropControl(Of T)

        Public Sub New(elementControl As T)

            Me.Control = New InteropControl(Of T)(elementControl)
        End Sub

        Private isFirst As Boolean = True
        Public Sub Show()
            Me.Pane.Visible = True
            Me.Control.ElementHost.Focus()
            If Me.isFirst Then System.Windows.Forms.SendKeys.Send(vbTab)
            Me.Control.ElementControl.Focus()

            Me.isFirst = False
        End Sub

    End Class

End Namespace