Namespace UI
    Public Class ElementControlToolWindow(Of T As Windows.UIElement)

        Public Property ToolWindow As System.Windows.Window
        Public ReadOnly Property Control As T

        Public Sub New(elementControl As T)

            Me.Control = elementControl
        End Sub

        Public Sub Show()
            If Me.ToolWindow Is Nothing Then
                Me.ToolWindow = New System.Windows.Window With {.WindowStyle = System.Windows.WindowStyle.ToolWindow, .Content = Me.Control, .Width = 350, .Height = 400, .Title = "Sample ToolWindow", .WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner}

                AddHandler Me.ToolWindow.Closed,
                    Sub(sender, e)
                        Me.ToolWindow = Nothing
                    End Sub

                Dim helper As New System.Windows.Interop.WindowInteropHelper(Me.ToolWindow)
                helper.Owner = Process.GetCurrentProcess().MainWindowHandle
                Me.ToolWindow.Show()

            End If
            Me.ToolWindow.Activate()
            Me.Control.Focus()
        End Sub
    End Class

End Namespace