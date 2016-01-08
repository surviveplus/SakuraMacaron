Namespace UI
    Public Class InteropControl(Of T As Windows.UIElement)
        Inherits System.Windows.Forms.UserControl

        Public Sub New(elementControl As T)
            If elementControl Is Nothing Then Throw New ArgumentNullException("elementControl")
            Me.ElementControl = elementControl

            Me.InitializeComponent()
        End Sub

        Private Sub InitializeComponent()

            Me.ElementHost = New System.Windows.Forms.Integration.ElementHost()
            Me.SuspendLayout()

            With Me.ElementHost
                .Dock = Windows.Forms.DockStyle.Fill
                .TabIndex = 0
                .Child = Me.ElementControl
            End With

            Me.Controls.Add(Me.ElementHost)
            Me.ResumeLayout()
        End Sub

        Public WithEvents ElementHost As Windows.Forms.Integration.ElementHost
        Public ElementControl As T

    End Class

End Namespace
