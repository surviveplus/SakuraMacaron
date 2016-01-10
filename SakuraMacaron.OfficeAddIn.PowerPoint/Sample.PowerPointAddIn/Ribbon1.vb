Imports Microsoft.Office.Tools.Ribbon
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.PowerPoint
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.UI

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        If Me.samplePane Is Nothing Then

            Dim c = New UserControl1
            AddHandler c.DoButtonClick,
                Sub(sender2, e2)

                    'ThisAddIn.Current.Application.ActiveWindow.Selection.TextRange.Text = "sample"

                    Dim macaron = New PowerPointMacaron(ThisAddIn.Current.Application)
                    macaron.ReplaceSelectionParagraphs(
                        Sub(a)
                            a.IsCanceled = a.Text.Contains("X")
                        End Sub,
                        Sub(a)
                            a.IsSkipped = a.Text = String.Empty OrElse (a.Text.StartsWith("[[") AndAlso a.Text.EndsWith("]]"))
                            a.InsertBeforeText = "[[ "
                            a.InsertAfterText = " ]]"
                        End Sub
                    )

                    macaron.ReplaceSelectionText(
                        Sub(a)
                            a.IsCanceled = a.Text.Contains("X")
                        End Sub,
                        Sub(a)
                            a.IsSkipped = a.Text = String.Empty OrElse (a.Text.StartsWith("START") AndAlso a.Text.EndsWith("END"))
                            a.InsertBeforeText = "START" & vbCrLf
                            a.InsertAfterText = vbCrLf & "END"
                        End Sub
                    )

                End Sub

            Me.samplePane = New ElementControlPane(Of UserControl1)(c)
            Me.samplePane.Pane = ThisAddIn.Current.CustomTaskPanes.Add(Me.samplePane.Control, "Sample Pane", ThisAddIn.Current.Application.ActiveWindow)
            Me.samplePane.Pane.Width = 350

        End If

        Me.samplePane?.Show()

    End Sub

    Private samplePane As ElementControlPane(Of UserControl1)

End Class
