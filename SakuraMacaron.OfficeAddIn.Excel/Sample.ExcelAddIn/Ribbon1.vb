﻿Imports Microsoft.Office.Tools.Ribbon
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.Excel
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.UI

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        If Me.samplePane Is Nothing Then
            Dim c = New UserControl1

            AddHandler c.DoButtonClick,
                Sub(sender2, e2)

                    Dim macaron = New ExcelMacaron(ThisAddIn.Current.Application)
                    macaron.ReplaceSelectionText(
                        Sub(a)

                        End Sub,
                        Sub(a)
                            a.Text = a.Text & "TEST"
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