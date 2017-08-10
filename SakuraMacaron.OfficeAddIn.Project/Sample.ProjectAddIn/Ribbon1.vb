Imports System.Diagnostics
Imports Microsoft.Office.Interop.MSProject
Imports Microsoft.Office.Tools.Ribbon
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.Project
Imports Net.Surviveplus.SakuraMacaron.OfficeAddIn.UI

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        Dim app = ThisAddIn.Current.Application

        Dim macaron = New ProjectMacaron(ThisAddIn.Current.Application)

        macaron.ReplaceSelectionText(
            Nothing,
            Sub(a)
                a.Text = a.Text & " Hello!!"
            End Sub)

    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click

        If Me.sampleTool Is Nothing Then
            Dim c = New UserControl1

            AddHandler c.DoButtonClick,
                Sub(sender2, e2)

                    Dim app = ThisAddIn.Current.Application

                    Dim macaron = New ProjectMacaron(ThisAddIn.Current.Application)

                    macaron.ReplaceSelectionText(
                        Nothing,
                        Sub(a)
                            a.Text = a.Text & " Good!!"
                        End Sub)


                End Sub

            Me.sampleTool = New ElementControlToolWindow(Of UserControl1)(c)

        End If
        Me.sampleTool?.Show()
    End Sub


    Private sampleTool As ElementControlToolWindow(Of UserControl1)

End Class
