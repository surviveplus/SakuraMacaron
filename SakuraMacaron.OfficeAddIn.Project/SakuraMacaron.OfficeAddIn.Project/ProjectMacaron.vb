Imports Microsoft.Office.Interop.MSProject
Imports Net.Surviveplus.SakuraMacaron.Core

Public Class ProjectMacaron
    Inherits Macaron

    ' Overrides or implements

#Region " Macaron members "

    Public Overrides Sub ReplaceSelectionParagraphs(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))
        Throw New NotImplementedException()
    End Sub

    Public Overrides Sub ReplaceSelectionText(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        If Me.app.ActiveSelection.Tasks Is Nothing Then Exit Sub

        Dim getFieldValue =
            Function(task As Task, fieldId As String) As String
                Try
                    Dim value = task.GetField(fieldId)
                    Return If(value IsNot Nothing, value, String.Empty)

                Catch ex As Runtime.InteropServices.COMException
                    Return Nothing
                End Try
            End Function

        Dim target =
            From task As Task In Me.app.ActiveSelection.Tasks
            From fieldId As String In Me.app.ActiveSelection.FieldIDList
            Let value = getFieldValue(task, fieldId)
            Where value IsNot Nothing
            Select New With {.Task = task, .FieldId = fieldId, .Value = value}


        If prepare IsNot Nothing Then
            For Each t In target
                Dim a As New TextActionsParameters With {.Text = t.Value}
                prepare(a)
                If a.IsCanceled Then Exit Sub
            Next
        End If

        For Each t In target
            Dim a As New TextActionsParameters With {.Text = t.Value}
            act(a)
            If a.IsCanceled Then Exit Sub
            If a.IsSkipped = False AndAlso
                (a.Text <> t.Value OrElse
                    String.IsNullOrEmpty(a.InsertBeforeText) = False OrElse
                    String.IsNullOrEmpty(a.InsertAfterText) = False) Then

                t.Task.SetField(
                    t.FieldId,
                    If(String.IsNullOrEmpty(a.InsertBeforeText) = False, a.InsertBeforeText, "") &
                        a.Text &
                        If(String.IsNullOrEmpty(a.InsertAfterText) = False, a.InsertAfterText, ""))
            End If
        Next

    End Sub
#End Region

    ' Class members

#Region " Constructors "

    ''' <summary>
    ''' Initializes a new instance of the class.
    ''' </summary>
    ''' <param name="app">Set Application object.</param>
    ''' <remarks></remarks>
    Public Sub New(app As Application)
        If app Is Nothing Then Throw New ArgumentNullException("app")
        Me.app = app
    End Sub

    Private app As Application

#End Region



End Class
