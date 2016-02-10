Imports Microsoft.Office.Interop.Excel
Imports Net.Surviveplus.SakuraMacaron.Core

Public Class ExcelMacaron
    Inherits Macaron

    ' Overrides or implements

#Region " Macaron members "

    Public Overrides Sub ReplaceSelectionParagraphs(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        Throw New NotImplementedException()

    End Sub

    Public Overrides Sub ReplaceSelectionText(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        Dim target As Range = Me.app.Selection

        Dim getText =
            Function(cell As Range) As String
                Return cell.Formula
            End Function

        Dim setText =
            Sub(cell As Range, a As TextActionsParameters)
                cell.Formula = a.InsertBeforeText & a.Text & a.InsertAfterText
            End Sub

        Dim r = ForEachRange(target, prepare, getText, Nothing)
        If r.IsCanceld = False Then ForEachRange(target, act, getText, setText)

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

    Private Shared Function ForEachRange(
        target As Range,
        act As Action(Of TextActionsParameters),
        getText As Func(Of Range, String),
        setText As Action(Of Range, TextActionsParameters)
        ) As ExecuteSelectionResult

        If target Is Nothing Then Throw New ArgumentNullException("target")
        If act Is Nothing Then Return New ExecuteSelectionResult With {.HasNoAction = True}

        If target.Count = 1 Then
            Dim a As New TextActionsParameters
            Dim text = getText(target)
            a.Text = text
            act(a)
            If a.IsCanceled Then Return New ExecuteSelectionResult With {.IsCanceld = True}
            If setText IsNot Nothing AndAlso a.IsSkipped = False Then
                If text <> a.Text OrElse
                            String.IsNullOrEmpty(a.InsertBeforeText) = False OrElse
                            String.IsNullOrEmpty(a.InsertAfterText) = False Then

                    setText(target, a)
                End If
            End If
        Else
            If target.Rows.Count * target.Columns.Count <> target.Count Then

                Dim c = 1
                For Each item As Range In target
                    Dim a As New TextActionsParameters With {
                        .IsBox = False,
                        .RowIndex = 1, .ColumnIndex = c
                    }
                    Dim text = getText(item)
                    a.Text = text
                    act(a)
                    If a.IsCanceled Then Return New ExecuteSelectionResult With {.IsCanceld = True}
                    If setText IsNot Nothing AndAlso a.IsSkipped = False Then
                        If text <> a.Text OrElse
                            String.IsNullOrEmpty(a.InsertBeforeText) = False OrElse
                            String.IsNullOrEmpty(a.InsertAfterText) = False Then

                            setText(item, a)
                        End If
                    End If
                    c += 1
                Next item
            Else
                For r = 1 To target.Rows.Count
                    For c = 1 To target.Columns.Count
                        Dim item As Range = target.Cells(r, c)

                        Dim a As New TextActionsParameters With {
                            .IsBox = True,
                            .RowIndex = r, .ColumnIndex = c}
                        Dim text = getText(item)
                        a.Text = text
                        act(a)
                        If a.IsCanceled Then Return New ExecuteSelectionResult With {.IsCanceld = True}
                        If setText IsNot Nothing AndAlso a.IsSkipped = False Then
                            If text <> a.Text OrElse
                            String.IsNullOrEmpty(a.InsertBeforeText) = False OrElse
                            String.IsNullOrEmpty(a.InsertAfterText) = False Then

                                setText(item, a)
                            End If
                        End If

                    Next c
                Next r
            End If
        End If

        Return New ExecuteSelectionResult
    End Function

End Class


Friend Class ExecuteSelectionResult
    Public Property IsCanceld As Boolean

    Public Property HasNoAction As Boolean
End Class