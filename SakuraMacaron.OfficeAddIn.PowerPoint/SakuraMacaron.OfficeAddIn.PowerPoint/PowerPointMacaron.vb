Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Net.Surviveplus.SakuraMacaron.Core

Public Class PowerPointMacaron
    Inherits Macaron

#Region " Macaron members "
    Public Overrides Sub ReplaceSelectionParagraphs(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        If prepare IsNot Nothing Then
            For Each r In (
                    From s In Me.SelectedTextRanges
                    From p As TextRange2 In s.Paragraphs
                    Where String.IsNullOrWhiteSpace(p.Text) = False
                    Select p).ToList()

                Dim a As New TextActionsParameters With {.Text = r.Text}
                prepare(a)
                If a.IsCanceled Then Exit Sub
            Next r
        End If

        For Each r In (
                    From s In Me.SelectedTextRanges
                    From p As TextRange2 In s.Paragraphs
                    Where String.IsNullOrWhiteSpace(p.Text) = False
                    Select p).ToList()

            Dim a As New TextActionsParameters With {.Text = r.Text}
            act(a)
            If a.IsCanceled Then Exit Sub
            If a.IsSkipped = False Then

                If String.IsNullOrEmpty(a.InsertAfterText) = False AndAlso r.Text.EndsWith(vbCr) Then
                    r.Text = a.InsertBeforeText & Strings.Left(a.Text, a.Text.Length - 1) & a.InsertAfterText
                Else
                    If r.Text <> a.Text Then r.Text = a.Text
                    If String.IsNullOrEmpty(a.InsertBeforeText) = False Then r.InsertBefore(a.InsertBeforeText)
                    If String.IsNullOrEmpty(a.InsertAfterText) = False Then r.InsertAfter(a.InsertAfterText)
                End If
            End If
        Next r

    End Sub

    Public Overrides Sub ReplaceSelectionText(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        If prepare IsNot Nothing Then
            For Each r In Me.SelectedTextRanges
                Dim a As New TextActionsParameters With {.Text = r.Text}
                prepare(a)
                If a.IsCanceled Then Exit Sub
            Next r
        End If

        For Each r In Me.SelectedTextRanges
            Dim a As New TextActionsParameters With {.Text = r.Text}
            act(a)
            If a.IsCanceled Then Exit Sub
            If a.IsSkipped = False Then
                If r.Text <> a.Text Then r.Text = a.Text
                If String.IsNullOrEmpty(a.InsertBeforeText) = False Then r.InsertBefore(a.InsertBeforeText)
                If String.IsNullOrEmpty(a.InsertAfterText) = False Then r.InsertAfter(a.InsertAfterText)
            End If
        Next r


    End Sub

#End Region

    ' Class members

#Region " Constructors "


    ''' <summary>
    ''' クラスの新しいインスタンスを初期化します。
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(app As Application)
        If app Is Nothing Then Throw New ArgumentNullException("app")
        Me.app = app
    End Sub

    Private app As Application

#End Region

#Region " Properties "
    Public ReadOnly Property SelectedTextRanges As IEnumerable(Of TextRange2)
        Get
            Dim r As TextRange2 = Nothing
            Try
                r = Me.app.ActiveWindow.Selection.TextRange2
            Catch ex As Exception
            End Try

            If r IsNot Nothing Then
                Select Case Me.app.ActiveWindow.Selection.Type
                    Case PpSelectionType.ppSelectionText
                        Return {r}
                    Case PpSelectionType.ppSelectionShapes
                        Return (From t As TextRange2 In r).ToList()
                    Case Else
                        Return New TextRange2() {}
                End Select
            Else
                Return New TextRange2() {}
            End If
        End Get
    End Property
#End Region



End Class
