Imports Microsoft.Office.Interop.Word
Imports Net.Surviveplus.SakuraMacaron.Core

Public Class WordMacaron
    Inherits Macaron

    ' Overrides or implements

#Region " Macaron members "

    Public Overrides Sub ReplaceSelectionParagraphs(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        If prepare IsNot Nothing Then
            For Each r In Me.SelectedParagraphsRanges

                If r.Text?.EndsWith(vbCr) Then
                    r.SetRange(r.Start, r.End - 1)
                End If

                Dim text = r.Text
                If text Is Nothing Then text = String.Empty

                Dim a As New TextActionsParameters With {.text = text}
                prepare(a)
                If a.IsCanceled Then Exit Sub
            Next
        End If

        For Each r In Me.SelectedParagraphsRanges

            If r.Text?.EndsWith(vbCr) Then
                r.SetRange(r.Start, r.End - 1)
            End If

            Dim text = r.Text
            If text Is Nothing Then text = String.Empty

            Dim a As New TextActionsParameters With {.text = text}
            act(a)
            If a.IsCanceled Then Exit Sub
            If a.IsSkipped = False Then
                If text <> a.Text Then r.Text = a.Text
                If String.IsNullOrEmpty(a.InsertBeforeText) = False Then r.InsertBefore(a.InsertBeforeText)
                If String.IsNullOrEmpty(a.InsertAfterText) = False Then r.InsertAfter(a.InsertAfterText)
            End If

        Next

    End Sub

    Public Overrides Sub ReplaceSelectionText(prepare As Action(Of TextActionsParameters), act As Action(Of TextActionsParameters))

        If prepare IsNot Nothing Then
            For Each r In Me.SelectedTextRanges

                If r.Text?.EndsWith(vbCr) Then
                    r.SetRange(r.Start, r.End - 1)
                End If

                Dim text = r.Text
                If text Is Nothing Then text = String.Empty

                Dim a As New TextActionsParameters With {.text = text}
                prepare(a)
                If a.IsCanceled Then Exit Sub
            Next
        End If

        For Each r In Me.SelectedTextRanges

            If r.Text?.EndsWith(vbCr) Then
                r.SetRange(r.Start, r.End - 1)
            End If

            Dim text = r.Text
            If text Is Nothing Then text = String.Empty

            Dim a As New TextActionsParameters With {.text = text}
            act(a)
            If a.IsCanceled Then Exit Sub
            If a.IsSkipped = False Then
                If text <> a.Text Then r.Text = a.Text
                If String.IsNullOrEmpty(a.InsertBeforeText) = False Then r.InsertBefore(a.InsertBeforeText)
                If String.IsNullOrEmpty(a.InsertAfterText) = False Then r.InsertAfter(a.InsertAfterText)
            End If

        Next
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
    Protected Sub New()
    End Sub

#End Region

    Protected Overridable Function GetSelection() As Selection
        Return Me.app.Selection
    End Function

#Region " Properties "

    Public ReadOnly Property SelectedTextRanges As IEnumerable(Of Range)
        Get
            Dim r = (From shape As Shape In Me.GetSelection()?.ShapeRange Select shape.TextFrame?.TextRange).ToArray()
            If r.Count = 0 Then r = {Me.GetSelection()?.Range}

            Return r
        End Get
    End Property

    Public ReadOnly Property SelectedParagraphsRanges As IEnumerable(Of Range)
        Get
            Dim r = (
                From shape As Shape In (From s In Me.GetSelection()?.ShapeRange).ToArray()
                From line As Range In (
                        From p As Paragraph In shape.TextFrame.TextRange.Paragraphs
                        Select p.Range).ToArray()
                Select line).ToArray()

            If r.Count = 0 Then r = (From p As Paragraph In Me.GetSelection()?.Paragraphs Select p.Range).ToArray()

            Return r
        End Get
    End Property
#End Region


End Class
