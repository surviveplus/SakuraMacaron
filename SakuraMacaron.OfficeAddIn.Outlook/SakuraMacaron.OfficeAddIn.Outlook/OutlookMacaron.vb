Imports Microsoft.Office.Interop.Outlook
Imports Net.Surviveplus.SakuraMacaron.Core

Public Class OutlookMacaron
    Inherits Net.Surviveplus.SakuraMacaron.OfficeAddIn.Word.WordMacaron

    ' Overrides or implements

#Region " WordMacaron members "

    Protected Overrides Function GetSelection() As Microsoft.Office.Interop.Word.Selection
        Return Me.WordSelection

    End Function

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

    Public ReadOnly Property WordSelection As Microsoft.Office.Interop.Word.Selection
        Get

            If Me.app.ActiveInspector.IsWordMail AndAlso
                Me.app.ActiveInspector.EditorType = OlEditorType.olEditorWord Then

                Dim word = Me.app.ActiveInspector.WordEditor
                Dim doc As Microsoft.Office.Interop.Word.Document = word
                Return doc.Windows(1).Selection
            Else
                Return Nothing
            End If
        End Get
    End Property


End Class
