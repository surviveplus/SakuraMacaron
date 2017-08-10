Imports Microsoft.Office.Interop.MSProject

Module TaskExtensions

    ''' <summary>
    ''' Do something, and return something.
    ''' </summary>
    ''' <param name="this">The instance of the type which is added this extension method.</param>
    ''' <param name="value">Set something.</param>
    ''' <remarks></remarks>
    <Runtime.CompilerServices.Extension()>
    Public Sub SetCustomFieldValue(ByVal this As Task, ByVal name As String, ByVal value As String)
        If this Is Nothing Then Throw New ArgumentNullException("this")

        Dim app = this.Application

        For index = 1 To 30
            Dim originalName = "Text" & index
            Dim id = app.FieldNameToFieldConstant(originalName)
            Dim customName = app.CustomFieldGetName(id)

            If customName = name Then
                this.SetField(id, value)
                Return
            End If
        Next index

        ' VSTS/TFS に接続した Project を開くと、Iteration Path は OutlineCode10
        ' OutlineCode1~10 では、選択肢に値がないときはセットに失敗する様子

        For index = 1 To 10
            Dim originalName = "OutlineCode" & index
            Dim id = app.FieldNameToFieldConstant(originalName)
            Dim customName = app.CustomFieldGetName(id)

            If customName = name Then
                this.SetField(id, value)
                Return
            End If
        Next index

        Throw New NotFoundCustomFieldException(name & " is not found.")
    End Sub

End Module

