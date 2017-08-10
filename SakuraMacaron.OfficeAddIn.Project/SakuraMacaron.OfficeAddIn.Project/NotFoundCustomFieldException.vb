''' <summary>
'''The exception that is thrown when problem occurs.
''' </summary>
''' <remarks>
''' <para>change log</para>
''' </remarks>
<Serializable()> Public Class NotFoundCustomFieldException
    Inherits Exception

    ' Override or Implement Interface

#Region " Exception Members "

    ''' <summary>
    ''' Sets the SerializationInfo object with the parameter name and additional exception information.
    ''' </summary>
    ''' <param name="info">The object that holds the serialized object data.</param>
    ''' <param name="context">The contextual information about the source or destination.</param>
    ''' <remarks></remarks>
    ''' <exception cref="System.ArgumentNullException">The info object is Nothing (a null reference in C#).</exception>
    Public Overrides Sub GetObjectData(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.GetObjectData(info, context)

        ' TODO: Save private members in info object. For example to save Value property, write the code as follows.
        'info.AddValue("Value", Me.Value)
    End Sub

#End Region

    ' Class Members

#Region " Constructors "

    ''' <summary>
    ''' Initializes a new instance of the NotFoundCustomFieldException class.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the NotFoundCustomFieldException class with a specified error message. 
    ''' </summary>
    ''' <param name="message">The error message that explains the reason for the exception.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String)
        MyBase.New(message)
    End Sub

    ''' <summary>
    ''' Initializes a new instance of the NotFoundCustomFieldException class with a specified error message and a reference to the inner exception that is the cause of this exception.
    ''' </summary>
    ''' <param name="message">The error message that explains the reason for the exception.</param>
    ''' <param name="innerException">The exception that is the cause of the current exception. Set Nothing (a null reference in C#) if the inner exception value was not supplied.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal message As String, ByVal innerException As Exception)
        MyBase.New(message, innerException)
    End Sub

    ''' <summary>
    '''  Initializes a new instance of the NotFoundCustomFieldException class with serialized data.
    ''' </summary>
    ''' <param name="info">The object that holds the serialized object data.</param>
    ''' <param name="context">The contextual information about the source or destination.</param>
    ''' <remarks></remarks>
    Public Sub New(ByVal info As System.Runtime.Serialization.SerializationInfo, ByVal context As System.Runtime.Serialization.StreamingContext)
        MyBase.New(info, context)

        ' TODO: Initialize private members by info object. For example to initialize Value property, write the code as follows.
        ' Me.Value = info.GetInt32("Value")

    End Sub
#End Region

End Class