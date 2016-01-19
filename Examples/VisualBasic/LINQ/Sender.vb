
Imports System.Collections.Generic
Imports System.Text

Namespace LINQ
    Public Class Sender
        Public Property Name() As [String]
            Get
                Return m_Name
            End Get
            Set(value As [String])
                m_Name = value
            End Set
        End Property
        Private m_Name As [String]
        Public Property Message() As [String]
            Get
                Return m_Message
            End Get
            Set(value As [String])
                m_Message = value
            End Set
        End Property
        Private m_Message As [String]
    End Class
End Namespace