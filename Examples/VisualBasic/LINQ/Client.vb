
Imports System.Collections.Generic
Imports System.Text

Namespace LINQ
    Public Class Client
        Public Property Name() As [String]
            Get
                Return m_Name
            End Get
            Set(value As [String])
                m_Name = value
            End Set
        End Property
        Private m_Name As [String]
    End Class
End Namespace
