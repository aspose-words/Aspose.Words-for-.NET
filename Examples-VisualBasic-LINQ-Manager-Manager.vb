' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Public Class Manager
    Public Property Name() As [String]
        Get
            Return m_Name
        End Get
        Set(value As [String])
            m_Name = value
        End Set
    End Property
    Private m_Name As [String]
    Public Property Age() As Integer
        Get
            Return m_Age
        End Get
        Set(value As Integer)
            m_Age = value
        End Set
    End Property
    Private m_Age As Integer
    Public Property Photo() As Byte()
        Get
            Return m_Photo
        End Get
        Set(value As Byte())
            m_Photo = value
        End Set
    End Property
    Private m_Photo As Byte()
    Public Property Contracts() As IEnumerable(Of Contract)
        Get
            Return m_Contracts
        End Get
        Set(value As IEnumerable(Of Contract))
            m_Contracts = value
        End Set
    End Property
    Private m_Contracts As IEnumerable(Of Contract)
End Class
