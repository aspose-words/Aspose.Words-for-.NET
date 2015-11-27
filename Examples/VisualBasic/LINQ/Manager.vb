'///////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'///////////////////////////////////////////////////////////////////////
Imports System.Collections.Generic
Imports System.Text

Namespace LINQ
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
End Namespace