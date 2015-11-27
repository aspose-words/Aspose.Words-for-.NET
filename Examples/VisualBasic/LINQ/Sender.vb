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