Imports Microsoft.VisualBasic
Imports System
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Fields

Namespace DocumentExplorerExample
	' Classes inherited from the Item class provide specialized representation of particular 
	' Document nodes by overriding virtual methods and properties of the base class.

	Public Class DocumentItem
		Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property IsRemovable() As Boolean
            Get
                Return False
            End Get
        End Property

    End Class

    Public Class SectionItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class HeaderFooterItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Protected Overrides ReadOnly Property IconName() As String
            Get
                If (CType(Node, HeaderFooter)).IsHeader Then
                    Return "Header"
                Else
                    Return "Footer"
                End If
            End Get
        End Property

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - {1}", MyBase.Name, (CType(Node, HeaderFooter)).HeaderFooterType.ToString())
            End Get
        End Property
    End Class

    Public Class BodyItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class TableItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class RowItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class CellItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class ParagraphItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property IsRemovable() As Boolean
            Get
                Dim para As Paragraph = CType(Node, Paragraph)
                Return Not para.IsEndOfSection
            End Get
        End Property
    End Class

    Public Class RunItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class FieldStartItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class FieldSeparatorItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class FieldEndItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class BookmarkStartItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - ""{1}""", MyBase.Name, (CType(Node, BookmarkStart)).Name)
            End Get
        End Property
    End Class

    Public Class BookmarkEndItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - ""{1}""", MyBase.Name, (CType(Node, BookmarkEnd)).Name)
            End Get
        End Property
    End Class

    Public Class CommentRangeStartItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - (Id = {1})", MyBase.Name, (CType(Node, CommentRangeStart)).Id)
            End Get
        End Property
    End Class

    Public Class CommentRangeEndItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - (Id = {1})", MyBase.Name, (CType(Node, CommentRangeEnd)).Id)
            End Get
        End Property
    End Class

    Public Class CommentItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - (Id = {1})", MyBase.Name, (CType(Node, Comment)).Id)
            End Get
        End Property
    End Class
    Public Class FootnoteItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class DrawingMLItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class StructuredDocumentTagItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class CustomXmlMarkupItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class OfficeMathItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class SmartTagItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class ShapeItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Dim shape As Shape = CType(Node, Shape)
                Select Case shape.ShapeType
                    Case ShapeType.OleObject, ShapeType.OleControl
                        Return shape.OleFormat.ProgId
                    Case Else
                        Return MyBase.IconName
                End Select
            End Get
        End Property

        Protected Overrides ReadOnly Property IconName() As String
            Get
                Dim shape As Shape = CType(Node, Shape)
                Select Case shape.ShapeType
                    Case ShapeType.OleObject
                        Return "OleObject"
                    Case ShapeType.OleControl
                        Return "OleControl"
                    Case Else
                        If shape.IsInline Then
                            Return "InlineShape"
                        Else
                            Return MyBase.IconName
                        End If
                End Select
            End Get
        End Property

    End Class

    Public Class GroupShapeItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class

    Public Class FormFieldItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub

        Public Overrides ReadOnly Property Name() As String
            Get
                Return String.Format("{0} - ""{1}""", MyBase.Name, (CType(Node, FormField)).Name)
            End Get
        End Property

        Protected Overrides ReadOnly Property IconName() As String
            Get
                Select Case (CType(Node, FormField)).Type
                    Case FieldType.FieldFormCheckBox
                        Return "FormCheckBox"
                    Case FieldType.FieldFormDropDown
                        Return "FormDropDown"
                    Case FieldType.FieldFormTextInput
                        Return "FormTextInput"
                    Case Else
                        Return MyBase.IconName
                End Select
            End Get
        End Property

    End Class

    Public Class SpecialCharItem
        Inherits Item
        Public Sub New(ByVal node As Node)
            MyBase.New(node)
        End Sub
    End Class
End Namespace