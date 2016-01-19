Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Drawing
Imports System.Reflection
Imports System.Collections
Imports System.Text
Imports System.Windows.Forms
Imports Aspose.Words

Namespace DocumentExplorerExample
    ''' <summary>
    ''' Base class to provide GUI representation for document nodes.
    ''' </summary>
    Public Class Item
        ''' <summary>
        ''' Creates Item for the document node.
        ''' </summary>
        ''' <param name="aNode">Document node which this Item will represent.</param>
        Public Sub New(ByVal aNode As Node)
            mNode = aNode
        End Sub

        ''' <summary>
        ''' Document node that this Item represents.
        ''' </summary>
        Public ReadOnly Property Node() As Node
            Get
                Return mNode
            End Get
        End Property

        ''' <summary>
        '''  DisplayName for this Item. Can be customized by overriding in inheriting classes.
        ''' </summary>
        Public Overridable ReadOnly Property Name() As String
            Get
                Return mNode.NodeType.ToString()
            End Get
        End Property

        ''' <summary>
        ''' Text contained by the corresponding document node.
        ''' </summary>
        Public ReadOnly Property Text() As String
            Get
                Dim result As New StringBuilder()

                ' All control characters are converted to human readable form.
                ' E.g. [!PageBreak!], [!ParagraphBreak!], etc.
                For Each c As Char In mNode.GetText()
                    Dim controlCharDisplay As String = CStr(gControlCharacters(c))
                    If controlCharDisplay Is Nothing Then
                        result.Append(c)
                    Else
                        result.Append(controlCharDisplay)
                    End If
                Next c

                Return result.ToString()
            End Get
        End Property

        ''' <summary>
        ''' Creates TreeNode for this item to be displayed in Document Explorer TreeView control.
        ''' </summary>
        Public ReadOnly Property TreeNode() As TreeNode
            Get
                If mTreeNode Is Nothing Then
                    mTreeNode = New TreeNode(Name)
                    If (Not gIconNames.Contains(IconName)) Then
                        gIconNames.Add(IconName)
                        ImageList.Images.Add(Icon)
                    End If
                    Dim index As Integer = gIconNames.IndexOf(IconName)
                    mTreeNode.ImageIndex = index
                    mTreeNode.SelectedImageIndex = index
                    mTreeNode.Tag = Me
                    If TypeOf mNode Is CompositeNode AndAlso (CType(mNode, CompositeNode)).ChildNodes.Count > 0 Then
                        mTreeNode.Nodes.Add("#dummy")
                    End If
                End If
                Return mTreeNode
            End Get
        End Property

        Public Shared ReadOnly Property ImageList() As ImageList
            Get
                If mImageList Is Nothing Then
                    mImageList = New ImageList()
                    mImageList.ColorDepth = ColorDepth.Depth32Bit
                    mImageList.ImageSize = New Size(16, 16)
                End If
                Return mImageList
            End Get
        End Property

        ''' <summary>
        ''' Icon to display in the Document Explorer TreeView control.
        ''' </summary>
        Public ReadOnly Property Icon() As Icon
            Get
                If mIcon Is Nothing Then
                    mIcon = LoadIcon(IconName)
                    If mIcon Is Nothing Then
                        mIcon = LoadIcon("Node")
                    End If
                End If
                Return mIcon
            End Get
        End Property

        ''' <summary>
        ''' Icon for this node can be customized by overriding this property in the inheriting classes.
        ''' The name represents name of .ico file without extension located in the Icons folder of the project.
        ''' </summary>
        Protected Overridable ReadOnly Property IconName() As String
            Get
                Return Me.GetType().Name.Replace("Item", "")
            End Get
        End Property

        ''' <summary>
        ''' Provides lazy on-expand loading of underlying tree nodes.
        ''' </summary>
        Public Sub OnExpand()
            ' Optimized to allow automatic conversion to VB.NET
            If TreeNode.Nodes(0).Text.Equals("#dummy") Then
                TreeNode.Nodes.Clear()
                For Each n As Node In (CType(mNode, CompositeNode)).ChildNodes
                    TreeNode.Nodes.Add(CreateItem(n).TreeNode)
                Next n
            End If
        End Sub

        ''' <summary>
        ''' Loads icon from assembly resource stream.
        ''' </summary>
        ''' <param name="anIconName">Name of the icon to load.</param>
        ''' <returns>Icon object or null if icon was not found in the resources.</returns>
        Private Shared Function LoadIcon(ByVal anIconName As String) As Icon
            Dim resourceName As String = "VisualBasic.Viewers_Visualizers.Document_Explorer.Icons." & anIconName & ".ico"
            Dim iconStream As Stream = FetchResourceStream(resourceName)

            If iconStream IsNot Nothing Then
                Return New Icon(iconStream)
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' Returns a resource stream from the executing assembly or throws if the resource cannot be found.
        ''' </summary>
        ''' <param name="resourceName">The name of the resource without the name of the assembly.</param>
        ''' <returns>The stream. Don't forget to close it when finished.</returns>
        Friend Shared Function FetchResourceStream(ByVal resourceName As String) As Stream
            Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
            Dim fullName As String = String.Format("{0}Example.{1}", asm.GetName().Name, resourceName)
            Dim stream As Stream = asm.GetManifestResourceStream(fullName)

            ' Ugly optimization so conversion to VB.NET can work.
            Do While stream Is Nothing
                Dim dotPos As Integer = fullName.IndexOf(".")
                If dotPos < 0 Then
                    Return Nothing
                End If

                fullName = fullName.Substring(dotPos + 1)
                stream = asm.GetManifestResourceStream(fullName)
            Loop

            Return stream
        End Function

        Public Sub Remove()
            If IsRemovable Then
                mNode.Remove()
                mTreeNode.Remove()
            End If
        End Sub

        Public Overridable ReadOnly Property IsRemovable() As Boolean
            Get
                Return True
            End Get
        End Property

        ''' <summary>
        ''' Static ctor.
        ''' </summary>
        Shared Sub New()
            ' Fill set of typenames of Item inheritors for Item class fabric.
            gItemSet = New ArrayList()
            For Each type As Type In System.Reflection.Assembly.GetExecutingAssembly().GetTypes()
                If type.IsSubclassOf(GetType(Item)) AndAlso (Not type.IsAbstract) Then
                    gItemSet.Add(type.Name)
                End If
            Next type

            ' Fill control chars fields set
            gControlCharacters.Add(ControlChar.CellChar, "[!Cell!]")
            gControlCharacters.Add(ControlChar.ColumnBreakChar, "[!ColumnBreak!]" & Constants.vbCrLf)
            gControlCharacters.Add(ControlChar.FieldEndChar, "[!FieldEnd!]")
            gControlCharacters.Add(ControlChar.FieldSeparatorChar, "[!FieldSeparator!]")
            gControlCharacters.Add(ControlChar.FieldStartChar, "[!FieldStart!]")
            gControlCharacters.Add(ControlChar.LineBreakChar, "[!LineBreak!]" & Constants.vbCrLf)
            gControlCharacters.Add(ControlChar.LineFeedChar, "[!LineFeed!]")
            gControlCharacters.Add(ControlChar.NonBreakingHyphenChar, "[!NonBreakingHyphen!]")
            gControlCharacters.Add(ControlChar.NonBreakingSpaceChar, "[!NonBreakingSpace!]")
            gControlCharacters.Add(ControlChar.OptionalHyphenChar, "[!OptionalHyphen!]")
            gControlCharacters.Add(ControlChar.ParagraphBreakChar, "¶" & Constants.vbCrLf)
            gControlCharacters.Add(ControlChar.SectionBreakChar, "[!SectionBreak!]" & Constants.vbCrLf)
            gControlCharacters.Add(ControlChar.TabChar, "[!Tab!]")
        End Sub

        ''' <summary>
        ''' Item class factory implementation.
        ''' </summary>
        Public Shared Function CreateItem(ByVal aNode As Node) As Item
            Dim typeName As String = aNode.NodeType.ToString() & "Item"
            If gItemSet.Contains(typeName) Then
                Return CType(Activator.CreateInstance(Type.GetType("DocumentExplorerExample." & typeName), New Object() {aNode}), Item)
            Else
                Return New Item(aNode)
            End If
        End Function

        Private ReadOnly mNode As Node
        Private mTreeNode As TreeNode
        Private Shared mImageList As ImageList
        Private mIcon As Icon

        Private Shared ReadOnly gItemSet As ArrayList
        Private Shared ReadOnly gIconNames As New ArrayList()
        ''' <summary>
        ''' Map of character to string that we use to display control MS Word control characters.
        ''' </summary>
        Private Shared ReadOnly gControlCharacters As New Hashtable()
    End Class
End Namespace