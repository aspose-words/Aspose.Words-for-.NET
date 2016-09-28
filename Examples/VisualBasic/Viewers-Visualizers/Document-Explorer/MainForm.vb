Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Windows.Forms
Imports Aspose.Words
Imports Aspose.Words.Saving

Namespace DocumentExplorerExample
    ''' <summary>
    ''' The main form of the DocumentExplorer demo.
    ''' 
    ''' DocumentExplorer allows to open documents using Aspose.Words.
    ''' Once a document is opened, you can explore its object model in the tree.
    ''' You can also save the document into DOC, DOCX, ODF, EPUB, PDF, SWF, RTF, WordML,
    ''' HTML, MHTML and plain text formats.
    ''' </summary>
    Public Class MainForm
        Inherits Form

#Region "Windows Form Designer generated code"

        Private WithEvents toolBar1 As System.Windows.Forms.ToolBar
        Private imageList1 As System.Windows.Forms.ImageList
        Private panel1 As System.Windows.Forms.Panel
        Private splitter1 As System.Windows.Forms.Splitter
        Private panel2 As System.Windows.Forms.Panel
        Public StatusBar As System.Windows.Forms.StatusBar
        Public WithEvents Tree As System.Windows.Forms.TreeView
        Private toolOpenDocument As System.Windows.Forms.ToolBarButton
        Private toolSaveDocument As System.Windows.Forms.ToolBarButton
        Private toolExpandAll As System.Windows.Forms.ToolBarButton
        Private toolCollapseAll As System.Windows.Forms.ToolBarButton
        Private toolSeparator1 As System.Windows.Forms.ToolBarButton
        Public Text1 As System.Windows.Forms.TextBox
        Private mainMenu1 As System.Windows.Forms.MainMenu
        Private menuFile As System.Windows.Forms.MenuItem
        Private WithEvents menuOpen As System.Windows.Forms.MenuItem
        Private WithEvents menuSaveAs As System.Windows.Forms.MenuItem
        Private menuBar1 As System.Windows.Forms.MenuItem
        Private WithEvents menuExit As System.Windows.Forms.MenuItem
        Private menuView As System.Windows.Forms.MenuItem
        Private WithEvents menuExpandAll As System.Windows.Forms.MenuItem
        Private WithEvents menuCollapseAll As System.Windows.Forms.MenuItem
        Private menuHelp As System.Windows.Forms.MenuItem
        Private WithEvents menuAbout As System.Windows.Forms.MenuItem
        Private toolSeparator2 As System.Windows.Forms.ToolBarButton
        Private toolViewInWord As System.Windows.Forms.ToolBarButton
        Private toolViewInPdf As System.Windows.Forms.ToolBarButton
        Private toolRemove As System.Windows.Forms.ToolBarButton
        Private WithEvents menuRemoveNode As System.Windows.Forms.MenuItem
        Private menuEdit As System.Windows.Forms.MenuItem
        Private menuItem1 As System.Windows.Forms.MenuItem
        Private WithEvents menuRender As System.Windows.Forms.MenuItem
        Private toolRenderDocument As System.Windows.Forms.ToolBarButton
        Private toolBarButton1 As System.Windows.Forms.ToolBarButton
        Private toolPreviewButton As System.Windows.Forms.ToolBarButton
        Private WithEvents menuPreview As System.Windows.Forms.MenuItem
        Private components As System.ComponentModel.IContainer

        Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container()
            Dim resources As New System.Resources.ResourceManager(GetType(MainForm))
            Me.StatusBar = New System.Windows.Forms.StatusBar()
            Me.toolBar1 = New System.Windows.Forms.ToolBar()
            Me.toolOpenDocument = New System.Windows.Forms.ToolBarButton()
            Me.toolSaveDocument = New System.Windows.Forms.ToolBarButton()
            Me.toolSeparator1 = New System.Windows.Forms.ToolBarButton()
            Me.toolRenderDocument = New System.Windows.Forms.ToolBarButton()
            Me.toolPreviewButton = New System.Windows.Forms.ToolBarButton()
            Me.toolBarButton1 = New System.Windows.Forms.ToolBarButton()
            Me.toolExpandAll = New System.Windows.Forms.ToolBarButton()
            Me.toolCollapseAll = New System.Windows.Forms.ToolBarButton()
            Me.toolSeparator2 = New System.Windows.Forms.ToolBarButton()
            Me.toolRemove = New System.Windows.Forms.ToolBarButton()
            Me.toolViewInWord = New System.Windows.Forms.ToolBarButton()
            Me.toolViewInPdf = New System.Windows.Forms.ToolBarButton()
            Me.imageList1 = New System.Windows.Forms.ImageList(Me.components)
            Me.panel1 = New System.Windows.Forms.Panel()
            Me.Tree = New System.Windows.Forms.TreeView()
            Me.splitter1 = New System.Windows.Forms.Splitter()
            Me.panel2 = New System.Windows.Forms.Panel()
            Me.Text1 = New System.Windows.Forms.TextBox()
            Me.mainMenu1 = New System.Windows.Forms.MainMenu()
            Me.menuFile = New System.Windows.Forms.MenuItem()
            Me.menuOpen = New System.Windows.Forms.MenuItem()
            Me.menuSaveAs = New System.Windows.Forms.MenuItem()
            Me.menuBar1 = New System.Windows.Forms.MenuItem()
            Me.menuRender = New System.Windows.Forms.MenuItem()
            Me.menuPreview = New System.Windows.Forms.MenuItem()
            Me.menuItem1 = New System.Windows.Forms.MenuItem()
            Me.menuExit = New System.Windows.Forms.MenuItem()
            Me.menuEdit = New System.Windows.Forms.MenuItem()
            Me.menuRemoveNode = New System.Windows.Forms.MenuItem()
            Me.menuView = New System.Windows.Forms.MenuItem()
            Me.menuExpandAll = New System.Windows.Forms.MenuItem()
            Me.menuCollapseAll = New System.Windows.Forms.MenuItem()
            Me.menuHelp = New System.Windows.Forms.MenuItem()
            Me.menuAbout = New System.Windows.Forms.MenuItem()
            Me.panel1.SuspendLayout()
            Me.panel2.SuspendLayout()
            Me.SuspendLayout()
            ' 
            ' StatusBar
            ' 
            Me.StatusBar.Location = New System.Drawing.Point(0, 665)
            Me.StatusBar.Name = "StatusBar"
            Me.StatusBar.ShowPanels = True
            Me.StatusBar.Size = New System.Drawing.Size(884, 24)
            Me.StatusBar.TabIndex = 0
            Me.StatusBar.Text = "statusBar1"
            ' 
            ' toolBar1
            ' 
            Me.toolBar1.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
            Me.toolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.toolOpenDocument, Me.toolSaveDocument, Me.toolSeparator1, Me.toolRenderDocument, Me.toolPreviewButton, Me.toolBarButton1, Me.toolExpandAll, Me.toolCollapseAll, Me.toolSeparator2, Me.toolRemove, Me.toolViewInWord, Me.toolViewInPdf})
            Me.toolBar1.DropDownArrows = True
            Me.toolBar1.ImageList = Me.imageList1
            Me.toolBar1.Location = New System.Drawing.Point(0, 0)
            Me.toolBar1.Name = "toolBar1"
            Me.toolBar1.ShowToolTips = True
            Me.toolBar1.Size = New System.Drawing.Size(884, 28)
            Me.toolBar1.TabIndex = 1
            '			Me.toolBar1.ButtonClick += New System.Windows.Forms.ToolBarButtonClickEventHandler(Me.toolBar1_ButtonClick);
            ' 
            ' toolOpenDocument
            ' 
            Me.toolOpenDocument.ImageIndex = 0
            Me.toolOpenDocument.ToolTipText = "Open Document"
            ' 
            ' toolSaveDocument
            ' 
            Me.toolSaveDocument.Enabled = False
            Me.toolSaveDocument.ImageIndex = 1
            Me.toolSaveDocument.ToolTipText = "Save Document As..."
            ' 
            ' toolSeparator1
            ' 
            Me.toolSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
            ' 
            ' toolRenderDocument
            ' 
            Me.toolRenderDocument.Enabled = False
            Me.toolRenderDocument.ImageIndex = 7
            Me.toolRenderDocument.ToolTipText = "Render Document"
            ' 
            ' toolPreviewButton
            ' 
            Me.toolPreviewButton.Enabled = False
            Me.toolPreviewButton.ImageIndex = 8
            Me.toolPreviewButton.ToolTipText = "Print Preview"
            ' 
            ' toolBarButton1
            ' 
            Me.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
            ' 
            ' toolExpandAll
            ' 
            Me.toolExpandAll.Enabled = False
            Me.toolExpandAll.ImageIndex = 2
            Me.toolExpandAll.ToolTipText = "Expand All"
            ' 
            ' toolCollapseAll
            ' 
            Me.toolCollapseAll.Enabled = False
            Me.toolCollapseAll.ImageIndex = 3
            Me.toolCollapseAll.ToolTipText = "Collapse All"
            ' 
            ' toolSeparator2
            ' 
            Me.toolSeparator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
            ' 
            ' toolRemove
            ' 
            Me.toolRemove.Enabled = False
            Me.toolRemove.ImageIndex = 4
            Me.toolRemove.ToolTipText = "Remove Node"
            ' 
            ' toolViewInWord
            ' 
            Me.toolViewInWord.ImageIndex = 4
            Me.toolViewInWord.ToolTipText = "View in MS Word"
            Me.toolViewInWord.Visible = False
            ' 
            ' toolViewInPdf
            ' 
            Me.toolViewInPdf.ImageIndex = 5
            Me.toolViewInPdf.ToolTipText = "View in Acrobat"
            Me.toolViewInPdf.Visible = False
            ' 
            ' imageList1
            ' 
            Me.imageList1.ImageSize = New System.Drawing.Size(16, 16)
            Me.imageList1.ImageStream = (CType(resources.GetObject("imageList1.ImageStream"), System.Windows.Forms.ImageListStreamer))
            Me.imageList1.TransparentColor = System.Drawing.Color.Transparent
            ' 
            ' panel1
            ' 
            Me.panel1.Controls.Add(Me.Tree)
            Me.panel1.Dock = System.Windows.Forms.DockStyle.Left
            Me.panel1.Location = New System.Drawing.Point(0, 28)
            Me.panel1.Name = "panel1"
            Me.panel1.Size = New System.Drawing.Size(324, 637)
            Me.panel1.TabIndex = 2
            ' 
            ' Tree
            ' 
            Me.Tree.Dock = System.Windows.Forms.DockStyle.Fill
            Me.Tree.HideSelection = False
            Me.Tree.ImageIndex = -1
            Me.Tree.Location = New System.Drawing.Point(0, 0)
            Me.Tree.Name = "Tree"
            Me.Tree.SelectedImageIndex = -1
            Me.Tree.Size = New System.Drawing.Size(324, 637)
            Me.Tree.TabIndex = 0
            '			Me.Tree.KeyDown += New System.Windows.Forms.KeyEventHandler(Me.Tree_KeyDown);
            '			Me.Tree.MouseDown += New System.Windows.Forms.MouseEventHandler(Me.Tree_MouseDown);
            '			Me.Tree.AfterSelect += New System.Windows.Forms.TreeViewEventHandler(Me.Tree_AfterSelect);
            '			Me.Tree.BeforeExpand += New System.Windows.Forms.TreeViewCancelEventHandler(Me.Tree_BeforeExpand);
            ' 
            ' splitter1
            ' 
            Me.splitter1.Location = New System.Drawing.Point(324, 28)
            Me.splitter1.Name = "splitter1"
            Me.splitter1.Size = New System.Drawing.Size(4, 637)
            Me.splitter1.TabIndex = 3
            Me.splitter1.TabStop = False
            ' 
            ' panel2
            ' 
            Me.panel2.Controls.Add(Me.Text1)
            Me.panel2.Dock = System.Windows.Forms.DockStyle.Fill
            Me.panel2.Location = New System.Drawing.Point(328, 28)
            Me.panel2.Name = "panel2"
            Me.panel2.Size = New System.Drawing.Size(556, 637)
            Me.panel2.TabIndex = 4
            ' 
            ' Text1
            ' 
            Me.Text1.BackColor = System.Drawing.SystemColors.Window
            Me.Text1.Dock = System.Windows.Forms.DockStyle.Fill
            Me.Text1.Font = New System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, (CByte(0)))
            Me.Text1.HideSelection = False
            Me.Text1.Location = New System.Drawing.Point(0, 0)
            Me.Text1.Multiline = True
            Me.Text1.Name = "Text1"
            Me.Text1.ReadOnly = True
            Me.Text1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            Me.Text1.Size = New System.Drawing.Size(556, 637)
            Me.Text1.TabIndex = 1
            Me.Text1.Text = ""
            ' 
            ' mainMenu1
            ' 
            Me.mainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuFile, Me.menuEdit, Me.menuView, Me.menuHelp})
            ' 
            ' menuFile
            ' 
            Me.menuFile.Index = 0
            Me.menuFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuOpen, Me.menuSaveAs, Me.menuBar1, Me.menuRender, Me.menuPreview, Me.menuItem1, Me.menuExit})
            Me.menuFile.Text = "&File"
            ' 
            ' menuOpen
            ' 
            Me.menuOpen.Index = 0
            Me.menuOpen.Text = "&Open"
            '			Me.menuOpen.Click += New System.EventHandler(Me.menuOpen_Click);
            ' 
            ' menuSaveAs
            ' 
            Me.menuSaveAs.Enabled = False
            Me.menuSaveAs.Index = 1
            Me.menuSaveAs.Text = "Save &As..."
            '			Me.menuSaveAs.Click += New System.EventHandler(Me.menuSaveAs_Click);
            ' 
            ' menuBar1
            ' 
            Me.menuBar1.Index = 2
            Me.menuBar1.Text = "-"
            ' 
            ' menuRender
            ' 
            Me.menuRender.Enabled = False
            Me.menuRender.Index = 3
            Me.menuRender.Text = "&Render..."
            '			Me.menuRender.Click += New System.EventHandler(Me.menuRender_Click);
            ' 
            ' menuPreview
            ' 
            Me.menuPreview.Enabled = False
            Me.menuPreview.Index = 4
            Me.menuPreview.Text = "&Print Preview..."
            '			Me.menuPreview.Click += New System.EventHandler(Me.menuPreview_Click);
            ' 
            ' menuItem1
            ' 
            Me.menuItem1.Index = 5
            Me.menuItem1.Text = "-"
            ' 
            ' menuExit
            ' 
            Me.menuExit.Index = 6
            Me.menuExit.Text = "E&xit"
            '			Me.menuExit.Click += New System.EventHandler(Me.menuExit_Click);
            ' 
            ' menuEdit
            ' 
            Me.menuEdit.Index = 1
            Me.menuEdit.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuRemoveNode})
            Me.menuEdit.Text = "Edit"
            ' 
            ' menuRemoveNode
            ' 
            Me.menuRemoveNode.Enabled = False
            Me.menuRemoveNode.Index = 0
            Me.menuRemoveNode.Text = "Remove Node"
            '			Me.menuRemoveNode.Click += New System.EventHandler(Me.menuRemoveNode_Click);
            ' 
            ' menuView
            ' 
            Me.menuView.Index = 2
            Me.menuView.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuExpandAll, Me.menuCollapseAll})
            Me.menuView.Text = "&View"
            ' 
            ' menuExpandAll
            ' 
            Me.menuExpandAll.Enabled = False
            Me.menuExpandAll.Index = 0
            Me.menuExpandAll.Text = "&Expand All"
            '			Me.menuExpandAll.Click += New System.EventHandler(Me.menuExpandAll_Click);
            ' 
            ' menuCollapseAll
            ' 
            Me.menuCollapseAll.Enabled = False
            Me.menuCollapseAll.Index = 1
            Me.menuCollapseAll.Text = "&Collapse All"
            '			Me.menuCollapseAll.Click += New System.EventHandler(Me.menuCollapseAll_Click);
            ' 
            ' menuHelp
            ' 
            Me.menuHelp.Index = 3
            Me.menuHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuAbout})
            Me.menuHelp.Text = "&Help"
            ' 
            ' menuAbout
            ' 
            Me.menuAbout.Index = 0
            Me.menuAbout.Text = "&About"
            '			Me.menuAbout.Click += New System.EventHandler(Me.menuAbout_Click);
            ' 
            ' MainForm
            ' 
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.ClientSize = New System.Drawing.Size(884, 689)
            Me.Controls.Add(Me.panel2)
            Me.Controls.Add(Me.splitter1)
            Me.Controls.Add(Me.panel1)
            Me.Controls.Add(Me.toolBar1)
            Me.Controls.Add(Me.StatusBar)
            Me.Icon = (CType(resources.GetObject("$this.Icon"), System.Drawing.Icon))
            Me.Menu = Me.mainMenu1
            Me.Name = "MainForm"
            Me.Text = "Document Explorer"
            Me.panel1.ResumeLayout(False)
            Me.panel2.ResumeLayout(False)
            Me.ResumeLayout(False)

        End Sub

        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If components IsNot Nothing Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

#End Region

        <STAThread> _
        Public Shared Sub Run()
            Application.EnableVisualStyles()
            Application.DoEvents()
            Application.Run(New MainForm())
        End Sub

        ''' <summary>
        ''' Ctor.
        ''' </summary>
        Public Sub New()
            InitializeComponent()

            FindAndApplyLicense()

            Tree.ImageList = Item.ImageList
        End Sub

        ''' <summary>
        ''' Search for Aspose.Words license in the application directory.
        ''' The File.Exists check is only needed in this demo so it will work
        ''' both when the license file is missing and when it is present.
        ''' In your real application you just need to call SetLicense.
        ''' </summary>
        Private Shared Sub FindAndApplyLicense()
            ' Try to find Aspose.Total license.
            Dim licenseFile As String = Path.Combine(Application.StartupPath, "Aspose.Total.lic")
            If File.Exists(licenseFile) Then
                LicenseAsposeWords(licenseFile)
                Return
            End If

            ' Try to find Aspose.Custom license.
            licenseFile = Path.Combine(Application.StartupPath, "Aspose.Custom.lic")
            If File.Exists(licenseFile) Then
                LicenseAsposeWords(licenseFile)
                Return
            End If

            licenseFile = Path.Combine(Application.StartupPath, "Aspose.Words.lic")
            If File.Exists(licenseFile) Then
                LicenseAsposeWords(licenseFile)
            End If

        End Sub

        ''' <summary>
        ''' This code activates Aspose.Words license.
        ''' If you don't specify a license, Aspose.Words will work in evaluation mode.
        ''' </summary>
        Private Shared Sub LicenseAsposeWords(ByVal licenseFile As String)
            Dim licenseWords As New Aspose.Words.License()
            licenseWords.SetLicense(licenseFile)
        End Sub

        Private Sub menuOpen_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuOpen.Click
            OpenDocument()
        End Sub

        Private Sub menuSaveAs_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuSaveAs.Click
            SaveDocument()
        End Sub

        Private Sub menuRender_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuRender.Click
            RenderDocument()
        End Sub

        Private Sub menuPreview_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles menuPreview.Click
            PrintPreview()
        End Sub

        Private Sub menuExpandAll_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuExpandAll.Click
            ExpandAll()
        End Sub

        Private Sub menuCollapseAll_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuCollapseAll.Click
            CollapseAll()
        End Sub

        Private Sub menuRemoveNode_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuRemoveNode.Click
            Remove()
        End Sub

        Private Sub menuAbout_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuAbout.Click
            Dim aboutForm As Form = New AboutForm()
            Try
                aboutForm.ShowDialog()
            Finally
                aboutForm.Dispose()
            End Try
        End Sub

        Private Sub menuExit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles menuExit.Click
            Close()
        End Sub

        Private Sub toolBar1_ButtonClick(ByVal sender As Object, ByVal e As ToolBarButtonClickEventArgs) Handles toolBar1.ButtonClick
            Select Case e.Button.ToolTipText
                Case "Open Document"
                    OpenDocument()
                Case "Save Document As..."
                    SaveDocument()
                Case "Render Document"
                    RenderDocument()
                Case "Print Preview"
                    PrintPreview()
                Case "Expand All"
                    ExpandAll()
                Case "Collapse All"
                    CollapseAll()
                Case "Remove Node"
                    Remove()
            End Select
        End Sub

        ''' <summary>
        ''' Shows the open file dialog box and opens a document.
        ''' </summary>
        Private Sub OpenDocument()
            Dim fileName As String = SelectOpenFileName()
            If fileName Is Nothing Then
                Return
            End If

            ' This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents()
            Dim cursor As Cursor = cursor.Current
            Cursor.Current = Cursors.WaitCursor


            ' Load document is put in a try-catch block to handle situations when it fails for some reason.
            Try
                ' Loads the document into Aspose.Words object model.
                mDocument = New Document(fileName)

                Tree.BeginUpdate()

                ' Clears the tree from the previously loaded documents.
                Tree.Nodes.Clear()
                ' Creates instance of an Item class which will control the GUI representation of the document root node.
                Tree.Nodes.Add(Item.CreateItem(mDocument).TreeNode)
                ' Shows the immediate children of the document node.
                Tree.Nodes(0).Expand()
                ' Selects root node of the tree.
                Tree.SelectedNode = Tree.Nodes(0)

                Tree.EndUpdate()

                Text = "Document Explorer - " + fileName

                menuSaveAs.Enabled = True
                toolBar1.Buttons(1).Enabled = True

                menuRender.Enabled = True
                toolBar1.Buttons(3).Enabled = True
                menuPreview.Enabled = True
                toolBar1.Buttons(4).Enabled = True

                menuExpandAll.Enabled = True
                toolBar1.Buttons(6).Enabled = True
                menuCollapseAll.Enabled = True
                toolBar1.Buttons(7).Enabled = True

            Catch ex As Exception
                CType(New ExceptionDialog(ex), ExceptionDialog).ShowDialog()
            End Try

            ' Restore cursor.
            Me.Cursor = cursor
        End Sub

        ''' <summary>
        ''' Saves the document with the name and format provided in standard Save As dialog.
        ''' </summary>
        Private Sub SaveDocument()
            If mDocument Is Nothing Then
                Return
            End If

            Dim fileName As String = SelectSaveFileName()
            If fileName Is Nothing Then
                Return
            End If

            ' This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents()
            Dim cursor As Cursor = Me.Cursor
            Me.Cursor = Cursors.WaitCursor

            ' This operation is put in try-catch block to handle situations when operation fails for some reason.
            Try
                Dim saveOptions As SaveOptions = SaveOptions.CreateSaveOptions(fileName)
                saveOptions.PrettyFormat = True

                ' For most of the save formats it is enough to just invoke Aspose.Words save.
                mDocument.Save(fileName, saveOptions)
            Catch ex As Exception
                CType(New ExceptionDialog(ex), ExceptionDialog).ShowDialog()
            End Try

            ' Restore cursor.
            Me.Cursor = cursor
        End Sub

        Private Sub RenderDocument()
            If mDocument Is Nothing Then
                Return
            End If

            Dim form As New ViewerForm()
            form.Document = mDocument
            form.ShowDialog()
        End Sub

        Private Sub PrintPreview()
            If mDocument Is Nothing Then
                Return
            End If

            Preview.Execute(mDocument)
        End Sub

        ''' <summary>
        ''' Selects file name for a document to open or null.
        ''' </summary>
        Private Function SelectOpenFileName() As String
            Dim dlg As New OpenFileDialog()
            Try
                dlg.CheckFileExists = True
                dlg.CheckPathExists = True
                dlg.Title = "Open Document"
                dlg.InitialDirectory = mInitialDirectory
                dlg.Filter = "All Supported Formats (*.doc;*.dot;*.docx;*.dotx;*.docm;*.dotm;*.xml;*.wml;*.rtf;*.odt;*.ott;*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)|*.doc;*.dot;*.docx;*.dotx;*.docm;*.dotm;*.xml;*.wml;*.rtf;*.odt;*.ott;*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml|" & "Word 97-2003 Documents (*.doc;*.dot)|*.doc;*.dot|" & "Word 2007 OOXML Documents (*.docx;*.dotx;*.docm;*.dotm)|*.docx;*.dotx;*.docm;*.dotm|" & "XML Documents (*.xml;*.wml)|*.xml;*.wml|" & "Rich Text Format (*.rtf)|*.rtf|" & "OpenDocument Text (*.odt;*.ott)|*.odt;*.ott|" & "Web Pages (" & "*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)|" & "*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml|" & "All Files (*.*)|*.*"

                Dim dlgResult As DialogResult = dlg.ShowDialog()
                ' Optimized to allow automatic conversion to VB.NET
                If dlgResult.Equals(System.Windows.Forms.DialogResult.OK) Then
                    mInitialDirectory = Path.GetDirectoryName(dlg.FileName)
                    Return dlg.FileName
                Else
                    Return Nothing
                End If
            Finally
                dlg.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Selects file name for saving currently opened document or null.
        ''' </summary>
        Private Function SelectSaveFileName() As String
            Dim dlg As New SaveFileDialog()
            Try
                dlg.CheckFileExists = False
                dlg.CheckPathExists = True
                dlg.Title = "Save Document As"
                dlg.InitialDirectory = mInitialDirectory
                dlg.Filter = "Word 97-2003 Document (*.doc)|*.doc|" & "Word 2007 OOXML Document (*.docx)|*.docx|" & "Word 2007 OOXML Macro-Enabled Document (*.docm)|*.docm|" & "PDF (*.pdf)|*.pdf|" & "XPS Document (*.xps)|*.xps|" & "OpenDocument Text (*.odt)|*.odt|" & "Web Page (*.html)|*.html|" & "Single File Web Page (*.mht)|*.mht|" & "Rich Text Format (*.rtf)|*.rtf|" & "Word 2003 WordprocessingML (*.xml)|*.xml|" & "FlatOPC XML Document (*.fopc)|*.fopc|" & "Plain Text (*.txt)|*.txt|" & "IDPF EPUB Document (*.epub)|*.epub|" & "Macromedia Flash File (*.swf)|*.swf|" & "XAML Fixed Document (*.xaml)|*.xaml"

                dlg.FileName = Path.GetFileNameWithoutExtension(mDocument.OriginalFileName)

                Dim dlgResult As DialogResult = dlg.ShowDialog()
                ' Optimized to allow automatic conversion to VB.NET
                If dlgResult.Equals(System.Windows.Forms.DialogResult.OK) Then
                    mInitialDirectory = Path.GetDirectoryName(dlg.FileName)
                    Return dlg.FileName
                Else
                    Return Nothing
                End If
            Finally
                dlg.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' Expands all nodes under currently selected node.
        ''' </summary>
        Private Sub ExpandAll()
            ' This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents()
            Dim cursor As Cursor = Me.Cursor
            Me.Cursor = Cursors.WaitCursor

            If Tree.SelectedNode IsNot Nothing Then
                Tree.BeginUpdate()
                Tree.SelectedNode.ExpandAll()
                Tree.SelectedNode.EnsureVisible()
                Tree.EndUpdate()
            End If

            ' Restore cursor.
            Me.Cursor = cursor
        End Sub

        ''' <summary>
        ''' Collapses all nodes under currently selected node.
        ''' </summary>
        Private Sub CollapseAll()
            ' This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents()
            Dim cursor As Cursor = Me.Cursor
            Me.Cursor = Cursors.WaitCursor

            If Tree.SelectedNode IsNot Nothing Then
                Tree.BeginUpdate()
                Tree.SelectedNode.Collapse()
                Tree.SelectedNode.EnsureVisible()
                Tree.EndUpdate()
            End If

            ' Restore cursor.
            Me.Cursor = cursor
        End Sub

        ''' <summary>
        ''' Removes currently selected node.
        ''' </summary>
        Private Sub Remove()
            If Tree.SelectedNode IsNot Nothing Then
                CType(Tree.SelectedNode.Tag, Item).Remove()
            End If
        End Sub

        ''' <summary>
        ''' Informs Item class, which provides GUI representation of a document node,
        ''' that the corresponding TreeNode was selected.
        ''' </summary>
        Private Sub Tree_AfterSelect(ByVal sender As Object, ByVal e As TreeViewEventArgs) Handles Tree.AfterSelect
            ' This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents()
            Dim cursor As Cursor = Me.Cursor
            Me.Cursor = Cursors.WaitCursor

            Dim selectedItem As Item = CType(e.Node.Tag, Item)

            ' Set 'Remove Node' menu and tool button visibility.
            menuRemoveNode.Enabled = selectedItem.IsRemovable
            toolBar1.Buttons(9).Enabled = selectedItem.IsRemovable

            ' Show the text contained by selected document node.
            Text1.Text = selectedItem.Text

            ' Restore cursor.
            Me.Cursor = cursor
        End Sub

        ''' <summary>
        ''' Informs Item class, which provides GUI representation of a document node,
        ''' that the corresponding TreeNode is about being expanded.
        ''' </summary>
        Private Sub Tree_BeforeExpand(ByVal sender As Object, ByVal e As TreeViewCancelEventArgs) Handles Tree.BeforeExpand
            CType(e.Node.Tag, Item).OnExpand()
        End Sub

        ''' <summary>
        ''' Ensures that tree nodes are selected by right-click also.
        ''' </summary>
        Private Sub Tree_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles Tree.MouseDown
            Dim treeNode As TreeNode = Tree.GetNodeAt(e.X, e.Y)
            ' Optimized to allow automatic conversion to VB.NET
            If e.Button.Equals(MouseButtons.Right) AndAlso treeNode IsNot Nothing Then
                Tree.SelectedNode = treeNode
            End If
        End Sub

        Private Sub Tree_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles Tree.KeyDown
            ' Optimized to allow automatic conversion to VB.NET
            If e.KeyCode.Equals(Keys.Delete) Then
                Remove()
            End If
        End Sub

        ''' <summary>
        ''' The currently opened document.
        ''' </summary>
        Private mDocument As Document
        ''' <summary>
        ''' Last selected directory in the open and save dialogs.
        ''' </summary>
        Private mInitialDirectory As String = RunExamples.GetDataDir_ViewersAndVisualizers()
    End Class
End Namespace