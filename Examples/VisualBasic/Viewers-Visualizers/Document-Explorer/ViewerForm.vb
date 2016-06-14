Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Aspose.Words
Imports Aspose.Words.Rendering

Namespace DocumentExplorerExample
	''' <summary>
	''' A simple form to show a Word document using Aspose.Words.Viewer.
	''' </summary>
	Public Class ViewerForm
		Inherits System.Windows.Forms.Form
		Private components As System.ComponentModel.IContainer
		Private mainMenu As System.Windows.Forms.MainMenu
		Private fileMenuItem As System.Windows.Forms.MenuItem
		Private WithEvents fileOpenMenuItem As System.Windows.Forms.MenuItem
		Private WithEvents filePrintMenuItem As System.Windows.Forms.MenuItem
		Private separator1MenuItem As System.Windows.Forms.MenuItem
		Private WithEvents fileExitMenuItem As System.Windows.Forms.MenuItem
		Private WithEvents navigationPreviousPageMenuItem As System.Windows.Forms.MenuItem
		Private WithEvents navigationNextPageMenuItem As System.Windows.Forms.MenuItem
		Private WithEvents toolBar As System.Windows.Forms.ToolBar
		Private fileOpenButton As System.Windows.Forms.ToolBarButton
		Private filePrintButton As System.Windows.Forms.ToolBarButton
		Private separator1 As System.Windows.Forms.ToolBarButton
		Private navigationPreviousPageButton As System.Windows.Forms.ToolBarButton
		Private navigationNextPageButton As System.Windows.Forms.ToolBarButton
		Private toolBarImages As System.Windows.Forms.ImageList
		Private statusBar As System.Windows.Forms.StatusBar
		Private separator2 As System.Windows.Forms.MenuItem
		Private navigationFirstPageButton As System.Windows.Forms.ToolBarButton
		Private navigationLastPageButton As System.Windows.Forms.ToolBarButton
		Private WithEvents navigationFirstPageMenuItem As System.Windows.Forms.MenuItem
		Private WithEvents navigationLastPageMenuItem As System.Windows.Forms.MenuItem
		Private openFileDialog As System.Windows.Forms.OpenFileDialog
		Private mainPanel As System.Windows.Forms.Panel
		Private docPagePictureBox As System.Windows.Forms.PictureBox
		Private separator3 As System.Windows.Forms.MenuItem
		Private WithEvents navigationGoToPageMenuItem As System.Windows.Forms.MenuItem
		Private navigationGoToPageButton As System.Windows.Forms.ToolBarButton
		Private mDocument As Document
		Private viewMenuItem As System.Windows.Forms.MenuItem
		Private mPageNumber As Integer

		Public Sub New()
			InitializeComponent()
		End Sub

		''' <summary>
		''' Gets or sets the Document to render.
		''' </summary>
		Public Property Document() As Document
			Get
				Return mDocument
			End Get
			Set(ByVal value As Document)
				mDocument = value
				mPageNumber = 1
				UpdatePage()
			End Set
		End Property

		''' <summary>
		''' Clean up any resources being used.
		''' </summary>
		Protected Overrides Overloads Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If components IsNot Nothing Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		Private Sub OpenDocument()
			If openFileDialog.ShowDialog().Equals(System.Windows.Forms.DialogResult.OK) Then
				Try
					Document = New Document(openFileDialog.FileName)
					Text = String.Format("Aspose.Words Rendering Demo - {0}", Path.GetFileNameWithoutExtension(openFileDialog.FileName))
				Catch e As Exception
					MessageBox.Show(String.Format("Unable to load file {0}. {1}", openFileDialog.FileName, e.Message), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
				End Try
			End If
		End Sub

		Private Sub PrintPreview()
			Preview.Execute(mDocument)
		End Sub

		Private Sub MoveToPreviousPage()
			mPageNumber -= 1
			UpdatePage()
		End Sub

		Private Sub MoveToNextPage()
			mPageNumber += 1
			UpdatePage()
		End Sub

		Private Sub MoveToFirstPage()
			mPageNumber = 1
			UpdatePage()
		End Sub

		Private Sub MoveToLastPage()
			mPageNumber = mDocument.PageCount
			UpdatePage()
		End Sub

		Private Sub GoToPage()
			Dim form As New GoToPageForm()
			form.MaxPageNumber = mDocument.PageCount

			If form.ShowDialog().Equals(System.Windows.Forms.DialogResult.OK) Then
				mPageNumber = form.PageNumber
				UpdatePage()
			End If
		End Sub

		Private Sub UpdatePage()
			' This operation can take some time (for the first page) so we set the Cursor to WaitCursor.
            Dim cursor As Cursor = Me.Cursor
            Me.Cursor = Cursors.WaitCursor

			Dim canMoveBack As Boolean = (mPageNumber > 1)
			navigationFirstPageMenuItem.Enabled = canMoveBack
			navigationFirstPageButton.Enabled = canMoveBack
			navigationPreviousPageMenuItem.Enabled = canMoveBack
			navigationPreviousPageButton.Enabled = canMoveBack

			Dim canMoveForward As Boolean = (mPageNumber < mDocument.PageCount)
			navigationLastPageMenuItem.Enabled = canMoveForward
			navigationLastPageButton.Enabled = canMoveForward
			navigationNextPageMenuItem.Enabled = canMoveForward
			navigationNextPageButton.Enabled = canMoveForward

			Dim pageIndex As Integer = mPageNumber - 1
			Dim pageInfo As PageInfo = mDocument.GetPageInfo(pageIndex)
			Const Resolution As Integer = 96
			Const scale As Single = 1.0f
			Dim imgSize As Size = pageInfo.GetSizeInPixels(scale, Resolution)

			Dim img As New Bitmap(imgSize.Width, imgSize.Height)
			img.SetResolution(Resolution, Resolution)
			Using gfx As Graphics = Graphics.FromImage(img)
				gfx.Clear(Color.White)
				mDocument.RenderToScale(pageIndex, gfx, 0, 0, scale)
			End Using

            docPagePictureBox.Width = System.Math.Max(img.Width + 100, SystemInformation.WorkingArea.Width - SystemInformation.VerticalScrollBarWidth)
			docPagePictureBox.Height = img.Height + 100
			docPagePictureBox.Image = img

			statusBar.Text = String.Format("Page {0} of {1}", mPageNumber, mDocument.PageCount)

			' Restore cursor.
            Me.Cursor = cursor
		End Sub

		#Region "Windows Form Designer generated code"
		''' <summary>
		''' Required method for Designer support - do not modify
		''' the contents of this method with the code editor.
		''' </summary>
		Private Sub InitializeComponent()
			Me.components = New System.ComponentModel.Container()
			Dim resources As New System.Resources.ResourceManager(GetType(ViewerForm))
			Me.mainMenu = New System.Windows.Forms.MainMenu()
			Me.fileMenuItem = New System.Windows.Forms.MenuItem()
			Me.fileOpenMenuItem = New System.Windows.Forms.MenuItem()
			Me.filePrintMenuItem = New System.Windows.Forms.MenuItem()
			Me.separator1MenuItem = New System.Windows.Forms.MenuItem()
			Me.fileExitMenuItem = New System.Windows.Forms.MenuItem()
			Me.viewMenuItem = New System.Windows.Forms.MenuItem()
			Me.navigationPreviousPageMenuItem = New System.Windows.Forms.MenuItem()
			Me.navigationNextPageMenuItem = New System.Windows.Forms.MenuItem()
			Me.separator2 = New System.Windows.Forms.MenuItem()
			Me.navigationFirstPageMenuItem = New System.Windows.Forms.MenuItem()
			Me.navigationLastPageMenuItem = New System.Windows.Forms.MenuItem()
			Me.separator3 = New System.Windows.Forms.MenuItem()
			Me.navigationGoToPageMenuItem = New System.Windows.Forms.MenuItem()
			Me.toolBar = New System.Windows.Forms.ToolBar()
			Me.fileOpenButton = New System.Windows.Forms.ToolBarButton()
			Me.filePrintButton = New System.Windows.Forms.ToolBarButton()
			Me.separator1 = New System.Windows.Forms.ToolBarButton()
			Me.navigationFirstPageButton = New System.Windows.Forms.ToolBarButton()
			Me.navigationPreviousPageButton = New System.Windows.Forms.ToolBarButton()
			Me.navigationNextPageButton = New System.Windows.Forms.ToolBarButton()
			Me.navigationLastPageButton = New System.Windows.Forms.ToolBarButton()
			Me.navigationGoToPageButton = New System.Windows.Forms.ToolBarButton()
			Me.toolBarImages = New System.Windows.Forms.ImageList(Me.components)
			Me.statusBar = New System.Windows.Forms.StatusBar()
			Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
			Me.mainPanel = New System.Windows.Forms.Panel()
			Me.docPagePictureBox = New System.Windows.Forms.PictureBox()
			Me.mainPanel.SuspendLayout()
			Me.SuspendLayout()
			' 
			' mainMenu
			' 
			Me.mainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.fileMenuItem, Me.viewMenuItem})
			' 
			' fileMenuItem
			' 
			Me.fileMenuItem.Index = 0
			Me.fileMenuItem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.fileOpenMenuItem, Me.filePrintMenuItem, Me.separator1MenuItem, Me.fileExitMenuItem})
			Me.fileMenuItem.Text = "&File"
			' 
			' fileOpenMenuItem
			' 
			Me.fileOpenMenuItem.Index = 0
			Me.fileOpenMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlO
			Me.fileOpenMenuItem.Text = "&Open..."
'			Me.fileOpenMenuItem.Click += New System.EventHandler(Me.fileOpenMenuItem_Click);
			' 
			' filePrintMenuItem
			' 
			Me.filePrintMenuItem.Index = 1
			Me.filePrintMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlP
			Me.filePrintMenuItem.Text = "&Print Preview"
'			Me.filePrintMenuItem.Click += New System.EventHandler(Me.filePrintMenuItem_Click);
			' 
			' separator1MenuItem
			' 
			Me.separator1MenuItem.Index = 2
			Me.separator1MenuItem.Text = "-"
			' 
			' fileExitMenuItem
			' 
			Me.fileExitMenuItem.Index = 3
			Me.fileExitMenuItem.Text = "&Exit"
'			Me.fileExitMenuItem.Click += New System.EventHandler(Me.fileExitMenuItem_Click);
			' 
			' viewMenuItem
			' 
			Me.viewMenuItem.Index = 1
			Me.viewMenuItem.MenuItems.AddRange(New System.Windows.Forms.MenuItem() { Me.navigationPreviousPageMenuItem, Me.navigationNextPageMenuItem, Me.separator2, Me.navigationFirstPageMenuItem, Me.navigationLastPageMenuItem, Me.separator3, Me.navigationGoToPageMenuItem})
			Me.viewMenuItem.Text = "&View"
			' 
			' navigationPreviousPageMenuItem
			' 
			Me.navigationPreviousPageMenuItem.Index = 0
			Me.navigationPreviousPageMenuItem.Text = "P&revious Page"
'			Me.navigationPreviousPageMenuItem.Click += New System.EventHandler(Me.navigationPreviousPageMenuItem_Click);
			' 
			' navigationNextPageMenuItem
			' 
			Me.navigationNextPageMenuItem.Index = 1
			Me.navigationNextPageMenuItem.Text = "&Next Page"
'			Me.navigationNextPageMenuItem.Click += New System.EventHandler(Me.navigationNextPageMenuItem_Click);
			' 
			' separator2
			' 
			Me.separator2.Index = 2
			Me.separator2.Text = "-"
			' 
			' navigationFirstPageMenuItem
			' 
			Me.navigationFirstPageMenuItem.Index = 3
			Me.navigationFirstPageMenuItem.Text = "&First Page"
'			Me.navigationFirstPageMenuItem.Click += New System.EventHandler(Me.navigationFirstPageMenuItem_Click);
			' 
			' navigationLastPageMenuItem
			' 
			Me.navigationLastPageMenuItem.Index = 4
			Me.navigationLastPageMenuItem.Text = "&Last Page"
'			Me.navigationLastPageMenuItem.Click += New System.EventHandler(Me.lastPageMenuItem_Click);
			' 
			' separator3
			' 
			Me.separator3.Index = 5
			Me.separator3.Text = "-"
			' 
			' navigationGoToPageMenuItem
			' 
			Me.navigationGoToPageMenuItem.Index = 6
			Me.navigationGoToPageMenuItem.Text = "&Go to Page..."
'			Me.navigationGoToPageMenuItem.Click += New System.EventHandler(Me.navigationGoToPageMenuItem_Click);
			' 
			' toolBar
			' 
			Me.toolBar.Appearance = System.Windows.Forms.ToolBarAppearance.Flat
			Me.toolBar.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() { Me.fileOpenButton, Me.filePrintButton, Me.separator1, Me.navigationFirstPageButton, Me.navigationPreviousPageButton, Me.navigationNextPageButton, Me.navigationLastPageButton, Me.navigationGoToPageButton})
			Me.toolBar.ButtonSize = New System.Drawing.Size(16, 16)
			Me.toolBar.DropDownArrows = True
			Me.toolBar.ImageList = Me.toolBarImages
			Me.toolBar.Location = New System.Drawing.Point(0, 0)
			Me.toolBar.Name = "toolBar"
			Me.toolBar.ShowToolTips = True
			Me.toolBar.Size = New System.Drawing.Size(712, 28)
			Me.toolBar.TabIndex = 0
'			Me.toolBar.ButtonClick += New System.Windows.Forms.ToolBarButtonClickEventHandler(Me.toolBar_ButtonClick);
			' 
			' fileOpenButton
			' 
			Me.fileOpenButton.ImageIndex = 0
			Me.fileOpenButton.ToolTipText = "Open a document"
			' 
			' filePrintButton
			' 
			Me.filePrintButton.ImageIndex = 8
			Me.filePrintButton.ToolTipText = "Print preview"
			' 
			' separator1
			' 
			Me.separator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
			' 
			' navigationFirstPageButton
			' 
			Me.navigationFirstPageButton.ImageIndex = 2
			Me.navigationFirstPageButton.ToolTipText = "Move to first page"
			' 
			' navigationPreviousPageButton
			' 
			Me.navigationPreviousPageButton.ImageIndex = 3
			Me.navigationPreviousPageButton.ToolTipText = "Move to previous page"
			' 
			' navigationNextPageButton
			' 
			Me.navigationNextPageButton.ImageIndex = 4
			Me.navigationNextPageButton.ToolTipText = "Move to next page"
			' 
			' navigationLastPageButton
			' 
			Me.navigationLastPageButton.ImageIndex = 5
			Me.navigationLastPageButton.ToolTipText = "Move to last page"
			' 
			' navigationGoToPageButton
			' 
			Me.navigationGoToPageButton.ImageIndex = 6
			Me.navigationGoToPageButton.ToolTipText = "Go to specified page"
			' 
			' toolBarImages
			' 
			Me.toolBarImages.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit
			Me.toolBarImages.ImageSize = New System.Drawing.Size(16, 16)
			Me.toolBarImages.ImageStream = (CType(resources.GetObject("toolBarImages.ImageStream"), System.Windows.Forms.ImageListStreamer))
			Me.toolBarImages.TransparentColor = System.Drawing.Color.Silver
			' 
			' statusBar
			' 
			Me.statusBar.Location = New System.Drawing.Point(0, 459)
			Me.statusBar.Name = "statusBar"
			Me.statusBar.Size = New System.Drawing.Size(712, 22)
			Me.statusBar.TabIndex = 3
			' 
			' openFileDialog
			' 
			Me.openFileDialog.Filter = "Microsoft Word Documents|*.doc|All files|*.*"
			' 
			' mainPanel
			' 
			Me.mainPanel.AutoScroll = True
			Me.mainPanel.BackColor = System.Drawing.Color.FromArgb((CByte(144)), (CByte(153)), (CByte(174)))
			Me.mainPanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
			Me.mainPanel.Controls.Add(Me.docPagePictureBox)
			Me.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill
			Me.mainPanel.Location = New System.Drawing.Point(0, 28)
			Me.mainPanel.Name = "mainPanel"
			Me.mainPanel.Size = New System.Drawing.Size(712, 431)
			Me.mainPanel.TabIndex = 4
			' 
			' docPagePictureBox
			' 
			Me.docPagePictureBox.BackColor = System.Drawing.Color.FromArgb((CByte(144)), (CByte(153)), (CByte(174)))
			Me.docPagePictureBox.Location = New System.Drawing.Point(0, 0)
			Me.docPagePictureBox.Name = "docPagePictureBox"
			Me.docPagePictureBox.Size = New System.Drawing.Size(56, 56)
			Me.docPagePictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
			Me.docPagePictureBox.TabIndex = 0
			Me.docPagePictureBox.TabStop = False
			' 
			' ViewerForm
			' 
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.ClientSize = New System.Drawing.Size(712, 481)
			Me.Controls.Add(Me.mainPanel)
			Me.Controls.Add(Me.statusBar)
			Me.Controls.Add(Me.toolBar)
			Me.Menu = Me.mainMenu
			Me.Name = "ViewerForm"
			Me.Text = "Aspose.Words Rendering Demo"
			Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
			Me.mainPanel.ResumeLayout(False)
			Me.ResumeLayout(False)

		End Sub
		#End Region

		Private Sub fileOpenMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fileOpenMenuItem.Click
			OpenDocument()
		End Sub

		Private Sub filePrintMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles filePrintMenuItem.Click
			PrintPreview()
		End Sub

		Private Sub fileExitMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles fileExitMenuItem.Click
			Close()
		End Sub

		Private Sub navigationPreviousPageMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navigationPreviousPageMenuItem.Click
			MoveToPreviousPage()
		End Sub

		Private Sub navigationNextPageMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navigationNextPageMenuItem.Click
			MoveToNextPage()
		End Sub

		Private Sub navigationFirstPageMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navigationFirstPageMenuItem.Click
			MoveToFirstPage()
		End Sub

		Private Sub lastPageMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navigationLastPageMenuItem.Click
			MoveToLastPage()
		End Sub

		Private Sub navigationGoToPageMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles navigationGoToPageMenuItem.Click
			GoToPage()
		End Sub

		Private Sub toolBar_ButtonClick(ByVal sender As Object, ByVal e As ToolBarButtonClickEventArgs) Handles toolBar.ButtonClick
			Select Case toolBar.Buttons.IndexOf(e.Button)
				Case 0
					OpenDocument()
				Case 1
					PrintPreview()
				Case 3
					MoveToFirstPage()
				Case 4
					MoveToPreviousPage()
				Case 5
					MoveToNextPage()
				Case 6
					MoveToLastPage()
				Case 7
					GoToPage()
				Case Else
					Throw New Exception("Unknown toolbar button index.")
			End Select
		End Sub
	End Class
End Namespace