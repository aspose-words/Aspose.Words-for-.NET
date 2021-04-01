using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Rendering;

namespace DocumentExplorer
{
	/// <summary>
	/// A simple form to show a Word document using Aspose.Words.Viewer.
	/// </summary>
	public class ViewerForm : System.Windows.Forms.Form
	{
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.MainMenu mainMenu;
		private System.Windows.Forms.MenuItem fileMenuItem;
		private System.Windows.Forms.MenuItem fileOpenMenuItem;
		private System.Windows.Forms.MenuItem filePrintMenuItem;
		private System.Windows.Forms.MenuItem separator1MenuItem;
		private System.Windows.Forms.MenuItem fileExitMenuItem;
		private System.Windows.Forms.MenuItem navigationPreviousPageMenuItem;
		private System.Windows.Forms.MenuItem navigationNextPageMenuItem;
		private System.Windows.Forms.ToolBar toolBar;
		private System.Windows.Forms.ToolBarButton fileOpenButton;
		private System.Windows.Forms.ToolBarButton filePrintButton;
		private System.Windows.Forms.ToolBarButton separator1;
		private System.Windows.Forms.ToolBarButton navigationPreviousPageButton;
		private System.Windows.Forms.ToolBarButton navigationNextPageButton;
		private System.Windows.Forms.ImageList toolBarImages;
		private System.Windows.Forms.StatusBar statusBar;
		private System.Windows.Forms.MenuItem separator2;
		private System.Windows.Forms.ToolBarButton navigationFirstPageButton;
		private System.Windows.Forms.ToolBarButton navigationLastPageButton;
		private System.Windows.Forms.MenuItem navigationFirstPageMenuItem;
		private System.Windows.Forms.MenuItem navigationLastPageMenuItem;
		private System.Windows.Forms.OpenFileDialog openFileDialog;
		private System.Windows.Forms.Panel mainPanel;
		private System.Windows.Forms.PictureBox docPagePictureBox;
		private System.Windows.Forms.MenuItem separator3;
		private System.Windows.Forms.MenuItem navigationGoToPageMenuItem;
		private System.Windows.Forms.ToolBarButton navigationGoToPageButton;
		private Document mDocument;
        private System.Windows.Forms.MenuItem viewMenuItem;
        private int mPageNumber;

		public ViewerForm()
		{
			InitializeComponent();
		}

		/// <summary>
		/// Gets or sets the Document to render.
		/// </summary>
		public Document Document
		{
			get { return mDocument; }
			set
			{
				mDocument = value;
                mPageNumber = 1;
				UpdatePage();
			}
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		private void OpenDocument()
		{
            if (openFileDialog.ShowDialog().Equals(DialogResult.OK))
            {
            	try
            	{
					Document = new Document(openFileDialog.FileName);
					Text = string.Format("Aspose.Words Rendering Demo - {0}", Path.GetFileNameWithoutExtension(openFileDialog.FileName));
            	}
				catch (Exception e)
            	{
            		MessageBox.Show(string.Format("Unable to load file {0}. {1}", openFileDialog.FileName, e.Message), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            	}
            }
		}
	    
	    private void PrintPreview()
	    {
            Preview.Execute(mDocument);
        }

		private void MoveToPreviousPage()
		{
            mPageNumber--;
			UpdatePage();
		}

		private void MoveToNextPage()
		{
            mPageNumber++;
			UpdatePage();
		}

		private void MoveToFirstPage()
		{
            mPageNumber = 1;
			UpdatePage();
		}

		private void MoveToLastPage()
		{
            mPageNumber = mDocument.PageCount;
			UpdatePage();
		}

		private void GoToPage()
		{
			GoToPageForm form = new GoToPageForm();
            form.MaxPageNumber = mDocument.PageCount;

			if (form.ShowDialog().Equals(DialogResult.OK))
            {
                mPageNumber = form.PageNumber;
				UpdatePage();
            }
		}

		private void UpdatePage()
		{
            // This operation can take some time (for the first page) so we set the Cursor to WaitCursor.
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            bool canMoveBack = (mPageNumber > 1);
            navigationFirstPageMenuItem.Enabled = canMoveBack;
            navigationFirstPageButton.Enabled = canMoveBack;
            navigationPreviousPageMenuItem.Enabled = canMoveBack;
            navigationPreviousPageButton.Enabled = canMoveBack;

            bool canMoveForward = (mPageNumber < mDocument.PageCount);
            navigationLastPageMenuItem.Enabled = canMoveForward;
            navigationLastPageButton.Enabled = canMoveForward;
            navigationNextPageMenuItem.Enabled = canMoveForward;
            navigationNextPageButton.Enabled = canMoveForward;

            int pageIndex = mPageNumber - 1;
            PageInfo pageInfo = mDocument.GetPageInfo(pageIndex);
            const int Resolution = 96;
            const float scale = 1.0f;
            Size imgSize = pageInfo.GetSizeInPixels(scale, Resolution);

            Bitmap img = new Bitmap(imgSize.Width, imgSize.Height);
            img.SetResolution(Resolution, Resolution);
            using (Graphics gfx = Graphics.FromImage(img))
            {
                gfx.Clear(Color.White);
                mDocument.RenderToScale(pageIndex, gfx, 0, 0, scale);
            }

            docPagePictureBox.Width = Math.Max(img.Width + 100, SystemInformation.WorkingArea.Width - SystemInformation.VerticalScrollBarWidth);
            docPagePictureBox.Height = img.Height + 100;
            docPagePictureBox.Image = img;

            statusBar.Text = string.Format("Page {0} of {1}", mPageNumber, mDocument.PageCount);

            // Restore cursor.
            Cursor.Current = cursor;
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// The contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(ViewerForm));
            this.mainMenu = new System.Windows.Forms.MainMenu();
            this.fileMenuItem = new System.Windows.Forms.MenuItem();
            this.fileOpenMenuItem = new System.Windows.Forms.MenuItem();
            this.filePrintMenuItem = new System.Windows.Forms.MenuItem();
            this.separator1MenuItem = new System.Windows.Forms.MenuItem();
            this.fileExitMenuItem = new System.Windows.Forms.MenuItem();
            this.viewMenuItem = new System.Windows.Forms.MenuItem();
            this.navigationPreviousPageMenuItem = new System.Windows.Forms.MenuItem();
            this.navigationNextPageMenuItem = new System.Windows.Forms.MenuItem();
            this.separator2 = new System.Windows.Forms.MenuItem();
            this.navigationFirstPageMenuItem = new System.Windows.Forms.MenuItem();
            this.navigationLastPageMenuItem = new System.Windows.Forms.MenuItem();
            this.separator3 = new System.Windows.Forms.MenuItem();
            this.navigationGoToPageMenuItem = new System.Windows.Forms.MenuItem();
            this.toolBar = new System.Windows.Forms.ToolBar();
            this.fileOpenButton = new System.Windows.Forms.ToolBarButton();
            this.filePrintButton = new System.Windows.Forms.ToolBarButton();
            this.separator1 = new System.Windows.Forms.ToolBarButton();
            this.navigationFirstPageButton = new System.Windows.Forms.ToolBarButton();
            this.navigationPreviousPageButton = new System.Windows.Forms.ToolBarButton();
            this.navigationNextPageButton = new System.Windows.Forms.ToolBarButton();
            this.navigationLastPageButton = new System.Windows.Forms.ToolBarButton();
            this.navigationGoToPageButton = new System.Windows.Forms.ToolBarButton();
            this.toolBarImages = new System.Windows.Forms.ImageList(this.components);
            this.statusBar = new System.Windows.Forms.StatusBar();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.mainPanel = new System.Windows.Forms.Panel();
            this.docPagePictureBox = new System.Windows.Forms.PictureBox();
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // MainMenu
            // 
            this.mainMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                     this.fileMenuItem,
                                                                                     this.viewMenuItem});
            // 
            // FileMenuItem
            // 
            this.fileMenuItem.Index = 0;
            this.fileMenuItem.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                         this.fileOpenMenuItem,
                                                                                         this.filePrintMenuItem,
                                                                                         this.separator1MenuItem,
                                                                                         this.fileExitMenuItem});
            this.fileMenuItem.Text = "&File";
            // 
            // FileOpenMenuItem
            // 
            this.fileOpenMenuItem.Index = 0;
            this.fileOpenMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlO;
            this.fileOpenMenuItem.Text = "&Open...";
            this.fileOpenMenuItem.Click += new System.EventHandler(this.fileOpenMenuItem_Click);
            // 
            // FilePrintMenuItem
            // 
            this.filePrintMenuItem.Index = 1;
            this.filePrintMenuItem.Shortcut = System.Windows.Forms.Shortcut.CtrlP;
            this.filePrintMenuItem.Text = "&Print Preview";
            this.filePrintMenuItem.Click += new System.EventHandler(this.filePrintMenuItem_Click);
            // 
            // Separator1MenuItem
            // 
            this.separator1MenuItem.Index = 2;
            this.separator1MenuItem.Text = "-";
            // 
            // FileExitMenuItem
            // 
            this.fileExitMenuItem.Index = 3;
            this.fileExitMenuItem.Text = "&Exit";
            this.fileExitMenuItem.Click += new System.EventHandler(this.fileExitMenuItem_Click);
            // 
            // ViewMenuItem
            // 
            this.viewMenuItem.Index = 1;
            this.viewMenuItem.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                         this.navigationPreviousPageMenuItem,
                                                                                         this.navigationNextPageMenuItem,
                                                                                         this.separator2,
                                                                                         this.navigationFirstPageMenuItem,
                                                                                         this.navigationLastPageMenuItem,
                                                                                         this.separator3,
                                                                                         this.navigationGoToPageMenuItem});
            this.viewMenuItem.Text = "&View";
            // 
            // NavigationPreviousPageMenuItem
            // 
            this.navigationPreviousPageMenuItem.Index = 0;
            this.navigationPreviousPageMenuItem.Text = "P&revious Page";
            this.navigationPreviousPageMenuItem.Click += new System.EventHandler(this.navigationPreviousPageMenuItem_Click);
            // 
            // NavigationNextPageMenuItem
            // 
            this.navigationNextPageMenuItem.Index = 1;
            this.navigationNextPageMenuItem.Text = "&Next Page";
            this.navigationNextPageMenuItem.Click += new System.EventHandler(this.navigationNextPageMenuItem_Click);
            // 
            // Separator2
            // 
            this.separator2.Index = 2;
            this.separator2.Text = "-";
            // 
            // NavigationFirstPageMenuItem
            // 
            this.navigationFirstPageMenuItem.Index = 3;
            this.navigationFirstPageMenuItem.Text = "&First Page";
            this.navigationFirstPageMenuItem.Click += new System.EventHandler(this.navigationFirstPageMenuItem_Click);
            // 
            // NavigationLastPageMenuItem
            // 
            this.navigationLastPageMenuItem.Index = 4;
            this.navigationLastPageMenuItem.Text = "&Last Page";
            this.navigationLastPageMenuItem.Click += new System.EventHandler(this.lastPageMenuItem_Click);
            // 
            // Separator3
            // 
            this.separator3.Index = 5;
            this.separator3.Text = "-";
            // 
            // NavigationGoToPageMenuItem
            // 
            this.navigationGoToPageMenuItem.Index = 6;
            this.navigationGoToPageMenuItem.Text = "&Go to Page...";
            this.navigationGoToPageMenuItem.Click += new System.EventHandler(this.navigationGoToPageMenuItem_Click);
            // 
            // ToolBar
            // 
            this.toolBar.Appearance = System.Windows.Forms.ToolBarAppearance.Flat;
            this.toolBar.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
                                                                                       this.fileOpenButton,
                                                                                       this.filePrintButton,
                                                                                       this.separator1,
                                                                                       this.navigationFirstPageButton,
                                                                                       this.navigationPreviousPageButton,
                                                                                       this.navigationNextPageButton,
                                                                                       this.navigationLastPageButton,
                                                                                       this.navigationGoToPageButton});
            this.toolBar.ButtonSize = new System.Drawing.Size(16, 16);
            this.toolBar.DropDownArrows = true;
            this.toolBar.ImageList = this.toolBarImages;
            this.toolBar.Location = new System.Drawing.Point(0, 0);
            this.toolBar.Name = "toolBar";
            this.toolBar.ShowToolTips = true;
            this.toolBar.Size = new System.Drawing.Size(712, 28);
            this.toolBar.TabIndex = 0;
            this.toolBar.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar_ButtonClick);
            // 
            // FileOpenButton
            // 
            this.fileOpenButton.ImageIndex = 0;
            this.fileOpenButton.ToolTipText = "Open a document";
            // 
            // FilePrintButton
            // 
            this.filePrintButton.ImageIndex = 8;
            this.filePrintButton.ToolTipText = "Print preview";
            // 
            // Separator1
            // 
            this.separator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // NavigationFirstPageButton
            // 
            this.navigationFirstPageButton.ImageIndex = 2;
            this.navigationFirstPageButton.ToolTipText = "Move to first page";
            // 
            // NavigationPreviousPageButton
            // 
            this.navigationPreviousPageButton.ImageIndex = 3;
            this.navigationPreviousPageButton.ToolTipText = "Move to previous page";
            // 
            // NavigationNextPageButton
            // 
            this.navigationNextPageButton.ImageIndex = 4;
            this.navigationNextPageButton.ToolTipText = "Move to next page";
            // 
            // NavigationLastPageButton
            // 
            this.navigationLastPageButton.ImageIndex = 5;
            this.navigationLastPageButton.ToolTipText = "Move to last page";
            // 
            // NavigationGoToPageButton
            // 
            this.navigationGoToPageButton.ImageIndex = 6;
            this.navigationGoToPageButton.ToolTipText = "Go to specified page";
            // 
            // ToolBarImages
            // 
            this.toolBarImages.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            this.toolBarImages.ImageSize = new System.Drawing.Size(16, 16);
            this.toolBarImages.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("toolBarImages.ImageStream")));
            this.toolBarImages.TransparentColor = System.Drawing.Color.Silver;
            // 
            // StatusBar
            // 
            this.statusBar.Location = new System.Drawing.Point(0, 459);
            this.statusBar.Name = "statusBar";
            this.statusBar.Size = new System.Drawing.Size(712, 22);
            this.statusBar.TabIndex = 3;
            // 
            // OpenFileDialog
            // 
            this.openFileDialog.Filter = "Microsoft Word Documents|*.doc|All files|*.*";
            // 
            // MainPanel
            // 
            this.mainPanel.AutoScroll = true;
            this.mainPanel.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(144)), ((System.Byte)(153)), ((System.Byte)(174)));
            this.mainPanel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.mainPanel.Controls.Add(this.docPagePictureBox);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 28);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(712, 431);
            this.mainPanel.TabIndex = 4;
            // 
            // DocPagePictureBox
            // 
            this.docPagePictureBox.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(144)), ((System.Byte)(153)), ((System.Byte)(174)));
            this.docPagePictureBox.Location = new System.Drawing.Point(0, 0);
            this.docPagePictureBox.Name = "docPagePictureBox";
            this.docPagePictureBox.Size = new System.Drawing.Size(56, 56);
            this.docPagePictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.docPagePictureBox.TabIndex = 0;
            this.docPagePictureBox.TabStop = false;
            // 
            // ViewerForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(712, 481);
            this.Controls.Add(this.mainPanel);
            this.Controls.Add(this.statusBar);
            this.Controls.Add(this.toolBar);
            this.Menu = this.mainMenu;
            this.Name = "ViewerForm";
            this.Text = "Aspose.Words Rendering Demo";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.mainPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }
		#endregion

		private void fileOpenMenuItem_Click(object sender, EventArgs e)
		{
			OpenDocument();
		}

		private void filePrintMenuItem_Click(object sender, EventArgs e)
		{
			PrintPreview();
		}

		private void fileExitMenuItem_Click(object sender, EventArgs e)
		{
            Close();
		}

		private void navigationPreviousPageMenuItem_Click(object sender, EventArgs e)
		{
            MoveToPreviousPage();
		}

		private void navigationNextPageMenuItem_Click(object sender, EventArgs e)
		{
            MoveToNextPage();
		}

		private void navigationFirstPageMenuItem_Click(object sender, EventArgs e)
		{
			MoveToFirstPage();
		}

		private void lastPageMenuItem_Click(object sender, EventArgs e)
		{
			MoveToLastPage();
		}

		private void navigationGoToPageMenuItem_Click(object sender, EventArgs e)
		{
			GoToPage();
		}

		private void toolBar_ButtonClick(object sender, ToolBarButtonClickEventArgs e)
		{
			switch (toolBar.Buttons.IndexOf(e.Button))
			{
				case 0:
					OpenDocument();
					break;
				case 1:
					PrintPreview();
					break;
				case 3:
					MoveToFirstPage();
					break;
				case 4:
					MoveToPreviousPage();
					break;
				case 5:
					MoveToNextPage();
					break;
				case 6:
					MoveToLastPage();
					break;
				case 7:
					GoToPage();
					break;
                default:
					throw new Exception("Unknown toolbar button index.");
			}
		}
	}
}