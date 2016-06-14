
using System;
using System.IO;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Examples.CSharp;

namespace DocumentExplorerExample
{
    /// <summary>
    /// The main form of the DocumentExplorer demo.
    /// 
    /// DocumentExplorer allows to open documents using Aspose.Words.
    /// Once a document is opened, you can explore its object model in the tree.
    /// You can also save the document into DOC, DOCX, ODF, EPUB, PDF, SWF, RTF, WordML,
    /// HTML, MHTML and plain text formats.
    /// </summary>
    public class MainForm : Form
    {

        #region Windows Form Designer generated code

        private System.Windows.Forms.ToolBar toolBar1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Splitter splitter1;
        private System.Windows.Forms.Panel panel2;
        public System.Windows.Forms.StatusBar StatusBar;
        public System.Windows.Forms.TreeView Tree;
        private System.Windows.Forms.ToolBarButton toolOpenDocument;
        private System.Windows.Forms.ToolBarButton toolSaveDocument;
        private System.Windows.Forms.ToolBarButton toolExpandAll;
        private System.Windows.Forms.ToolBarButton toolCollapseAll;
        private System.Windows.Forms.ToolBarButton toolSeparator1;
        public System.Windows.Forms.TextBox Text1;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem menuFile;
        private System.Windows.Forms.MenuItem menuOpen;
        private System.Windows.Forms.MenuItem menuSaveAs;
        private System.Windows.Forms.MenuItem menuBar1;
        private System.Windows.Forms.MenuItem menuExit;
        private System.Windows.Forms.MenuItem menuView;
        private System.Windows.Forms.MenuItem menuExpandAll;
        private System.Windows.Forms.MenuItem menuCollapseAll;
        private System.Windows.Forms.MenuItem menuHelp;
        private System.Windows.Forms.MenuItem menuAbout;
        private System.Windows.Forms.ToolBarButton toolSeparator2;
        private System.Windows.Forms.ToolBarButton toolViewInWord;
        private System.Windows.Forms.ToolBarButton toolViewInPdf;
        private System.Windows.Forms.ToolBarButton toolRemove;
        private System.Windows.Forms.MenuItem menuRemoveNode;
        private System.Windows.Forms.MenuItem menuEdit;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuRender;
        private System.Windows.Forms.ToolBarButton toolRenderDocument;
        private System.Windows.Forms.ToolBarButton toolBarButton1;
        private System.Windows.Forms.ToolBarButton toolPreviewButton;
        private System.Windows.Forms.MenuItem menuPreview;
        private System.ComponentModel.IContainer components;

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(MainForm));
            this.StatusBar = new System.Windows.Forms.StatusBar();
            this.toolBar1 = new System.Windows.Forms.ToolBar();
            this.toolOpenDocument = new System.Windows.Forms.ToolBarButton();
            this.toolSaveDocument = new System.Windows.Forms.ToolBarButton();
            this.toolSeparator1 = new System.Windows.Forms.ToolBarButton();
            this.toolRenderDocument = new System.Windows.Forms.ToolBarButton();
            this.toolPreviewButton = new System.Windows.Forms.ToolBarButton();
            this.toolBarButton1 = new System.Windows.Forms.ToolBarButton();
            this.toolExpandAll = new System.Windows.Forms.ToolBarButton();
            this.toolCollapseAll = new System.Windows.Forms.ToolBarButton();
            this.toolSeparator2 = new System.Windows.Forms.ToolBarButton();
            this.toolRemove = new System.Windows.Forms.ToolBarButton();
            this.toolViewInWord = new System.Windows.Forms.ToolBarButton();
            this.toolViewInPdf = new System.Windows.Forms.ToolBarButton();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.Tree = new System.Windows.Forms.TreeView();
            this.splitter1 = new System.Windows.Forms.Splitter();
            this.panel2 = new System.Windows.Forms.Panel();
            this.Text1 = new System.Windows.Forms.TextBox();
            this.mainMenu1 = new System.Windows.Forms.MainMenu();
            this.menuFile = new System.Windows.Forms.MenuItem();
            this.menuOpen = new System.Windows.Forms.MenuItem();
            this.menuSaveAs = new System.Windows.Forms.MenuItem();
            this.menuBar1 = new System.Windows.Forms.MenuItem();
            this.menuRender = new System.Windows.Forms.MenuItem();
            this.menuPreview = new System.Windows.Forms.MenuItem();
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuExit = new System.Windows.Forms.MenuItem();
            this.menuEdit = new System.Windows.Forms.MenuItem();
            this.menuRemoveNode = new System.Windows.Forms.MenuItem();
            this.menuView = new System.Windows.Forms.MenuItem();
            this.menuExpandAll = new System.Windows.Forms.MenuItem();
            this.menuCollapseAll = new System.Windows.Forms.MenuItem();
            this.menuHelp = new System.Windows.Forms.MenuItem();
            this.menuAbout = new System.Windows.Forms.MenuItem();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // StatusBar
            // 
            this.StatusBar.Location = new System.Drawing.Point(0, 665);
            this.StatusBar.Name = "StatusBar";
            this.StatusBar.ShowPanels = true;
            this.StatusBar.Size = new System.Drawing.Size(884, 24);
            this.StatusBar.TabIndex = 0;
            this.StatusBar.Text = "statusBar1";
            // 
            // toolBar1
            // 
            this.toolBar1.Appearance = System.Windows.Forms.ToolBarAppearance.Flat;
            this.toolBar1.Buttons.AddRange(new System.Windows.Forms.ToolBarButton[] {
                                                                                        this.toolOpenDocument,
                                                                                        this.toolSaveDocument,
                                                                                        this.toolSeparator1,
                                                                                        this.toolRenderDocument,
                                                                                        this.toolPreviewButton,
                                                                                        this.toolBarButton1,
                                                                                        this.toolExpandAll,
                                                                                        this.toolCollapseAll,
                                                                                        this.toolSeparator2,
                                                                                        this.toolRemove,
                                                                                        this.toolViewInWord,
                                                                                        this.toolViewInPdf});
            this.toolBar1.DropDownArrows = true;
            this.toolBar1.ImageList = this.imageList1;
            this.toolBar1.Location = new System.Drawing.Point(0, 0);
            this.toolBar1.Name = "toolBar1";
            this.toolBar1.ShowToolTips = true;
            this.toolBar1.Size = new System.Drawing.Size(884, 28);
            this.toolBar1.TabIndex = 1;
            this.toolBar1.ButtonClick += new System.Windows.Forms.ToolBarButtonClickEventHandler(this.toolBar1_ButtonClick);
            // 
            // toolOpenDocument
            // 
            this.toolOpenDocument.ImageIndex = 0;
            this.toolOpenDocument.ToolTipText = "Open Document";
            // 
            // toolSaveDocument
            // 
            this.toolSaveDocument.Enabled = false;
            this.toolSaveDocument.ImageIndex = 1;
            this.toolSaveDocument.ToolTipText = "Save Document As...";
            // 
            // toolSeparator1
            // 
            this.toolSeparator1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // toolRenderDocument
            // 
            this.toolRenderDocument.Enabled = false;
            this.toolRenderDocument.ImageIndex = 7;
            this.toolRenderDocument.ToolTipText = "Render Document";
            // 
            // toolPreviewButton
            // 
            this.toolPreviewButton.Enabled = false;
            this.toolPreviewButton.ImageIndex = 8;
            this.toolPreviewButton.ToolTipText = "Print Preview";
            // 
            // toolBarButton1
            // 
            this.toolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // toolExpandAll
            // 
            this.toolExpandAll.Enabled = false;
            this.toolExpandAll.ImageIndex = 2;
            this.toolExpandAll.ToolTipText = "Expand All";
            // 
            // toolCollapseAll
            // 
            this.toolCollapseAll.Enabled = false;
            this.toolCollapseAll.ImageIndex = 3;
            this.toolCollapseAll.ToolTipText = "Collapse All";
            // 
            // toolSeparator2
            // 
            this.toolSeparator2.Style = System.Windows.Forms.ToolBarButtonStyle.Separator;
            // 
            // toolRemove
            // 
            this.toolRemove.Enabled = false;
            this.toolRemove.ImageIndex = 4;
            this.toolRemove.ToolTipText = "Remove Node";
            // 
            // toolViewInWord
            // 
            this.toolViewInWord.ImageIndex = 4;
            this.toolViewInWord.ToolTipText = "View in MS Word";
            this.toolViewInWord.Visible = false;
            // 
            // toolViewInPdf
            // 
            this.toolViewInPdf.ImageIndex = 5;
            this.toolViewInPdf.ToolTipText = "View in Acrobat";
            this.toolViewInPdf.Visible = false;
            // 
            // imageList1
            // 
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.Tree);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel1.Location = new System.Drawing.Point(0, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(324, 637);
            this.panel1.TabIndex = 2;
            // 
            // Tree
            // 
            this.Tree.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Tree.HideSelection = false;
            this.Tree.ImageIndex = -1;
            this.Tree.Location = new System.Drawing.Point(0, 0);
            this.Tree.Name = "Tree";
            this.Tree.SelectedImageIndex = -1;
            this.Tree.Size = new System.Drawing.Size(324, 637);
            this.Tree.TabIndex = 0;
            this.Tree.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Tree_KeyDown);
            this.Tree.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Tree_MouseDown);
            this.Tree.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.Tree_AfterSelect);
            this.Tree.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.Tree_BeforeExpand);
            // 
            // splitter1
            // 
            this.splitter1.Location = new System.Drawing.Point(324, 28);
            this.splitter1.Name = "splitter1";
            this.splitter1.Size = new System.Drawing.Size(4, 637);
            this.splitter1.TabIndex = 3;
            this.splitter1.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.Text1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(328, 28);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(556, 637);
            this.panel2.TabIndex = 4;
            // 
            // Text1
            // 
            this.Text1.BackColor = System.Drawing.SystemColors.Window;
            this.Text1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.Text1.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
            this.Text1.HideSelection = false;
            this.Text1.Location = new System.Drawing.Point(0, 0);
            this.Text1.Multiline = true;
            this.Text1.Name = "Text1";
            this.Text1.ReadOnly = true;
            this.Text1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.Text1.Size = new System.Drawing.Size(556, 637);
            this.Text1.TabIndex = 1;
            this.Text1.Text = "";
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                      this.menuFile,
                                                                                      this.menuEdit,
                                                                                      this.menuView,
                                                                                      this.menuHelp});
            // 
            // menuFile
            // 
            this.menuFile.Index = 0;
            this.menuFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                     this.menuOpen,
                                                                                     this.menuSaveAs,
                                                                                     this.menuBar1,
                                                                                     this.menuRender,
                                                                                     this.menuPreview,
                                                                                     this.menuItem1,
                                                                                     this.menuExit});
            this.menuFile.Text = "&File";
            // 
            // menuOpen
            // 
            this.menuOpen.Index = 0;
            this.menuOpen.Text = "&Open";
            this.menuOpen.Click += new System.EventHandler(this.menuOpen_Click);
            // 
            // menuSaveAs
            // 
            this.menuSaveAs.Enabled = false;
            this.menuSaveAs.Index = 1;
            this.menuSaveAs.Text = "Save &As...";
            this.menuSaveAs.Click += new System.EventHandler(this.menuSaveAs_Click);
            // 
            // menuBar1
            // 
            this.menuBar1.Index = 2;
            this.menuBar1.Text = "-";
            // 
            // menuRender
            // 
            this.menuRender.Enabled = false;
            this.menuRender.Index = 3;
            this.menuRender.Text = "&Render...";
            this.menuRender.Click += new System.EventHandler(this.menuRender_Click);
            // 
            // menuPreview
            // 
            this.menuPreview.Enabled = false;
            this.menuPreview.Index = 4;
            this.menuPreview.Text = "&Print Preview...";
            this.menuPreview.Click += new System.EventHandler(this.menuPreview_Click);
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 5;
            this.menuItem1.Text = "-";
            // 
            // menuExit
            // 
            this.menuExit.Index = 6;
            this.menuExit.Text = "E&xit";
            this.menuExit.Click += new System.EventHandler(this.menuExit_Click);
            // 
            // menuEdit
            // 
            this.menuEdit.Index = 1;
            this.menuEdit.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                     this.menuRemoveNode});
            this.menuEdit.Text = "Edit";
            // 
            // menuRemoveNode
            // 
            this.menuRemoveNode.Enabled = false;
            this.menuRemoveNode.Index = 0;
            this.menuRemoveNode.Text = "Remove Node";
            this.menuRemoveNode.Click += new System.EventHandler(this.menuRemoveNode_Click);
            // 
            // menuView
            // 
            this.menuView.Index = 2;
            this.menuView.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                     this.menuExpandAll,
                                                                                     this.menuCollapseAll});
            this.menuView.Text = "&View";
            // 
            // menuExpandAll
            // 
            this.menuExpandAll.Enabled = false;
            this.menuExpandAll.Index = 0;
            this.menuExpandAll.Text = "&Expand All";
            this.menuExpandAll.Click += new System.EventHandler(this.menuExpandAll_Click);
            // 
            // menuCollapseAll
            // 
            this.menuCollapseAll.Enabled = false;
            this.menuCollapseAll.Index = 1;
            this.menuCollapseAll.Text = "&Collapse All";
            this.menuCollapseAll.Click += new System.EventHandler(this.menuCollapseAll_Click);
            // 
            // menuHelp
            // 
            this.menuHelp.Index = 3;
            this.menuHelp.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
                                                                                     this.menuAbout});
            this.menuHelp.Text = "&Help";
            // 
            // menuAbout
            // 
            this.menuAbout.Index = 0;
            this.menuAbout.Text = "&About";
            this.menuAbout.Click += new System.EventHandler(this.menuAbout_Click);
            // 
            // MainForm
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(884, 689);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.splitter1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.toolBar1);
            this.Controls.Add(this.StatusBar);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.mainMenu1;
            this.Name = "MainForm";
            this.Text = "Document Explorer";
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #endregion

        [STAThread]
        public static void Run()
        {
            Application.EnableVisualStyles();
            Application.DoEvents();
            Application.Run(new MainForm());
        }

        /// <summary>
        /// Ctor.
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            FindAndApplyLicense();

            Tree.ImageList = Item.ImageList;
        }

        /// <summary>
        /// Search for Aspose.Words license in the application directory.
        /// The File.Exists check is only needed in this demo so it will work
        /// both when the license file is missing and when it is present.
        /// In your real application you just need to call SetLicense.
        /// </summary>
        private static void FindAndApplyLicense()
        {
            // Try to find Aspose.Total license.
            string licenseFile = Path.Combine(Application.StartupPath, "Aspose.Total.lic");
            if (File.Exists(licenseFile))
            {
                LicenseAsposeWords(licenseFile);
                return;
            }

            // Try to find Aspose.Custom license.
            licenseFile = Path.Combine(Application.StartupPath, "Aspose.Custom.lic");
            if (File.Exists(licenseFile))
            {
                LicenseAsposeWords(licenseFile);
                return;
            }

            licenseFile = Path.Combine(Application.StartupPath, "Aspose.Words.lic");
            if (File.Exists(licenseFile))
                LicenseAsposeWords(licenseFile);

        }

        /// <summary>
        /// This code activates Aspose.Words license.
        /// If you don't specify a license, Aspose.Words will work in evaluation mode.
        /// </summary>
        private static void LicenseAsposeWords(string licenseFile)
        {
            Aspose.Words.License licenseWords = new Aspose.Words.License();
            licenseWords.SetLicense(licenseFile);
        }

        private void menuOpen_Click(object sender, EventArgs e)
        {
            OpenDocument();
        }

        private void menuSaveAs_Click(object sender, EventArgs e)
        {
            SaveDocument();
        }
        
        private void menuRender_Click(object sender, System.EventArgs e)
        {
            RenderDocument();        
        }

        private void menuPreview_Click(object sender, System.EventArgs e)
        {
            PrintPreview();
        }

        private void menuExpandAll_Click(object sender, EventArgs e)
        {
            ExpandAll();
        }

        private void menuCollapseAll_Click(object sender, EventArgs e)
        {
            CollapseAll();
        }

        private void menuRemoveNode_Click(object sender, EventArgs e)
        {
            Remove();
        }

        private void menuAbout_Click(object sender, EventArgs e)
        {
            Form aboutForm = new AboutForm();
            try
            {
                aboutForm.ShowDialog();
            }
            finally
            {
                aboutForm.Dispose();
            }
        }

        private void menuExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void toolBar1_ButtonClick(object sender, ToolBarButtonClickEventArgs e)
        {
            switch (e.Button.ToolTipText)
            {
                case "Open Document":
                    OpenDocument();
                    break;
                case "Save Document As...":
                    SaveDocument();
                    break;
                case "Render Document":
                    RenderDocument();
                    break;
                case "Print Preview":
                    PrintPreview();
                    break;
                case "Expand All":
                    ExpandAll();
                    break;
                case "Collapse All":
                    CollapseAll();
                    break;
                case "Remove Node":
                    Remove();
                    break;
            }
        }

        /// <summary>
        /// Shows the open file dialog box and opens a document.
        /// </summary>
        private void OpenDocument()
        {
            string fileName = SelectOpenFileName();
            if (fileName == null)
                return;

            // This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents();
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            // Load document is put in a try-catch block to handle situations when it fails for some reason.
            try
            {
                // Loads the document into Aspose.Words object model.
                mDocument = new Document(fileName);

                Tree.BeginUpdate();

                // Clears the tree from the previously loaded documents.
                Tree.Nodes.Clear();
                // Creates instance of an Item class which will control the GUI representation of the document root node.
                Tree.Nodes.Add(Item.CreateItem(mDocument).TreeNode);
                // Shows the immediate children of the document node.
                Tree.Nodes[0].Expand();
                // Selects root node of the tree.
                Tree.SelectedNode = Tree.Nodes[0];

                Tree.EndUpdate();

                Text = "Document Explorer - " + fileName;

                menuSaveAs.Enabled = true;
                toolBar1.Buttons[1].Enabled = true;
                
                menuRender.Enabled = true;
                toolBar1.Buttons[3].Enabled = true;
                menuPreview.Enabled = true;
                toolBar1.Buttons[4].Enabled = true;
                
                menuExpandAll.Enabled = true;
                toolBar1.Buttons[6].Enabled = true;
                menuCollapseAll.Enabled = true;
                toolBar1.Buttons[7].Enabled = true;
            }
            catch (Exception ex)
            {
                new ExceptionDialog(ex).ShowDialog();
            }

            // Restore cursor.
            Cursor.Current = cursor;
        }

        /// <summary>
        /// Saves the document with the name and format provided in standard Save As dialog.
        /// </summary>
        private void SaveDocument()
        {
            if (mDocument == null)
                return;

            string fileName = SelectSaveFileName();
            if (fileName == null)
                return;

            // This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents();
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            // This operation is put in try-catch block to handle situations when operation fails for some reason.
            try
            {
                SaveOptions saveOptions = SaveOptions.CreateSaveOptions(fileName);
                saveOptions.PrettyFormat = true;

                // For most of the save formats it is enough to just invoke Aspose.Words save.
                mDocument.Save(fileName, saveOptions);
            }
            catch (Exception ex)
            {
                new ExceptionDialog(ex).ShowDialog();
            }

            // Restore cursor.
            Cursor.Current = cursor;
        }
        
        private void RenderDocument()
        {
            if (mDocument == null)
                return;
            
            ViewerForm form = new ViewerForm();
            form.Document = mDocument;
            form.ShowDialog();
        }
        
        private void PrintPreview()
        {
            if (mDocument == null)
                return;

            Preview.Execute(mDocument);
        }

        /// <summary>
        /// Selects file name for a document to open or null.
        /// </summary>
        private string SelectOpenFileName()
        {
            OpenFileDialog dlg = new OpenFileDialog();
            try
            {
                dlg.CheckFileExists = true;
                dlg.CheckPathExists = true;
                dlg.Title = "Open Document";
                dlg.InitialDirectory = mInitialDirectory;
                dlg.Filter =
                    "All Supported Formats (*.doc;*.dot;*.docx;*.dotx;*.docm;*.dotm;*.xml;*.wml;*.rtf;*.odt;*.ott;*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)|*.doc;*.dot;*.docx;*.dotx;*.docm;*.dotm;*.xml;*.wml;*.rtf;*.odt;*.ott;*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml|" +
            "Word 97-2003 Documents (*.doc;*.dot)|*.doc;*.dot|" +
            "Word 2007 OOXML Documents (*.docx;*.dotx;*.docm;*.dotm)|*.docx;*.dotx;*.docm;*.dotm|" +
            "XML Documents (*.xml;*.wml)|*.xml;*.wml|" +
            "Rich Text Format (*.rtf)|*.rtf|" +
            "OpenDocument Text (*.odt;*.ott)|*.odt;*.ott|" +
            "Web Pages (" +
            "*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml)|" +
            "*.htm;*.html;*.xhtml;*.mht;*.mhtm;*.mhtml|" +
                    "All Files (*.*)|*.*";
                    
                DialogResult dlgResult = dlg.ShowDialog();
                // Optimized to allow automatic conversion to VB.NET
                if (dlgResult.Equals(DialogResult.OK))
                {
                    mInitialDirectory = Path.GetDirectoryName(dlg.FileName);
                    return dlg.FileName;
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dlg.Dispose();
            }
        }

        /// <summary>
        /// Selects file name for saving currently opened document or null.
        /// </summary>
        private string SelectSaveFileName()
        {
            SaveFileDialog dlg = new SaveFileDialog();
            try
            {
                dlg.CheckFileExists = false;
                dlg.CheckPathExists = true;
                dlg.Title = "Save Document As";
                dlg.InitialDirectory = mInitialDirectory;
                dlg.Filter =
                    "Word 97-2003 Document (*.doc)|*.doc|" +
                    "Word 2007 OOXML Document (*.docx)|*.docx|" +
                    "Word 2007 OOXML Macro-Enabled Document (*.docm)|*.docm|" +
                    "PDF (*.pdf)|*.pdf|" +
                    "XPS Document (*.xps)|*.xps|" +
                    "OpenDocument Text (*.odt)|*.odt|" +
                    "Web Page (*.html)|*.html|" +
                    "Single File Web Page (*.mht)|*.mht|" +
                    "Rich Text Format (*.rtf)|*.rtf|" +
                    "Word 2003 WordprocessingML (*.xml)|*.xml|" +
                    "FlatOPC XML Document (*.fopc)|*.fopc|" +
                    "Plain Text (*.txt)|*.txt|" +
                    "IDPF EPUB Document (*.epub)|*.epub|" +
                    "Macromedia Flash File (*.swf)|*.swf|" +
                    "XAML Fixed Document (*.xaml)|*.xaml";

                dlg.FileName = Path.GetFileNameWithoutExtension(mDocument.OriginalFileName);

                DialogResult dlgResult = dlg.ShowDialog();
                // Optimized to allow automatic conversion to VB.NET
                if (dlgResult.Equals(DialogResult.OK))
                {
                    mInitialDirectory = Path.GetDirectoryName(dlg.FileName);
                    return dlg.FileName;
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                dlg.Dispose();
            }
        }

        /// <summary>
        /// Expands all nodes under currently selected node.
        /// </summary>
        private void ExpandAll()
        {
            // This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents();
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            if (Tree.SelectedNode != null)
            {
                Tree.BeginUpdate();
                Tree.SelectedNode.ExpandAll();
                Tree.SelectedNode.EnsureVisible();
                Tree.EndUpdate();
            }

            // Restore cursor.
            Cursor.Current = cursor;
        }

        /// <summary>
        /// Collapses all nodes under currently selected node.
        /// </summary>
        private void CollapseAll()
        {
            // This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents();
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            if (Tree.SelectedNode != null)
            {
                Tree.BeginUpdate();
                Tree.SelectedNode.Collapse();
                Tree.SelectedNode.EnsureVisible();
                Tree.EndUpdate();
            }

            // Restore cursor.
            Cursor.Current = cursor;
        }

        /// <summary>
        /// Removes currently selected node.
        /// </summary>
        private void Remove()
        {
            if (Tree.SelectedNode != null)
                ((Item)Tree.SelectedNode.Tag).Remove();
        }

        /// <summary>
        /// Informs Item class, which provides GUI representation of a document node,
        /// that the corresponding TreeNode was selected.
        /// </summary>
        private void Tree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // This operation can take some time so we set the Cursor to WaitCursor.
            Application.DoEvents();
            Cursor cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            Item selectedItem = (Item)e.Node.Tag;

            // Set 'Remove Node' menu and tool button visibility.
            menuRemoveNode.Enabled = selectedItem.IsRemovable;
            toolBar1.Buttons[9].Enabled = selectedItem.IsRemovable;

            // Show the text contained by selected document node.
            Text1.Text = selectedItem.Text;

            // Restore cursor.
            Cursor.Current = cursor;
        }

        /// <summary>
        /// Informs Item class, which provides GUI representation of a document node,
        /// that the corresponding TreeNode is about being expanded.
        /// </summary>
        private void Tree_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            ((Item)e.Node.Tag).OnExpand();
        }

        /// <summary>
        /// Ensures that tree nodes are selected by right-click also.
        /// </summary>
        private void Tree_MouseDown(object sender, MouseEventArgs e)
        {
            TreeNode treeNode = Tree.GetNodeAt(e.X, e.Y);
            // Optimized to allow automatic conversion to VB.NET
            if (e.Button.Equals(MouseButtons.Right) && treeNode != null)
                Tree.SelectedNode = treeNode;
        }

        private void Tree_KeyDown(object sender, KeyEventArgs e)
        {
            // Optimized to allow automatic conversion to VB.NET
            if (e.KeyCode.Equals(Keys.Delete))
                Remove();
        }

        /// <summary>
        /// The currently opened document.
        /// </summary>
        private Document mDocument;
        /// <summary>
        /// Last selected directory in the open and save dialogs.
        /// </summary>
        private string mInitialDirectory = RunExamples.GetDataDir_ViewersAndVisualizers();
    }
}