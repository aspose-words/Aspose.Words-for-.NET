// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AsposeVisualStudioPluginWords.Core;
using System.IO;
using System.Threading;
using System.Xml;
using AsposeVisualStudioPluginWords.XML;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Xml.Linq;
using EnvDTE80;

namespace AsposeVisualStudioPluginWords.GUI
{
    public partial class SampleWizardPage : Form
    {
        private bool examplesNotAvailable = false;
        private bool downloadTaskCompleted = false;
        private DTE2 _application = null;
        CancellationTokenSource cancelToken = new CancellationTokenSource();
        TreeNode selectedNode = null;

        Task progressTask;

        public SampleWizardPage()
        {
            InitializeComponent();
            //AsposeComponents components = new AsposeComponents();
            SetComponentsAPIs();

            //AsposeComponent component;
            //AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            toolStripStatusMessage.Visible = true;
            toolStripStatusMessage.Text = "";
            progressBar1.Visible = false;
            progressBar1.Value = 0;

            //rdbCSharp.Enabled = false;
            //rdbVisualBasic.Enabled = false;

            //checkAndUpdateRepo(component);

            //rdbCSharp.Enabled = true;
            //rdbVisualBasic.Enabled = true;
        }

        public SampleWizardPage(DTE2 application)
        {
            _application = application;
            InitializeComponent();
            //AsposeComponents components = new AsposeComponents();
            SetComponentsAPIs();

            textBoxLocation.Text = GetExamplesRootPath();

            //AsposeComponent component;
            //AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);

            toolStripStatusMessage.Visible = true;
            toolStripStatusMessage.Text = "";
            progressBar1.Visible = false;
            progressBar1.Value = 0;

            //rdbCSharp.Enabled = false;
            //rdbVisualBasic.Enabled = false;

            //checkAndUpdateRepo(component);

            //rdbCSharp.Enabled = true;
            //rdbVisualBasic.Enabled = true;

            AsposeComponent component;
            AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            string repoPath = GitHelper.getLocalRepositoryPath(component);
            PopulateTreeView(repoPath + "/Examples/" + (rdbCSharp.Checked ? "CSharp" : "VisualBasic"));
        }

        private void SetComponentsAPIs()
        {
            if (string.IsNullOrEmpty(GlobalData.SelectedComponent))
            {
                ComponentWizardPage componentWizardPage = new ComponentWizardPage();
                componentWizardPage.FormClosed += new FormClosedEventHandler(components_FormClosed);
                componentWizardPage.ShowDialog();

                //if (!GlobalData.isComponentFormAborted)
                //{
                //    SetComponentsAPIs();
                //}
            }
        }
        private string GetExamplesRootPath()
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\Aspose";
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            return path;
        }

        private void components_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (GlobalData.isComponentFormAborted)
                this.Close();
        }

        public DialogResult showMessage(string title, string message, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(message, title, buttons, icon);
        }

        private bool validateForm()
        {
            clearError();
            return true;
        }

        private void setErrorMessage(string message)
        {
            toolStripStatusMessage.Text = message;
            ContinueButton.Enabled = false;
        }

        private void clearError()
        {

            toolStripStatusMessage.Text = "";
            ContinueButton.Enabled = true;
        }

        private void CloneOrCheckOutRepo(AsposeComponent component)
        {
            downloadTaskCompleted = false;
            timer1.Start();
            Task repoUpdateWorker = new Task(delegate { CloneOrCheckOutRepoWorker(component); });
            repoUpdateWorker.Start();
            progressTask = new Task(delegate { progressDisplayTask(); });
            progressBar1.Enabled = true;
            progressTask.Start();
            ContinueButton.Enabled = false;
            toolStripStatusMessage.Text = "Please wait while the Examples are being downloaded...";
        }

        private void RepositoryUpdateCompleted()
        {
            ContinueButton.Enabled = true;
            toolStripStatusMessage.Text = "Examples downloaded successfully.";
            downloadTaskCompleted = true;
            progressBar1.Value = 0;
            progressBar1.Visible = false;
        }

        private void progressDisplayTask()
        {
            try
            {
                this.Invoke(new MethodInvoker(delegate() { progressBar1.Visible = true; toolStripStatusMessage.Visible = true; toolStripStatusMessage.Text = "Fetching Examples: - Please wait while we configure you preferences"; }));
            }
            catch (Exception)
            {
            }
        }

        private void CloneOrCheckOutRepoWorker(AsposeComponent component)
        {
            GitHelper.CloneOrCheckOutRepo(component);
            downloadTaskCompleted = true;

        }

        private void checkAndUpdateRepo(AsposeComponent component)
        {
            if (null == component)
                return;
            if (null == component.get_remoteExamplesRepository() || component.RemoteExamplesRepository == string.Empty)
            {
                showMessage("Examples not available", component.get_name() + " - " + Constants.EXAMPLES_NOT_AVAILABLE_MESSAGE, MessageBoxButtons.OK, MessageBoxIcon.Information);
                examplesNotAvailable = true;
                return;
            }
            else
            {
                examplesNotAvailable = false;
            }

            if (AsposeComponentsManager.isIneternetConnected())
            {
                CloneOrCheckOutRepo(component);
            }
            else
            {
                showMessage(Constants.INTERNET_CONNECTION_REQUIRED_MESSAGE_TITLE, component.get_name() + " - " + Constants.EXAMPLES_INTERNET_CONNECTION_REQUIRED_MESSAGE, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (downloadTaskCompleted)
            {
                timer1.Stop();
                RepositoryUpdateCompleted();
            }
        }

        private string GetDestinationPath(string destinationRoot, string selectedProject)
        {
            if (!Directory.Exists(destinationRoot))
                Directory.CreateDirectory(destinationRoot);

            string path = destinationRoot + "\\" + Path.GetFileName(selectedProject);

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            else
            {
                int index = 1;
                while (Directory.Exists(path + index))
                {
                    index++;
                }
                path = path + index;
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
            }

            return path;
        }

        bool CopyAndCreateProject()
        {
            progressBar1.Visible = true;
            progressBar1.Value = 10;
            toolStripStatusMessage.Visible = true;
            toolStripStatusMessage.Text = "Fetching Examples: - Please wait while we configure you preferences";
            AsposeComponent component;
            AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            //TreeNodeData nodeData = (TreeNodeData)selectedNode.Tag;
            string sampleSourcePath = "Examples";
            string repoPath = GitHelper.getLocalRepositoryPath(component);
            string destinationPath = GetDestinationPath(textBoxLocation.Text + "\\" + Constants.ASPOSE_COMPONENT, sampleSourcePath);
            progressBar1.Value = 30;

            bool isSuccessfull = false;
            try
            {
                CopyFolderContents(Path.Combine(repoPath, sampleSourcePath), destinationPath);
                progressBar1.Value = 40;

                string dllsRootPath = AsposeComponentsManager.getLibaryDownloadPath();
                string[] dllsPaths = Directory.GetFiles(Path.Combine(dllsRootPath, component.Name + "/lib/net40/"), "*.dll");
                for (int i = 0; i < dllsPaths.Length; i++)
                {
                    if (!Directory.Exists(Path.Combine(destinationPath, "Bin")))
                        Directory.CreateDirectory(Path.Combine(destinationPath, "Bin"));
                    File.Copy(dllsPaths[i], Path.Combine(destinationPath, "Bin", Path.GetFileName(dllsPaths[i])), true);
                }

                progressBar1.Value = 50;

                string[] projectFiles = Directory.GetFiles(Path.Combine(destinationPath, (rdbCSharp.Checked ? "CSharp" : "VisualBasic")), (rdbCSharp.Checked ? "*.csproj" : "*.vbproj"));
                for (int i = 0; i < projectFiles.Length; i++)
                {
                    UpdatePrjReferenceHintPath(projectFiles[i], component);
                }
                progressBar1.Value = 70;

                int vsVersion = GetVSVersion();

                if (vsVersion >= 2010) vsVersion = 2010; // Since our examples mostly have 2010 solution files

                string[] solutionFiles = Directory.GetFiles(Path.Combine(destinationPath, (rdbCSharp.Checked ? "CSharp" : "VisualBasic")), (rdbCSharp.Checked ? "*.sln" : "*.sln"));
                progressBar1.Value = 80;

                try
                {
                    if (solutionFiles.Length > 0)
                    {
                        foreach (string sFile in solutionFiles)
                        {
                            if (sFile.Contains(vsVersion.ToString()))
                            {
                                _application.Solution.Open(sFile);
                                isSuccessfull = true;
                                break;
                            }
                        }

                        if (!isSuccessfull)
                        {
                            System.Diagnostics.Process.Start(solutionFiles[solutionFiles.Count() - 1]);
                            isSuccessfull = true;
                        }
                    }
                    else if (projectFiles.Length > 0)
                    {
                        System.Diagnostics.Process.Start(projectFiles[0]);
                        isSuccessfull = true;
                    }
                    progressBar1.Value = 90;

                }
                catch (Exception)
                {
                    if (solutionFiles.Length > 0)
                    {
                        System.Diagnostics.Process.Start(solutionFiles[0]);
                        isSuccessfull = true;
                    }
                    else if (projectFiles.Length > 0)
                    {
                        System.Diagnostics.Process.Start(projectFiles[0]);
                        isSuccessfull = true;
                    }
                    progressBar1.Value = 90;

                }
            }
            catch (Exception)
            { }

            if (!isSuccessfull)
            {
                MessageBox.Show("Oops! We are unable to open the example project. Please open it manually from " + destinationPath);
                return false;
            }
            progressBar1.Value = 100;

            return true;
        }

        private int GetVSVersion()
        {
            switch (_application.Version)
            {
                case "11.0":
                    return 2012;
                case "10.0":
                    return 2010;
                case "9.0":
                    return 2008;
                case "8.0":
                    return 2005;
            }
            return 2003;
        }

        private void UpdatePrjReferenceHintPath(string projectFilePath, AsposeComponent component)
        {
            if (!File.Exists(projectFilePath))
                return;

            XmlDocument xdDoc = new XmlDocument();
            xdDoc.Load(projectFilePath);

            XmlNamespaceManager xnManager =
             new XmlNamespaceManager(xdDoc.NameTable);
            xnManager.AddNamespace("tu",
             "http://schemas.microsoft.com/developer/msbuild/2003");

            XmlNode xnRoot = xdDoc.DocumentElement;
            XmlNodeList xnlPages = xnRoot.SelectNodes("//tu:ItemGroup", xnManager);
            foreach (XmlNode node in xnlPages)
            {
                XmlNodeList nodelist = node.SelectNodes("//tu:HintPath", xnManager);
                foreach (XmlNode nd in nodelist)
                {
                    string innter = nd.InnerText;
                    nd.InnerText = "..\\Bin\\" + component.Name + ".dll";
                }
            }
            xdDoc.Save(projectFilePath);
        }

        private bool CopyFolderContents(string SourcePath, string DestinationPath)
        {
            SourcePath = SourcePath.EndsWith(@"\") ? SourcePath : SourcePath + @"\";
            DestinationPath = DestinationPath.EndsWith(@"\") ? DestinationPath : DestinationPath + @"\";

            try
            {
                if (Directory.Exists(SourcePath))
                {
                    if (Directory.Exists(DestinationPath) == false)
                    {
                        Directory.CreateDirectory(DestinationPath);
                    }

                    foreach (string files in Directory.GetFiles(SourcePath))
                    {
                        FileInfo fileInfo = new FileInfo(files);
                        fileInfo.CopyTo(string.Format(@"{0}\{1}", DestinationPath, fileInfo.Name), true);
                    }

                    foreach (string drs in Directory.GetDirectories(SourcePath))
                    {
                        DirectoryInfo directoryInfo = new DirectoryInfo(drs);
                        if (CopyFolderContents(drs, DestinationPath + directoryInfo.Name) == false)
                        {
                            return false;
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        private void buttonGetComponents_Click(object sender, EventArgs e)
        {

        }

        private void examplesTree_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            /*selectedNode = examplesTree.SelectedNode;
            validateForm();*/
        }

        private void AbortButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ContinueButton_Click(object sender, EventArgs e)
        {
            progressDisplayTask();
            if (CopyAndCreateProject())
            {
                progressBar1.Visible = true;
                progressBar1.Value = 0;
                toolStripStatusMessage.Visible = true;
                toolStripStatusMessage.Text = "";
                Close();
            }
        }

        private void GetComponentsButton_Click(object sender, EventArgs e)
        {
            GlobalData.isAutoOpened = false;
            ComponentWizardPage wizardpage = new ComponentWizardPage();
            wizardpage.ShowDialog();
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
            {
                textBoxLocation.Text = folderBrowserDialog1.SelectedPath;
                validateForm();
            }
        }

        private void PopulateTreeView(string dirPath)
        {
            treeView1.Nodes.Clear();

            TreeNode rootNode;

            DirectoryInfo info = new DirectoryInfo(@dirPath);
            if (info.Exists)
            {
                rootNode = new TreeNode("Aspose.Words");
                rootNode.Tag = info;
                treeView1.Nodes.Add(rootNode);
                GetDirectories(info.GetDirectories(), rootNode);
                rootNode.ExpandAll();
            }
        }

        private void GetDirectories(DirectoryInfo[] subDirs,
   TreeNode nodeToAddTo)
        {
            TreeNode aNode;
            DirectoryInfo[] subSubDirs;
            foreach (DirectoryInfo subDir in subDirs)
            {
                if (!subDir.Name.ToLower().Equals("data") && !subDir.Name.ToLower().Equals("properties"))
                {
                    aNode = new TreeNode(AddSpacesToSentence(subDir.Name.Replace("-", "")), 0, 0);
                    aNode.Tag = subDir;
                    aNode.ImageKey = "folder";
                    aNode.ImageIndex = 0;
                    aNode.SelectedImageIndex = 0;
                    subSubDirs = subDir.GetDirectories();

                    if (subDir.GetFiles().Count() > 0)
                    {
                        GetFiles(subDir, aNode);
                    }
                    if (subSubDirs.Length != 0)
                    {
                        GetDirectories(subSubDirs, aNode);
                    }
                    aNode.ExpandAll();
                    nodeToAddTo.Nodes.Add(aNode);
                }
            }
            nodeToAddTo.ExpandAll();
        }

        private void GetFiles(DirectoryInfo subDirs,
   TreeNode nodeToAddTo)
        {
            TreeNode aNode;
            ListViewItem item = null;
            foreach (FileInfo file in subDirs.GetFiles())
            {
                if (file.Name.Contains(".cs") || file.Name.Contains(".vb"))
                {
                    aNode = new TreeNode(AddSpacesToSentence(file.Name.Replace(".cs", "").Replace(".vb", "").Replace("-", "")), 0, 0);
                    aNode.Tag = file;
                    aNode.ImageKey = "File";
                    aNode.ImageIndex = (rdbCSharp.Checked ? 1 : 2);
                    aNode.SelectedImageIndex = (rdbCSharp.Checked ? 1 : 2);
                    nodeToAddTo.Nodes.Add(aNode);
                }
            }
            nodeToAddTo.ExpandAll();
        }

        private void rdbCSharp_CheckedChanged(object sender, EventArgs e)
        {
            AsposeComponent component;
            AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            string repoPath = GitHelper.getLocalRepositoryPath(component);
            PopulateTreeView(repoPath + "/Examples/" + (rdbCSharp.Checked ? "CSharp" : "VisualBasic"));

        }

        private void rdbVisualBasic_CheckedChanged(object sender, EventArgs e)
        {
            AsposeComponent component;
            AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            string repoPath = GitHelper.getLocalRepositoryPath(component);
            PopulateTreeView(repoPath + "/Examples/" + (rdbCSharp.Checked ? "CSharp" : "VisualBasic"));

        }

        string AddSpacesToSentence(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            StringBuilder newText = new StringBuilder(text.Length * 2);
            newText.Append(text[0]);
            for (int i = 1; i < text.Length; i++)
            {
                if (char.IsUpper(text[i]) && text[i - 1] != ' ')
                    newText.Append(' ');
                newText.Append(text[i]);
            }
            return newText.ToString().Replace("L I N Q", "LINQ").Replace("X M L", "XML");
        }
    }
}
