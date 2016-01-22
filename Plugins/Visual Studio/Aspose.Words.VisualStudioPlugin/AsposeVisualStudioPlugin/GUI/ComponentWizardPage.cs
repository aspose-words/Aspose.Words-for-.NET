// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AsposeVisualStudioPluginWords.Core;
using EnvDTE80;

namespace AsposeVisualStudioPluginWords.GUI
{
    public partial class ComponentWizardPage : Form
    {
        Task progressTask;
        private bool examplesNotAvailable = false;
        private bool downloadTaskCompleted = false;
        //AsyncDownloadList downloadList = new AsyncDownloadList();
        AsyncDownload asyncActiveDownload = null;
        private DTE2 _application = null;
        public ComponentWizardPage()
        {
            InitializeComponent();
            AsyncDownloadList.list.Clear();
            AsposeComponents components = new AsposeComponents();

            if (!GlobalData.isAutoOpened)
            {
                AbortButton.Text = "Close";
            }

            GlobalData.SelectedComponent = null;
            ContinueButton_Click(new object(), new EventArgs());


        }

        public ComponentWizardPage(DTE2 application)
        {
            _application = application;
            InitializeComponent();

            AsyncDownloadList.list.Clear();
            AsposeComponents components = new AsposeComponents();

            if (!GlobalData.isAutoOpened)
            {
                AbortButton.Text = "Close";
            }

            GlobalData.SelectedComponent = null;
            ContinueButton_Click(new object(), new EventArgs());


        }
        private void performPostFinish()
        {
            AbortButton.Enabled = true;

            AsposeComponent component;
            AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            checkAndUpdateRepo(component);
        }

        private bool performFinish()
        {
            ContinueButton.Enabled = false;
            processComponents();

            if (!AsposeComponentsManager.isIneternetConnected())
            {
                this.showMessage(Constants.INTERNET_CONNECTION_REQUIRED_MESSAGE_TITLE, Constants.INTERNET_CONNECTION_REQUIRED_MESSAGE, System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return false;
            }

            GlobalData.backgroundWorker = new BackgroundWorker();
            GlobalData.backgroundWorker.WorkerReportsProgress = true;
            GlobalData.backgroundWorker.WorkerSupportsCancellation = true;

            GlobalData.backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker_DoWork);
            GlobalData.backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker_ProgressChanged);
            GlobalData.backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker_RunWorkerCompleted);
            GlobalData.backgroundWorker.RunWorkerAsync();

            return true;
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                UpdateProgress(1);
                int total = 10;
                int index = 0;

                AsposeComponentsManager comManager = new AsposeComponentsManager(this);
                foreach (AsposeComponent component in AsposeComponents.list.Values)
                {
                    if (component.is_selected())
                    {
                        GlobalData.SelectedComponent = component.get_name();

                        if (AsposeComponentsManager.libraryAlreadyExists(component.get_downloadFileName()))
                        {
                            component.set_downloaded(true);
                        }
                        else
                        {
                            AsposeComponentsManager.addToDownloadList(component, component.get_downloadUrl(), component.get_downloadFileName());
                        }
                    }

                    decimal percentage = ((decimal)(index + 1) / (decimal)total) * 100;
                    UpdateProgress(Convert.ToInt32(percentage));

                    index++;
                }

                UpdateProgress(100);
                UpdateText("All operations completed");
            }
            catch (Exception) { }
        }

        private void UpdateText(string textToUpdate)
        {
            if (GlobalData.backgroundWorker != null)
            {
                toolStripStatusMessage.BeginInvoke(new TaskDescriptionCallback(this.TaskDescriptionLabel),
                     new object[] { textToUpdate });
            }
        }

        private void UpdateProgress(int progressValue)
        {
            if (GlobalData.backgroundWorker != null)
            {
                progressBar.BeginInvoke(new UpdateCurrentProgressBarCallback(this.UpdateCurrentProgressBar), new object[] { progressValue });
            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
                progressBar.Value = 1;
            else
                progressBar.Value = e.ProgressPercentage;
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            processDownloadList();
        }


        public delegate void TaskDescriptionCallback(string value);
        private void TaskDescriptionLabel(string value)
        {
            toolStripStatusMessage.Text = value;
        }

        public delegate void UpdateCurrentProgressBarCallback(int value);
        private void UpdateCurrentProgressBar(int value)
        {
            if (value < 0) value = 0;
            if (value > 100) value = 100;

            progressBar.Value = value;
        }


        private void processDownloadList()
        {
            if (AsyncDownloadList.list.Count > 0)
            {
                asyncActiveDownload = AsyncDownloadList.list[0];
                AsyncDownloadList.list.Remove(asyncActiveDownload);
                downloadFileFromWeb(asyncActiveDownload.Url, asyncActiveDownload.LocalPath);
                toolStripStatusMessage.Text = "Downloading " + asyncActiveDownload.Component.Name + " API";
            }
            else
            {
                performPostFinish();
            }
        }

        private void downloadFileFromWeb(string sourceURL, string destinationPath)
        {
            progressBar.Visible = true;

            // do nuget download
            //IServiceProvider serviceProvider=new System.IServiceProvider;
            //var componentModel = (IComponentModel)GetService(typeof(SComponentModel));
            //IVsPackageInstallerServices installerServices = componentModel.GetService<IVsPackageInstallerServices>();
            //var installedPackages = installerServices.GetInstalledPackages();
            //Console.WriteLine(installedPackages.FirstOrDefault().Title);


            WebClient webClient = new WebClient();
            webClient.DownloadFileCompleted += new AsyncCompletedEventHandler(Completed);
            webClient.DownloadProgressChanged += new DownloadProgressChangedEventHandler(ProgressChanged);
            Uri url = new Uri("http://packages.nuget.org/api/v1/package/" + asyncActiveDownload.Component.get_name());
            webClient.DownloadFileAsync(url, destinationPath);
        }

        private void ProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        private void Completed(object sender, AsyncCompletedEventArgs e)
        {
            progressBar.Value = 100;
            asyncActiveDownload.Component.Downloaded = true;
            AsposeComponentsManager.storeVersion(asyncActiveDownload.Component);
            UnZipDownloadedFile(asyncActiveDownload);
            AbortButton.Enabled = true;
            processDownloadList();
        }

        private void UnZipDownloadedFile(AsyncDownload download)
        {
            AsposeComponentsManager.unZipFile(download.LocalPath, Path.Combine(Path.GetDirectoryName(download.LocalPath), download.Component.Name));
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

        void processComponents()
        {
            AsposeComponent component = new AsposeComponent();
            AsposeComponents.list.TryGetValue(Constants.ASPOSE_COMPONENT, out component);
            component.Selected = true;
        }



        private void setErrorMessage(string message)
        {
            toolStripStatusMessage.Text = message;
            ContinueButton.Enabled = false;
        }

        private void clearError()
        {
            ContinueButton.Enabled = true;
        }

        private void setTooltip(Control control, string message)
        {
            ToolTip buttonToolTip = new ToolTip();
            buttonToolTip.ToolTipTitle = control.Text;
            buttonToolTip.UseFading = true;
            buttonToolTip.UseAnimation = true;
            buttonToolTip.IsBalloon = true;
            buttonToolTip.ToolTipIcon = ToolTipIcon.Info;

            buttonToolTip.ShowAlways = true;

            buttonToolTip.AutoPopDelay = 90000;
            buttonToolTip.InitialDelay = 100;
            buttonToolTip.ReshowDelay = 100;

            buttonToolTip.SetToolTip(control, message);

        }
        #region events
        private void linkLabelAspose_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkLabelAspose.LinkVisited = true;
            System.Diagnostics.Process.Start("http://www.aspose.com/.net/total-component.aspx");
        }


        #endregion

        private void logoButton_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.aspose.com");
        }

        private void textBoxProjectName_TextChanged(object sender, EventArgs e)
        {
            validateForm();
        }

        private void ContinueButton_Click(object sender, EventArgs e)
        {
            progressBar.Value = 1;
            progressBar.Visible = true;
            toolStripStatusMessage.Visible = true;
            toolStripStatusMessage.Text = "Fetching API info: - Please wait while we configure you preferences";

            GlobalData.isComponentFormAborted = false;
            performFinish();
        }

        private void AbortButton_Click(object sender, EventArgs e)
        {
            if (GlobalData.isAutoOpened)
                GlobalData.isComponentFormAborted = true;

            Close();
        }

        #region Download Examples from GitHub

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

        private void CloneOrCheckOutRepo(AsposeComponent component)
        {
            UpdateProgress(0);
            downloadTaskCompleted = false;
            timer1.Start();
            Task repoUpdateWorker = new Task(delegate { CloneOrCheckOutRepoWorker(component); });
            repoUpdateWorker.Start();
            progressTask = new Task(delegate { progressDisplayTask(); });
            progressBar.Enabled = true;
            progressTask.Start();
            ContinueButton.Enabled = false;
            toolStripStatusMessage.Text = "Please wait while the Examples are being downloaded...";
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            UpdateProgress((progressBar.Value < 90 ? progressBar.Value + 1 : 90));

            if (downloadTaskCompleted)
            {
                timer1.Stop();
                RepositoryUpdateCompleted();
            }
        }

        private void progressDisplayTask()
        {
            try
            {
                this.Invoke(new MethodInvoker(delegate() { progressBar.Visible = true; }));
            }
            catch (Exception)
            {
            }
        }

        private void RepositoryUpdateCompleted()
        {
            UpdateProgress(100);

            ContinueButton.Enabled = true;
            toolStripStatusMessage.Text = "Examples downloaded successfully.";
            downloadTaskCompleted = true;
            progressBar.Value = 0;
            progressBar.Visible = false;
            UpdateProgress(0);

            Close();
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

        private void CloneOrCheckOutRepoWorker(AsposeComponent component)
        {
            GitHelper.CloneOrCheckOutRepo(component);
            downloadTaskCompleted = true;
        }


        #endregion
    }
}
