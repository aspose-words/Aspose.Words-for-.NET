using Aspose.Words;
using Aspose.Words.Saving;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Hosting;
using System.Web.Services;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DocumentComparison
{
    public partial class Default : System.Web.UI.Page
    {
        public string CurrentFolder
        {
            get
            {
                return ViewState["CurrentFolder"] as string;
            }
            set
            {
                ViewState["CurrentFolder"] = value;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            Common.SetLicense();

            if (!IsPostBack)
            {
                this.CurrentFolder = Common.DataDir;
            }
            
            // Handle file upload, ONLY in case of post back
            if (IsPostBack)
                UploadFile(sender, e);

            PopulateFoldersAndFiles();
        }

        private void PopulateFoldersAndFiles()
        {
            lblCurrentFolder.Text = Server.MapPath(this.CurrentFolder);
            // Get the list of files & folders in the CurrentFolder
            var currentDirInfo = new DirectoryInfo(GetFullyQualifiedFolderPath(this.CurrentFolder));
            var folders = currentDirInfo.GetDirectories();
            var files = currentDirInfo.GetFiles();

            var fsItems = new List<FileSystemItem>(folders.Length + files.Length);

            // Add the ".." option, if needed
            if (!TwoFoldersAreEquivalent(currentDirInfo.FullName, GetFullyQualifiedFolderPath(Common.DataDir)))
            {
                var parentFolder = new FileSystemItem(currentDirInfo.Parent);
                parentFolder.Name = "..";
                fsItems.Add(parentFolder);
            }

            foreach (var folder in folders)
                fsItems.Add(new FileSystemItem(folder));

            foreach (var file in files)
                fsItems.Add(new FileSystemItem(file));

            GridView1.DataSource = fsItems;
            GridView1.DataBind();
        }

        private bool TwoFoldersAreEquivalent(string folderPath1, string folderPath2)
        {
            // Chop off any trailing slashes...
            if (folderPath1.EndsWith("\\") || folderPath1.EndsWith("/"))
                folderPath1 = folderPath1.Substring(0, folderPath1.Length - 1);

            if (folderPath2.EndsWith("\\") || folderPath2.EndsWith("/"))
                folderPath2 = folderPath1.Substring(0, folderPath2.Length - 1);

            return string.CompareOrdinal(folderPath1, folderPath2) == 0;
        }

        private string GetFullyQualifiedFolderPath(string folderPath)
        {
            if (folderPath.StartsWith("~"))
                return Server.MapPath(folderPath);
            else
                return folderPath;
        }

        /// <summary>
        /// Upload a file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        protected void UploadFile(object sender, EventArgs e)
        {
            // If there is no file, then return
            if (FileUpload1.HasFile == false)
                return;

            // Save the file
            string fileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
            FileUpload1.PostedFile.SaveAs(Server.MapPath(this.CurrentFolder) + Path.DirectorySeparatorChar + fileName);
            //PopulateFoldersAndFiles();
        }

        protected void GridView1_ItemDataBound(object sender, ListViewItemEventArgs e)
        {
            var item = e.Item.DataItem as FileSystemItem;
            var lbFolderItem = e.Item.FindControl("lbFolderItem") as LinkButton;
            var lnkDownload = e.Item.FindControl("lnkDownload") as LinkButton;
            var lnkDelete = e.Item.FindControl("lnkDelete") as LinkButton;
            var chSelect = e.Item.FindControl("chSelect") as CheckBox;

            if (item.IsFolder)
            {
                lbFolderItem.Text = "<span class='glyphicon glyphicon-folder-open' aria-hidden='true'></span> " + item.Name;
                lnkDownload.Text = "";
                chSelect.Visible = false;
                // If folder is .., empty the delete link text
                if (item.Name == "..")
                {
                    lnkDelete.Text = "";
                }
                    
            }
            else
            {
                //lbFolderItem.Text = "";
                //if (this.CurrentFolder.StartsWith("~"))
                //    lbFolderItem.Text = string.Format(@"<a href=""{0}"" target=""_blank"">{1}</a>",
                //            Page.ResolveClientUrl(string.Concat(this.CurrentFolder, "/", item.Name).Replace("//", "/")),
                //            item.Name);
                //else
                //    lbFolderItem.Text = item.Name;
            }
        }

        protected void GridView1_ItemCommand(object sender, ListViewCommandEventArgs e)
        {
            // Separate the command arguments
            string[] commandArguments = e.CommandArgument.ToString().Split(Common.separator, StringSplitOptions.None);
            string fileName = commandArguments[0]; // File or folder name
            bool isFolder = bool.Parse(commandArguments[1]); // is Folder?

            if (e.CommandName == "OpenFolder")
            {
                // If it is a folder
                if (isFolder == true)
                {
                    if (string.CompareOrdinal(fileName, "..") == 0)
                    {
                        var currentFullPath = this.CurrentFolder;
                        if (currentFullPath.EndsWith("\\") || currentFullPath.EndsWith("/"))
                            currentFullPath = currentFullPath.Substring(0, currentFullPath.Length - 1);

                        currentFullPath = currentFullPath.Replace("/", "\\");

                        var folders = currentFullPath.Split("\\".ToCharArray());

                        this.CurrentFolder = string.Join("\\", folders, 0, folders.Length - 1);
                    }
                    else
                        this.CurrentFolder = Path.Combine(this.CurrentFolder, fileName as string);


                    PopulateFoldersAndFiles(); 
                }
            }
            else if (e.CommandName == "DownloadFile")
            {
                // Only download the file
                if (isFolder == false)
                {
                    Response.ContentType = ContentType;
                    Response.AppendHeader("Content-Disposition", "attachment; filename=" + Path.GetFileName(fileName));
                    Response.WriteFile(fileName);
                    Response.End(); 
                }
            }
            else if (e.CommandName == "DeleteFile")
            {
                // Handle delete file
                if (isFolder == false && File.Exists(fileName))
                {
                    File.Delete(fileName);
                    //PopulateFoldersAndFiles();
                }
                // Handle delete folder
                else
                {
                    if (Directory.Exists(fileName))
                        Directory.Delete(fileName, true);
                }
                // Delete works in thread, sleep for .5 sec, to give it time to delete the files
                // Otherwise, populate method fetches the files, which were just deleted.
                System.Threading.Thread.Sleep(500);
                // Refresh the files and folders list after deleting
                PopulateFoldersAndFiles();
            }
        }

        /// <summary>
        /// Call this folder from client using Ajax
        /// </summary>
        /// <param name="currentFolder"></param>
        /// <param name="folderName"></param>
        /// <returns></returns>
        [WebMethod]
        public static string CreateFolder(string currentFolder, string folderName)
        {
            try
            {
                Directory.CreateDirectory(currentFolder + Path.DirectorySeparatorChar + folderName);
                return Common.Success;
            }
            catch(Exception ex)
            {
                return ex.Message;
            }
        }

        /// <summary>
        /// Read the document from server using Aspose and return data
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        [WebMethod]
        public static ArrayList GetDocumentData(string filePath, string sessionID)
        {
            Common.SetLicense();

            ArrayList result = new ArrayList();
            try
            {
                // Create a temporary folder
                string documentFolder = CreateTempFolders(filePath, sessionID);

                // Load the document in Aspose.Words
                Document doc = new Document(filePath);
                // Convert the document to images
                ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);
                options.PageCount = 1;
                // Save each page of the document as image.
                for (int i = 0; i < doc.PageCount; i++)
                {
                    options.PageIndex = i;
                    doc.Save(string.Format(@"{0}\{1}.png", documentFolder, i), options);
                }
                result.Add(Common.Success); // 0. Result
                result.Add(doc.PageCount.ToString()); // 1. Page count
                result.Add(MapPathReverse(documentFolder)); // 2. Images Folder path
            }
            catch (Exception ex)
            {
                result.Clear();
                result.Add(Common.Error + ": " + ex.Message); // 0. Result
            }
            return result;
        }

        public static string MapPathReverse(string path)
        {
            string appPath = HttpContext.Current.Server.MapPath("~");
            var scheme = HttpContext.Current.Request.Url.Scheme;
            var host = HttpContext.Current.Request.Url.Host;
            var port = HttpContext.Current.Request.Url.Port;
            var virtualPath = path.Replace(appPath, "").Replace("\\", "/");
            string res = string.Format("{3}://{0}:{1}/{2}", host, port, virtualPath, scheme);
            return res;
        }

        /// <summary>
        /// Create a temporary folder to store the images for the document
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="sessionID"></param>
        private static string CreateTempFolders(string filePath, string sessionID)
        {
            // Create the folder unique to the user's session
            string tempFolder = HostingEnvironment.MapPath(Common.tempDir);
            string sessionFolder = tempFolder + Path.DirectorySeparatorChar + sessionID;
            if (Directory.Exists(sessionFolder) == false)
                Directory.CreateDirectory(sessionFolder);

            // In session folder, re-create the folder for the document
            string documentFolder = Path.Combine(sessionFolder, Path.GetFileName(filePath));
            if (Directory.Exists(documentFolder) == true)
                Directory.Delete(documentFolder, true);
            Directory.CreateDirectory(documentFolder);

            return documentFolder;
        }

        /// <summary>
        /// Compare two documents using Aspose.Words
        /// </summary>
        /// <param name="document1"></param>
        /// <param name="document2"></param>
        /// <returns></returns>
        [WebMethod]
        public static ArrayList CompareDocuments(string document1, string document2)
        {
            Common.SetLicense();

            ArrayList result = new ArrayList();
            try
            {
                // Create a temporary folder
                string comparisonDocument = GetCompareDocumentName(document1, document2);

                // Call the util class for comparison
                DocumentComparisonUtil docCompUtil = new DocumentComparisonUtil();
                int added = 0, deleted = 0;
                docCompUtil.Compare(document1, document2, comparisonDocument, ref added, ref deleted);

                result.Add(Common.Success); // 0. Result
                result.Add((comparisonDocument)); // 1. Path of the comparison document
                result.Add(added); // 2. Number of additions
                result.Add(deleted); // 3. Number of deletions
            }
            catch (Exception ex)
            {
                result.Clear();
                result.Add(Common.Error + ": " + ex.Message); // 0. Result
            }
            return result;
        }

        private static string GetCompareDocumentName(string document1, string document2)
        {
            return HostingEnvironment.MapPath(Common.tempDir) + Path.GetFileNameWithoutExtension(document1) + " Compared to " +
                Path.GetFileNameWithoutExtension(document2) + ".docx";
        }
    }
}