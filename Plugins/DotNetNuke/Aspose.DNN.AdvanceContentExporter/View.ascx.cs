using System;
using DotNetNuke.Security;
using DotNetNuke.Services.Exceptions;
using DotNetNuke.Entities.Modules;
using DotNetNuke.Entities.Modules.Actions;
using DotNetNuke.Services.Localization;
using System.Web.UI.HtmlControls;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using Aspose.Words;
using System.Collections;
using Aspose.Words.Saving;

namespace Aspose.Modules.AsposeDotNetNukeContentExport
{

    public partial class View : AsposeDotNetNukeContentExportModuleBase, IActionable
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                if (!Page.IsPostBack)
                {
                    SetLocalizationText();
                    LoadPanes();
                }
            }
            catch (Exception exc) //Module failed to load
            {
                Exceptions.ProcessModuleLoadException(this, exc);
            }
        }

        private void SetLocalizationText()
        {
            ExportButton.CssClass = Settings["ExportButtonCssClass"] != null ? Settings["ExportButtonCssClass"].ToString() : string.Empty;

            if (Settings["PaneSelectionDropDownCssClass"] != null)
            {
                if (!string.IsNullOrEmpty(Settings["PaneSelectionDropDownCssClass"].ToString()))
                    PanesDropDownList.CssClass = Settings["PaneSelectionDropDownCssClass"].ToString();
            }
            if (Settings["ExportTypeDropDownCssClass"] != null)
            {
                if (!string.IsNullOrEmpty(Settings["ExportTypeDropDownCssClass"].ToString()))
                    ExportTypeDropDown.CssClass = Settings["ExportTypeDropDownCssClass"].ToString();
            }
        }

        private void LoadPanes()
        {
            PanesDropDownList.Items.Add(new ListItem(LocalizeString("FullPage"), "dnn_full_page"));

            foreach (string pane in PortalSettings.ActiveTab.Panes)
            {
                Control obj = (Control)DotNetNuke.Common.Globals.FindControlRecursiveDown(Page, pane);

                PanesDropDownList.Items.Add(new ListItem(pane, obj.ClientID));
            }

            if (Settings["DefaultPane"] != null)
            {
                PanesDropDownList.SelectedValue = Settings["DefaultPane"].ToString();
            }

            Session["PanesDropDown_" + TabId.ToString()] = PanesDropDownList.Items;

            PanesDropDownList.Attributes.Remove("style");

            if (Settings["HideDefaultPane"] != null)
            {
                if (Convert.ToBoolean(Settings["HideDefaultPane"].ToString()))
                    PanesDropDownList.Attributes.Add("style", "display: none;");
            }
        }

        public ModuleActionCollection ModuleActions
        {
            get
            {
                var actions = new ModuleActionCollection
                    {
                        {
                            GetNextActionID(), Localization.GetString("EditModule", LocalResourceFile), "", "", "",
                            EditUrl(), false, SecurityAccessLevel.Edit, true, false
                        }
                    };
                return actions;
            }
        }

        private string GetOutputFileName(string extension)
        {
            string name = System.Guid.NewGuid().ToString() + extension;
            return name;
        }

        private string BaseURL
        {
            get
            {
                string url = Request.Url.Authority;

                if (Request.ServerVariables["HTTPS"] == "on")
                {
                    url = "https://" + url;
                }
                else
                {
                    url = "http://" + url;
                }

                return url;
            }
        }

        private void ExportContent(string exportType)
        {
            string pageSource = PageSourceHiddenField.Value;
            pageSource = "<html>" + pageSource.Replace("#g#", ">").Replace("#l#", "<") + "</html>";

            pageSource = pageSource.Replace("<div class=" + "\"exportButton\"" + ">", "<div class=" + "\"exportButton\"" + "style=" + "\"display: none\"" + ">");

            // To make the relative image paths work, base URL must be included in head section
            pageSource = pageSource.Replace("</head>", string.Format("<base href='{0}'></base></head>", BaseURL));

            // Check for license and apply if exists
            string licenseFile = Server.MapPath("~/App_Data/Aspose.Words.lic");
            if (File.Exists(licenseFile))
            {
                License license = new License();
                license.SetLicense(licenseFile);
            }

            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(pageSource));
            Document doc = new Document(stream);
            string fileName = GetOutputFileName("." + exportType);

            if (doc.PageCount > 1)
            {
                Directory.CreateDirectory(Server.MapPath("~/App_Data/" + "Zip"));
                if (exportType.Equals("Jpeg") || exportType.Equals("Png"))
                {
                    fileName = GetOutputFileName(exportType).Replace(exportType, "");
                    // Convert the html , get page count and save PNG's in Images folder
                    for (int i = 0; i < doc.PageCount; i++)
                    {
                        if (exportType.Equals("Jpeg"))
                        {
                            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg);
                            saveOptions.PageIndex = i;
                            doc.Save(Server.MapPath("~/App_Data/Zip/") + fileName + "/" + (i + 1).ToString() + "." + exportType, saveOptions);
                        }
                        else
                        {
                            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
                            saveOptions.PageIndex = i;
                            doc.Save(Server.MapPath("~/App_Data/Zip/") + fileName + "/" + (i + 1).ToString() + "." + exportType, saveOptions);
                        }
                    }
                    string filepath = Server.MapPath("~/App_Data/Zip/" + fileName + "/");
                    string downloadDirectory = Server.MapPath("~/App_Data/");
                    ZipFile.CreateFromDirectory(filepath, downloadDirectory + fileName + ".zip", CompressionLevel.Optimal, false);
                    Directory.Delete(Server.MapPath("~/App_Data/Zip/" + fileName), true);
                    System.Web.HttpResponse response = System.Web.HttpContext.Current.Response;
                    response.ClearContent();
                    response.Clear();
                    response.ContentType = "App_Data/" + exportType;
                    response.AddHeader("Content-Disposition", "attachment; filename=" + fileName + ".zip;");
                    response.TransmitFile("~/App_Data/" + fileName + ".zip");
                    //File.Delete(Server.MapPath("~/App_Data/" + fileName + ".zip"));
                    response.End();
                    return;
                }
            }
            doc.Save(Response, fileName, ContentDisposition.Attachment, null);
            Response.End();
        }

        private string GetPortalRootSavePath()
        {
            string rootPath = Server.MapPath(PortalSettings.HomeDirectory) + "\\" + "AsposeExport";
            if (!Directory.Exists(rootPath))
                Directory.CreateDirectory(rootPath);
            return rootPath;
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {
            ExportContent(GetSaveFormat(ExportTypeDropDown.SelectedValue));
        }

        // get file export types/extenssions 
        public static string GetSaveFormat(string format)
        {
            try
            {
                string saveOption = SaveFormat.Pdf.ToString();
                switch (format)
                {
                    case "Pdf":
                        saveOption = SaveFormat.Pdf.ToString(); break;
                    case "Doc":
                        saveOption = SaveFormat.Doc.ToString(); break;
                    case "Docx":
                        saveOption = SaveFormat.Docx.ToString(); break;
                    case "Odt":
                        saveOption = SaveFormat.Odt.ToString(); break;
                    case "Xps":
                        saveOption = SaveFormat.Xps.ToString(); break;
                    case "Tiff":
                        saveOption = SaveFormat.Tiff.ToString(); break;
                    case "Png":
                        saveOption = SaveFormat.Png.ToString(); break;
                    case "Jpeg":
                        saveOption = SaveFormat.Jpeg.ToString(); break;

                    // there are many document formats supported, check SaveFormat property for more
                }

                return saveOption;
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

    }
}