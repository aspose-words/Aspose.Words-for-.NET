using System;
using System.Threading;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.HtmlControls;
using System.Web;
using System.ComponentModel;
using System.Collections.Generic;
using System.Collections;
using System.IO;
using System.Configuration;
using Aspose.Words;
using System.Text;
using System.Drawing;

namespace Aspose.Words.RepeaterExport
{
    public enum WordOutputFormat
    {
        Doc, Dot, Docx, Docm, Dotx, Dotm, Rtf, Odt, Ott, Txt
    }

    [ToolboxBitmap(typeof(ExportRepeaterToWord), "icon.bmp")]
    public class ExportRepeaterToWord : Repeater, INamingContainer, IPostBackDataHandler
    {
        Button wordExportButton;

        /// <summary>
        /// Css Class that is applied to the outer div of the export button. To apply css on button you can use .yourClass input {  }
        /// </summary>
        public string ExportButtonCssClass { get; set; }

        /// <summary>
        /// Heading that is used only in the exported output word file.
        /// </summary>
        public string ExportFileHeading { get; set; }

        /// <summary>
        /// If you have paging enabled then the default output is current page. To export all pages set this datasource to all rows you want to export to Word document
        /// </summary>
        public object ExportDataSource
        {
            get { return (object)ViewState["Aspose_ExportDataSource"]; }
            set { ViewState["Aspose_ExportDataSource"] = value; }
        }

        /// <summary>
        /// Output format of the exported document. Supported formats are Doc, Dot, Docx, Docm, Dotx, Dotm, Rtf, Odt, Ott, Txt
        /// </summary>
        public WordOutputFormat ExportOutputFormat { get; set; }

        /// <summary>
        /// Local output path e.g. "c:\\temp" Disk path on server where a copy of the export is automatically saved. Application must have write access to this path.
        /// </summary>
        public string ExportOutputPathOnServer { get; set; }

        /// <summary>
        /// If true it changes the orientation of the output document to landscape. Default is Portrait
        /// </summary>
        public bool ExportInLandscape { get; set; }

        /// <summary>
        /// Export button text
        /// </summary>
        public string ExportButtonText { get; set; }

        /// <summary>
        /// Path to Aspose.Words license file e.g. c:\\Aspose.Words.lic
        /// </summary>
        public string LicenseFilePath { get; set; }

        protected override void CreateChildControls()
        {
            base.CreateChildControls();

            if (Page.IsPostBack)
            {
                if (wordExportButton == null)
                    CreateExportButton();
                Controls.Add(wordExportButton);
            }
        }

        public bool LoadPostData(string postDataKey, System.Collections.Specialized.NameValueCollection postCollection)
        {
            return false;
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            CreateExportButton();
        }

        private void CreateExportButton()
        {
            wordExportButton = new Button();
            wordExportButton.Text = string.IsNullOrEmpty(ExportButtonText) ? "Export to Word" : ExportButtonText;
            wordExportButton.ID = "__aspose_export_to_word_Repeater";
            wordExportButton.Click += new EventHandler(ExportButton_Click);
        }

        public void RaisePostDataChangedEvent()
        {
        }



        protected override void Render(HtmlTextWriter writer)
        {
            if (!Page.IsPostBack)
            {
                if (wordExportButton == null)
                    CreateExportButton();
                Controls.Add(wordExportButton);
            }

            writer.Write("<div class='" + ExportButtonCssClass + "'>");
            wordExportButton.RenderControl(writer);
            wordExportButton.Visible = false;
            writer.Write("</div>");
            writer.Write("<div>");
            base.Render(writer);
            writer.Write("</div>");
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);

            hw.RenderBeginTag("table");

            try
            {
                Control HeaderTemplate = this.Controls[0].Controls[0];
                HeaderTemplate.RenderControl(hw);
            }
            catch (Exception)
            { }

            foreach (RepeaterItem rItem in this.Items)
            {
                rItem.RenderControl(hw);
            }
            hw.RenderEndTag();

            string heading = string.IsNullOrEmpty(ExportFileHeading) ? string.Empty : ExportFileHeading;

            string pageSource = "<html><head></head><body>" + heading + sw.ToString() + "</body></html>";

            // Check for license and apply if exists
            if (File.Exists(LicenseFilePath))
            {
                License license = new License();
                license.SetLicense(LicenseFilePath);
            }

            MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(pageSource));
            Document doc = new Document(stream);

            string extension = ExportOutputFormat.ToString().ToLower();

            if (string.IsNullOrEmpty(extension)) extension = "doc";
            string fileName = System.Guid.NewGuid() + "." + extension;

            if (!string.IsNullOrEmpty(ExportOutputPathOnServer) && Directory.Exists(ExportOutputPathOnServer))
            {
                try
                {
                    doc.Save(ExportOutputPathOnServer + "\\" + fileName);
                }
                catch (Exception) { }
            }

            if (ExportInLandscape)
            {
                foreach (Section section in doc)
                    section.PageSetup.Orientation = Orientation.Landscape;
            }

            doc.Save(HttpContext.Current.Response, fileName, ContentDisposition.Inline, null);
            HttpContext.Current.Response.End();
        }
    }
}
