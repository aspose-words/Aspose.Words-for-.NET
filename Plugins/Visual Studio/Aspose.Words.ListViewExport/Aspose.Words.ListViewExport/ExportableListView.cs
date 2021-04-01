using System;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Drawing;

namespace Aspose.Words.ListViewExport
{
    public enum WordOutputFormat
    {
        Doc, Dot, Docx, Docm, Dotx, Dotm, Rtf, Odt, Ott, Txt
    }

    [ProvideToolboxControl("Aspose", false)]
    [ToolboxBitmap(typeof(ExportListViewToWord), "icon.bmp")]
    public class ExportListViewToWord : ListView, INamingContainer
    {
        Button wordExportButton;

        /// <summary>
        /// Css Class that is applied to the outer div of the export button. To apply css on button you can use .yourClass input {  }
        /// </summary>
        public string ExportButtonCssClass { get; set; }

        /// <summary>
        /// Heading that is used only in the exported output word file.
        /// </summary>
        [DefaultValue("<h4>Example Report</h4>")]
        public string ExportFileHeading { get; set; }

        /// <summary>
        /// If you have paging enabled then the default output is current page. To export all pages set this datasource, all rows will be exported to Word document.
        /// </summary>
        public object ExportDataSource
        {
            get { return (object)ViewState["Aspose_ExportDataSource"]; }
            set { ViewState["Aspose_ExportDataSource"] = value; }
        }

        /// <summary>
        /// If DataPager is enabled then: Maximum number of records to export, integer value only
        /// </summary>
        [DefaultValue(6500)]
        public int ExportMaximumRecords { get; set; }

        /// <summary>
        /// Output format of the exported document. Supported formats are Doc, Dot, Docx, Docm, Dotx, Dotm, Rtf, Odt, Ott, Txt
        /// </summary>
        [DefaultValue("c:\\temp")]
        public WordOutputFormat ExportOutputFormat { get; set; }

        /// <summary>
        /// Local output path e.g. "c:\\temp" Disk path on server where a copy of the export is automatically saved. Application must have write access to this path.
        /// </summary>
        [DefaultValue("c:\\temp")]
        public string ExportOutputPathOnServer { get; set; }

        /// <summary>
        /// If true it changes the orientation of the output document to landscape. Default is Portrait
        /// </summary>
        [DefaultValue(true)]
        public bool ExportInLandscape { get; set; }

        /// <summary>
        /// Export button text
        /// </summary>
        [DefaultValue("Export to Word")]
        public string ExportButtonText { get; set; }

        /// <summary>
        /// Path to Aspose.Words license file e.g. c:\\Aspose.Words.lic
        /// </summary>
        [DefaultValue("e:\\Aspose\\Aspose.Words.lic")]
        public string LicenseFilePath { get; set; }

        protected override int CreateChildControls(System.Collections.IEnumerable dataSource, bool dataBinding)
        {
            var rowCount = base.CreateChildControls(dataSource, dataBinding);
            if (wordExportButton == null)
                CreateExportButton();
            Controls.Add(wordExportButton);
            return rowCount;
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
            wordExportButton.ID = "__aspose_export_to_word_listview";
            wordExportButton.Click += new EventHandler(ExportButton_Click);
        }

        private String CalculateWidth()
        {
            string strWidth = "auto";
            if (!this.Width.IsEmpty)
            {
                strWidth = String.Format("{0}{1}", this.Width.Value, ((this.Width.Type == UnitType.Percentage) ? "%" : "px"));
            }
            return strWidth;
        }

        protected override void Render(HtmlTextWriter writer)
        {
            writer.Write("<div style='width:" + CalculateWidth() + "'>");
            writer.Write("<div class='" + ExportButtonCssClass + "'>");
            wordExportButton.RenderControl(writer);
            wordExportButton.Visible = false;
            writer.Write("</div>");
            writer.Write("<div>");
            base.RenderContents(writer);
            writer.Write("</div></div>");
        }

        protected void ExportButton_Click(object sender, EventArgs e)
        {
            StringWriter sw = new StringWriter();
            HtmlTextWriter hw = new HtmlTextWriter(sw);

            DataPager objDataPager = FindControl<DataPager>(this.Controls);
            if (ExportDataSource != null)
            {
                if (objDataPager != null)
                {
                    if (ExportMaximumRecords == 0)
                        ExportMaximumRecords = 65000;
                    objDataPager.Visible = false;
                    objDataPager.PageSize = ExportMaximumRecords;
                    objDataPager.SetPageProperties(0, ExportMaximumRecords, false);
                    this.DataBind();
                }
            }
            else
            {
                if (objDataPager != null)
                {
                    objDataPager.Visible = false;
                }
            }

            this.Render(hw);

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

        // find DataPager typeof control in child controls
        public static T FindControl<T>(System.Web.UI.ControlCollection controls) where T : class
        {
            T found = default(T);

            if (controls != null && controls.Count > 0)
            {
                for (int i = 0; i < controls.Count; i++)
                {
                    if (found != null) break;
                    if (controls[i] is T)
                    {
                        found = controls[i] as T;
                        break;
                    }
                    found = FindControl<T>(controls[i].Controls);
                }
            }

            return found;
        }
    }
}
