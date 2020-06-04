using Aspose.Words.Live.Demos.UI.Config;
using System;
using System.Text;
using System.Web;
using System.Web.UI;
using Aspose.Words.Live.Demos.UI.Models;

namespace Aspose.Words.Live.Demos.UI.Editor
{
    public partial class Default : BasePage
    {
        public string FileName;
        public string FolderName;
        public string CallbackURL;
        public string DownloadOriginalURL;
        public string ProductName;
        public string AsposeEditorApp = Configuration.AsposeAppLiveDemosPath + "api/EditorHelper/";
        public string FileDownloadLink = Configuration.AsposeAppLiveDemosPath + "common/download";

        public string DownloadMainType;

        protected void Page_Load(object sender, EventArgs e)
        {
            Title = PageProductTitle + " " + Resources["EditorAPPName"];
			

            if (!IsPostBack)
            {
                ProductName = Resources["EditorAPPName"];
                Page.Title = Resources[Product + "EditorPageTitle"];
                Page.MetaDescription = Resources[Product + "EditorMetaDescription"];
                FileName = Request.QueryString["FileName"];
                FolderName = Request.QueryString["FolderName"];

                if (Request.QueryString["callbackURL"] != null)
                    CallbackURL = Request.QueryString["callbackURL"];
                else
                    CallbackURL = GetRouteUrl("AsposeToolsEditorApp", new { Product });

                var url = new StringBuilder(Configuration.AsposeAppLiveDemosPath + "common/download");
                url.Append("?FileName=");
                url.Append(HttpUtility.UrlEncode(FileName));
                url.Append("&FolderName=");
                url.Append(FolderName);
                DownloadOriginalURL = url.ToString();
            }

            var downloadItemsBuilder = new StringBuilder();

            switch (Product)
            {
	            case "words":
		            Page.Title = string.Format(Page.Title, "Word");
		            Page.MetaDescription = string.Format(Page.MetaDescription, "Word");
		            downloadItemsBuilder.AppendLine(CreateListItem("DOCX"));
		            downloadItemsBuilder.AppendLine(CreateListItem("PDF"));
		            break;
	          
	            default:
		            break;
            }

            downloadItemsBuilder.AppendLine(CreateListItem("HTML"));

            litToDropdownItem.Text = downloadItemsBuilder.ToString();
        }

        string CreateListItem(string extension)
        {
            return $@"<li><a href=""#"" ng-click=""Download('.{extension.ToLowerInvariant()}')"">as {extension.ToUpperInvariant()}</a></li>";
        }
    }
}
