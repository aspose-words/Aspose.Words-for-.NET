using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.Words.Live.Demos.UI.Config;

namespace Aspose.Words.Live.Demos.UI.Viewer
{
	public partial class Default : BasePage
	{
		public string FileName;
		public string FolderName;
		public string CallbackURL;
		public string DownloadOriginalURL;
		public string ProductName;
		public string AsposeViewerApp = Configuration.AsposeAppLiveDemosPath + "api/AsposeWordsViewer/";
		public string FileDownloadLink = Configuration.AsposeAppLiveDemosPath + "common/download";

		protected void Page_Load(object sender, EventArgs e)
		{
			Product = "words";
			Title = PageProductTitle + " " + Resources["ViewerAPPName"];

			if (!IsPostBack)
			{
				ProductName = Resources["ViewerAPPName"];

				FileName = Request.QueryString["FileName"].Split(',')[0];
				FolderName = Request.QueryString["FolderName"].Split(',')[0];

				if (Request.QueryString["callbackURL"] != null)
				{
					CallbackURL = Request.QueryString["callbackURL"];
					if (!CallbackURL.Contains("?"))
						CallbackURL += HttpUtility.HtmlEncode($"?id={FolderName}&FileName={FileName}");
				}
				else
					CallbackURL = GetRouteUrl("AsposeWordsViewerRoute", null);

				var url = new StringBuilder("");
				url.Append("?FileName=");
				url.Append(HttpUtility.UrlEncode(FileName));
				url.Append("&FolderName=");
				url.Append(FolderName);
				DownloadOriginalURL = url.ToString();
			}
		}
	}
}
