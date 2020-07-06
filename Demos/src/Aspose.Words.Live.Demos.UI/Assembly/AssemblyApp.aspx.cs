using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.Words.Live.Demos.UI.Config;

namespace Aspose.Words.Live.Demos.UI
{
	public partial class AssemblyApp : BasePage
	{
		protected void Page_Load(object sender, EventArgs e)
		{
			var fileMaxSize = "100MB";

			var productTitle = Resources["Aspose" + TitleCase(Product)];

			Page.Title = Resources[Product + "AssemblyPageTitle"];
			Page.MetaDescription = Resources[Product + "AssemblyMetaDescription"];
			ProductTitle.InnerText = Resources[Product + "AssemblyH1"];
			ProductTitleSub.InnerText = Resources[Product + "AssemblyH4"];

			AsposeProductTitle.InnerText = productTitle + " " + Resources["AssemblyAPPName"];
			ProductImage.Src = "~/img/aspose-" + Product + "-app.png";
			PoweredBy.InnerText = productTitle + ". ";
			PoweredBy.HRef = "https://products.aspose.com/" + Product;


			ViewState["fileMaxSize"] = fileMaxSize;
			ViewState["product"] = Product;
			ViewState["ASPOSE_FILEDOWNLOADLINK"] = Configuration.FileDownloadLink + "?";
			ViewState["ASPOSE_ASSEMBLY_API"] = Configuration.AsposeAppLiveDemosPath + "api/AsposeWordsAssembly/";

			var validFileExtensions = GetValidFileExtensions(Resources[Product + "ValidationExpression"]);
			ViewState["validFileExtensions"] = validFileExtensions;

			// Template
			ViewState["templateValidFileExtensions"] = Resources[Product + "ValidationExpression"].Replace('|', ',');
			TemplateInvalidExtension.InnerText = Resources["InvalidFileExtension"] + " " + validFileExtensions;
			TemplateInvalidSize.InnerText = Resources["ErrorFileSizeMax"] + " " + fileMaxSize;

			// DataSource
			ViewState["dataValidFileExtensions"] = Resources[Product + "AssemblyValidationExpression"].Replace('|', ',');
			DataInvalidExtension.InnerText = Resources["InvalidFileExtension"] + " " + validFileExtensions;
			DataInvalidSize.InnerText = Resources["ErrorFileSizeMax"] + " " + fileMaxSize;

			
		}//
	}
}
