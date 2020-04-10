using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using Aspose.Words.Live.Demos.UI.Config;


namespace Aspose.Words.Live.Demos.UI
{
	public class Global : HttpApplication
	{
		
		protected void Application_Error(object sender, EventArgs e)
		{			
			
		}

		void Application_Start(object sender, EventArgs e)
		{
			AreaRegistration.RegisterAllAreas();
			GlobalConfiguration.Configure(WebApiConfig.Register);
			RouteConfig.RegisterRoutes(RouteTable.Routes);			
			BundleConfig.RegisterBundles(BundleTable.Bundles);
			RegisterCustomRoutes(RouteTable.Routes);
		}
		void Session_Start(object sender, EventArgs e)
		{
			//Check URL to set language resource file
			string _language = "EN";
			
			SetResourceFile(_language);
		}

		private void SetResourceFile(string strLanguage)
		{
			if (Session["AsposeAppResources"] == null)
				Session["AsposeAppResources"] = new GlobalAppHelper(HttpContext.Current, Application, Configuration.ResourceFileSessionName, strLanguage);
		}
		
			void RegisterCustomRoutes(RouteCollection routes)
		{
			routes.RouteExistingFiles = true;
			routes.Ignore("{resource}.axd/{*pathInfo}");
					

			routes.MapRoute(
				name: "Default",
				url: "Default",
				defaults: new { controller = "Home", action = "Default" }
			);
			
			routes.MapRoute(
				"AsposeWordsConversionRoute",
				"{product}/Conversion",
				 new { controller = "Conversion", action = "Conversion" }
			);
			routes.MapRoute(
				"AsposeWordsUnlockRoute",
				"{product}/unlock",
				 new { controller = "Unlock", action = "Unlock" }
			);
			routes.MapRoute(
				"AsposeWordsRedactionRoute",
				"{product}/redaction",
				 new { controller = "Redaction", action = "Redaction" }
			);
			routes.MapPageRoute(
				"AsposeWordsWatermarkRoute",
				"words/watermark",
				"~/Watermark/WatermarkWords.aspx"
			);
			routes.MapRoute(
				"AsposeWordsSearchRoute",
				"{product}/search",
				 new { controller = "Search", action = "Search" }
			);
			routes.MapRoute(
				"AsposeWordsProtectRoute",
				"{product}/protect",
				 new { controller = "Protect", action = "Protect" }
			);
			routes.MapRoute(
				"AsposeWordsMergerRoute",
				"{product}/merger",
				 new { controller = "Merger", action = "Merger" }
			);
			routes.MapPageRoute(
			  "AsposeWordsAssemblyRoute",
			  "{Product}/assembly",
			  "~/Assembly/AssemblyApp.aspx"
			);
			routes.MapRoute(
				"AsposeWordsAnnotationRoute",
				"{product}/annotation",
				 new { controller = "Annotation", action = "Annotation" }
			);
			routes.MapRoute(
				"AsposeWordsRemoveAnnotationRoute",
				"annotation/remove",
				 new { controller = "Annotation", action = "Remove" }
			);
			routes.MapRoute(
				"AsposeWordsParserRoute",
				"{product}/parser",
				 new { controller = "Parser", action = "Parser" }
			);
			routes.MapRoute(
				"AsposeWordsComparisonRoute",
				"{product}/comparison",
				 new { controller = "Comparison", action = "Comparison" }
			);
			routes.MapRoute(
				"AsposeWordsViewerRoute",
				"{product}/viewer",
				 new { controller = "Viewer", action = "Viewer" }
			);
			routes.MapRoute(
				"AsposeWordsMetadataRoute",
				"{product}/metadata",
				 new { controller = "Metadata", action = "Metadata" }
			);
			routes.MapRoute(
				"AsposeWordsEditorRoute",
				"{product}/editor",
				 new { controller = "Editor", action = "Editor" }
			);
			routes.MapRoute(
				"AsposeWordsSplitterRoute",
				"{product}/splitter",
				 new { controller = "Splitter", action = "Splitter" }
			);
			routes.MapPageRoute(
			  "AsposeWordsDefaultViewerRoute",
			  "words/view",
			  "~/ViewerApp/Default.aspx"
			);
			routes.MapPageRoute(
				"AsposeWordsDefaultEditorRoute",
				"{Product}/edit",
				"~/Editor/Default.aspx"
			);
			routes.MapRoute(
				"AsposeWordsSignatureRoute",
				"{Product}/signature",
				 new { controller = "Signature", action = "Signature" }
			);
			routes.MapRoute(
				"AsposeWordsDocumentInfoViewerRoute",
				"{product}/viewer/documentInfo",
				 new { controller = "Viewer", action = "DocumentInfo" }
			);

			routes.MapRoute(
				"DownloadFileRoute",
				"common/download",
				new { controller = "Common", action = "DownloadFile" }				
				
			);
			routes.MapRoute(
				"UploadFileRoute",
				"common/uploadfile",
				new { controller = "Common", action = "UploadFile" }

			);
		}

		private void MapProductToolPageRoute(RouteCollection routes, string routeName, string routeUrl, string physicalFile, string productRegex)
		{
			routes.MapPageRoute(routeName, routeUrl, physicalFile, false, null, new RouteValueDictionary { { "Product", productRegex } });
		}
	}
}
