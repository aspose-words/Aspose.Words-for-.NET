using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Optimization;
using System.Web.UI;

namespace Aspose.Words.Live.Demos.UI
{
	public class BundleConfig
	{
		// For more information on Bundling, visit https://go.microsoft.com/fwlink/?LinkID=303951
		public static void RegisterBundles(BundleCollection bundles)
		{
			bundles.Add(new ScriptBundle("~/bundles/WebFormsJs").Include(
							"~/js/WebForms/WebForms.js",
							"~/js/WebForms/WebUIValidation.js",
							"~/js/WebForms/MenuStandards.js",
							"~/js/WebForms/Focus.js",
							"~/js/WebForms/GridView.js",
							"~/js/WebForms/DetailsView.js",
							"~/js/WebForms/TreeView.js",
							"~/js/WebForms/WebParts.js"));

			// Order is very important for these files to work, they have explicit dependencies
			bundles.Add(new ScriptBundle("~/bundles/MsAjaxJs").Include(
					"~/js/WebForms/MsAjax/MicrosoftAjax.js",
					"~/js/WebForms/MsAjax/MicrosoftAjaxApplicationServices.js",
					"~/js/WebForms/MsAjax/MicrosoftAjaxTimer.js",
					"~/js/WebForms/MsAjax/MicrosoftAjaxWebForms.js"));

			// Use the Development version of Modernizr to develop with and learn from. Then, when you’re
			// ready for production, use the build tool at https://modernizr.com to pick only the tests you need
			bundles.Add(new ScriptBundle("~/bundles/modernizr").Include(
							"~/js/modernizr-*"));

			bundles.Add(
				new ScriptBundle("~/bundles/jquery").
					Include(
						//"~/Scripts/jquery-{version}.js",
						"~/Scripts/jquery.form.min.js",
						"~/Scripts/jquery.unobtrusive-ajax.min.js"
					)
			);

			ScriptManager.ScriptResourceMapping.AddDefinition(
				"respond",
				new ScriptResourceDefinition
				{
					Path = "~/js/respond.min.js",
					DebugPath = "~/js/respond.js",
				});


			//bundles.Add(
			//	new ScriptBundle("~/bundles/AsposeShared")
			//	.Include(					
			//		"~/Scripts/Shared/Alert.js",
			//		"~/Scripts/Shared/Resources.js",
			//		"~/Scripts/Shared/DownloadResult.js",
			//		"~/Scripts/Shared/Loader.js",
			//		"~/Scripts/Shared/UploadFile.js",
			//		"~/Scripts/Shared/Work.js"
			//	)
			//);
		}
	}
}
