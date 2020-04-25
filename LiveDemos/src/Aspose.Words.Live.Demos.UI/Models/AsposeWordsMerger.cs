using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using Aspose.Words;
using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Controllers;


namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsMerger class to merge word document
	///</Summary>
	public class AsposeWordsMerger : AsposeWordsBase
  {
    ///<Summary>
    /// Merge method to merge word document
    ///</Summary>
   

		public Response Merge(Document[] docs, string outputType, string sourceFolder)
    {

			if (docs == null)
				return PasswordProtectedResponse;
			if (docs.Length <= 1 || docs.Length > MaximumUploadFiles)
				return MaximumFileLimitsResponse;

			SetDefaultOptions(docs, outputType);
			
			Opts.ResultFileName = $"Merged document{Opts.OutputType}";
			Opts.CreateZip = false;
			Opts.ZipFileName = "Merged document";
			Opts.AppName = " Merger";
			Opts.MethodName = "Merge";
			Opts.FolderName = sourceFolder;

			return  Process((inFilePath, outPath, zipOutFolder) =>
			{
				var doc = docs[0];
				for (var i = 1; i < docs.Length; i++)
					doc.AppendDocument(docs[i], ImportFormatMode.KeepSourceFormatting);
				SaveDocument(doc, outPath, zipOutFolder);
			});


		}   
  }
}