using System;
using Aspose.Words.Saving;
using System.IO;
using System.Web.Http;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq;
using Aspose.Words.Live.Demos.UI.Controllers;

namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsConversion class to convert words files to different formats
	///</Summary>
	public class AsposeWordsConversion : AsposeWordsBase
	{ 

    public Response ConvertFile(Document[] docs, string outputType, string sourceFolder)
		{
			
			if (docs == null)
				return PasswordProtectedResponse;
			if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
				return MaximumFileLimitsResponse;

			SetDefaultOptions(docs, outputType);
			Opts.AppName = " Conversion";
			Opts.MethodName = "ConvertFile";
			Opts.ZipFileName = docs.Length > 1 ? "Converted documents" : Path.GetFileNameWithoutExtension(docs[0].OriginalFileName);
			Opts.FolderName = sourceFolder;

			return  Process((inFilePath, outPath, zipOutFolder) =>
			{
				var tasks = docs.Select(doc => Task.Factory.StartNew(() => SaveDocument(doc, outPath, zipOutFolder))).ToArray();
				Task.WaitAll(tasks);
			});
		}

		
  }
}
