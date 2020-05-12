using System;
using Aspose.Words.Saving;
using System.IO;
using System.Web.Http;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq;
using Aspose.Words.Live.Demos.UI.Controllers;
using System.Web.Mvc;

namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsComparison class to convert words files to different formats
	///</Summary>
	public class AsposeWordsComparison : AsposeWordsBase
	{ 

    public Response Compare(Document[] docs, string sourceFolder)
		{

			if (docs == null)
				return PasswordProtectedResponse;
			if (docs.Length != 2)
				return new Response()
				{
					Status = "Number of files should be 2",
					StatusCode = 500
				};

			SetDefaultOptions(docs, "");
			Opts.AppName = "Comparison";
			
			Opts.MethodName = "Compare";
			Opts.ResultFileName = $"{Path.GetFileNameWithoutExtension(docs[0].OriginalFileName)} compared to {Path.GetFileNameWithoutExtension(docs[1].OriginalFileName)}.docx";
			Opts.OutputType = "docx";
			Opts.CreateZip = false;
			Opts.FolderName = sourceFolder;

			return  Process((inFilePath, outPath, zipOutFolder) =>
			{
				docs[0].Revisions.AcceptAll();
				docs[1].Revisions.AcceptAll();
				docs[0].Compare(docs[1], "a", DateTime.Now);
				docs[0].Save(outPath);
			});
		}

		
  }
}
