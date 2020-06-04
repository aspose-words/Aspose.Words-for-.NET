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
	/// AsposeWordsParser class to merge word document
	///</Summary>
	public class AsposeWordsParser : AsposeWordsBase
  {

		public Response Parse(Document[] docs, string outputType, string sourceFolder)
		{
			
			if (docs == null)
				return PasswordProtectedResponse;
			if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
				return MaximumFileLimitsResponse;

			SetDefaultOptions(docs, outputType);
			Opts.AppName = " Parse";
			Opts.MethodName = "Parse";
			Opts.ZipFileName = docs.Length > 1 ? "Parser" : Path.GetFileNameWithoutExtension(docs[0].OriginalFileName);
			Opts.OutputType = ".txt";
			Opts.CreateZip = true;
			Opts.FolderName = sourceFolder;

			return  Process((inFilePath, outPath, zipOutFolder) =>
			{
				var tasks = docs.Select(x => Task.Factory.StartNew(() => ParseDocument(x, zipOutFolder))).ToArray();
				Task.WaitAll(tasks);
			});
		}

		/// <summary>
		/// Parse Document
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="outPath"></param>
		private void ParseDocument(Document doc, string outPath)
		{
			try
			{
				var (filename, folder) = PrepareFolder(doc, outPath);
				var comments = doc.GetChildNodes(NodeType.Comment, true);
				comments.Clear();
				doc.Save($"{folder}/{Path.GetFileNameWithoutExtension(filename)}.txt", SaveFormat.Text);
				ExtractImages(doc, folder);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
		
  }
}