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
	/// AsposeWordsAnnotation class to merge word document
	///</Summary>
	public class AsposeWordsAnnotation : AsposeWordsBase
  {
  
   

		public Response Remove(Document[] docs,  string sourceFolder)
    {

			if (docs == null)
				return PasswordProtectedResponse;
			if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
				return MaximumFileLimitsResponse;

			SetDefaultOptions(docs, "");
			Opts.AppName = "Annotation";
			Opts.MethodName = "Remove";
			Opts.ZipFileName = docs.Length > 1 ? "Removed Annotations" : Path.GetFileNameWithoutExtension(docs[0].OriginalFileName);
			Opts.CreateZip = true;
			Opts.FolderName = sourceFolder;

			return  Process((inFilePath, outPath, zipOutFolder) =>
			{
				var tasks = docs.Select(x => Task.Factory.StartNew(() => RemoveAnnotations(x, zipOutFolder))).ToArray();
				Task.WaitAll(tasks);
			});


		}
		/// <summary>
		/// Remove annotations in document
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="outPath"></param>
		private void RemoveAnnotations(Document doc, string outPath)
		{
			try
			{
				var (filename, folder) = PrepareFolder(doc, outPath);
				var comments = doc.GetChildNodes(NodeType.Comment, true);
				var collectedComments = comments.Select(x => x as Comment).Select(
				  comment => comment.Author + " " + comment.DateTime + Environment.NewLine + comment.ToString(SaveFormat.Text)).ToArray();
				comments.Clear();
				doc.Save($"{folder}/{filename}");
				if (collectedComments.Length > 0)
					File.WriteAllLines($"{folder}/comments.txt", collectedComments);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}
	}
}