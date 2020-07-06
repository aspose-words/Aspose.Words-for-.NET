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
using Aspose.Words.Properties;
using Aspose.Words.Replacing;
using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Controllers;
using System.Text.RegularExpressions;
using System.Globalization;


namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsRedaction class to merge word document
	///</Summary>
	public class AsposeWordsRedaction : AsposeWordsBase
  {
  
   

		public Response Redact(Document[] docs,  string sourceFolder, string outputType, string searchQuery, string replaceText,
	  bool caseSensitive, bool text, bool comments, bool metadata)
    {

			if (docs == null)
				return PasswordProtectedResponse;
			if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
				return MaximumFileLimitsResponse;

			SetDefaultOptions(docs, outputType);
			Opts.AppName = "Redaction";
			Opts.MethodName = "Redact";
			Opts.ZipFileName = "Redacted documents";
			Opts.FolderName = sourceFolder;

			if (replaceText == null)
				replaceText = "";

			var statusValue = "OK";
			var statusCodeValue = 200;
			var fileProcessingErrorCode = FileProcessingErrorCode.OK;
			var lck = new object();
			var catchedException = false;

			if (IsValidRegex(searchQuery))
			{
				var regex = new Regex(searchQuery, caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase);
				var matchesFound = new int[docs.Length];
				var tasks = Enumerable.Range(0, docs.Length).Select(i => Task.Factory.StartNew(() =>
				{
					try
					{
						if (text || comments)
						{
							var findings = new RedactionCallback(text, comments);
							var options = new FindReplaceOptions()
							{
								ReplacingCallback = findings,
								Direction = FindReplaceDirection.Forward,
								MatchCase = caseSensitive
							};
							docs[i].Range.Replace(regex, replaceText, options);
							matchesFound[i] += findings.MatchesFound;
						}

						if (metadata)
							matchesFound[i] += ProcessMetadata(docs[i], regex, replaceText);
					}
					catch (Exception ex)
					{
						lock (lck)
							catchedException = true;
						Console.WriteLine(ex.Message);
					}
				})).ToArray();
				Task.WaitAll(tasks);

				if (!catchedException)
				{
					if (matchesFound.Sum() > 0)
						return  Process((inFilePath, outPath, zipOutFolder) =>
						{
							foreach (var doc in docs)
								SaveDocument(doc, outPath, zipOutFolder);
						});

					fileProcessingErrorCode = FileProcessingErrorCode.NoSearchResults;
				}
				else
				{
					statusCodeValue = 500;
					statusValue = "500 Exception during processing";
				}
			}
			else
				fileProcessingErrorCode = FileProcessingErrorCode.WrongRegExp;

			return new Response
			{
				Status = statusValue,
				StatusCode = statusCodeValue,
				FileProcessingErrorCode = fileProcessingErrorCode
			};


		}
		private static int ProcessMetadata(Document doc, Regex regex, string replaceText)
		{
			var matchesFound = 0;
			foreach (var prop in doc.BuiltInDocumentProperties.Union(doc.CustomDocumentProperties))
			{
				var value = Convert.ToString(prop.Value);
				if (regex.IsMatch(value))
					try
					{
						value = regex.Replace(value, replaceText);
						switch (prop.Type)
						{
							case PropertyType.String:
								prop.Value = value;
								break;
							case PropertyType.Boolean:
								prop.Value = Convert.ToBoolean(value);
								break;
							case PropertyType.Number:
								prop.Value = Convert.ToInt32(value);
								break;
							case PropertyType.Double:
								prop.Value = Convert.ToDouble(value, CultureInfo.InvariantCulture);
								break;
							case PropertyType.DateTime:
								prop.Value = Convert.ToDateTime(value);
								break;
						}
						matchesFound++;
					}
					catch { }
			}
			return matchesFound;
		}
		///<Summary>
		/// RedactionCallback class to redact word document
		///</Summary>
		public class RedactionCallback : IReplacingCallback
		{
			///<Summary>
			/// MatchesFound variable
			///</Summary>
			public int MatchesFound;
			///<Summary>
			/// Text variable
			///</Summary>
			public bool Text;
			///<Summary>
			/// Comments variable
			///</Summary>
			public bool Comments;
			///<Summary>
			/// init RedactionCallback
			///</Summary>
			public RedactionCallback(bool text, bool comments)
			{
				Text = text;
				Comments = comments;
			}

			/// <summary>
			/// Recursion check if ParentNode is a comment
			/// </summary>
			/// <param name="node"></param>
			/// <returns></returns>
			public bool IsComment(Node node)
			{
				if (node == null)
					return false;
				if (node.NodeType == NodeType.Comment)
					return true;
				return IsComment(node.ParentNode);
			}
			///<Summary>
			/// Replacing
			///</Summary>
			public ReplaceAction Replacing(ReplacingArgs args)
			{
				if (Text && Comments)
				{
					MatchesFound += 1;
					return ReplaceAction.Replace;
				}

				var isComment = IsComment(args.MatchNode);
				if (Text && isComment || Comments && !isComment)
					return ReplaceAction.Skip;

				MatchesFound += 1;
				return ReplaceAction.Replace;
			}
		}

	}
}