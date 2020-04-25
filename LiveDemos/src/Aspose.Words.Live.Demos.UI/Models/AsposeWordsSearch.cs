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
using Aspose.Words.Fields;
using Aspose.Words.Replacing;
using System.Text.RegularExpressions;
using System.Drawing;


namespace Aspose.Words.Live.Demos.UI.Models
{

	///<Summary>
	/// AsposeWordsSearch class to perform search operations on words document
	///</Summary>
	public class AsposeWordsSearch : AsposeWordsBase
	{
		// Yellow Color
		private static readonly Color HiglightedColor = Color.FromArgb(246, 247, 146);


		/// <summary>
		/// Search documents
		/// </summary>

		public Response Search(Document[] docs, string sourceFolder, string query)
    {
      
      if (docs == null)
        return PasswordProtectedResponse;
      if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
        return MaximumFileLimitsResponse;

      SetDefaultOptions(docs, "");
      Opts.AppName = "Search";
	  Opts.MethodName = "Search";
			Opts.FolderName = sourceFolder;
      Opts.OutputType = ".docx";
      Opts.ResultFileName = "Search Results";
      Opts.CreateZip = false;

      var statusValue = "OK";
      var statusCodeValue = 200;
      var fileProcessingErrorCode = FileProcessingErrorCode.OK;

      if (IsValidRegex(query))
        try
        {
          var findings = new FindCallback();
          var options = new FindReplaceOptions()
          {
            ReplacingCallback = findings,
            Direction = FindReplaceDirection.Forward,
            MatchCase = false
          };
          foreach (var doc in docs)
            doc.Range.Replace(new Regex(query, RegexOptions.IgnoreCase), "", options);

					if (findings.MatchesFound > 0)
						return  Process((inFilePath, outPath, zipOutFolder) => findings.Save(outPath));

					fileProcessingErrorCode = FileProcessingErrorCode.NoSearchResults;
        }
        catch (Exception ex)
        {
					Console.WriteLine(ex.Message);
          statusCodeValue = 500;
          statusValue = "500 " + ex.Message;
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

    
		///<Summary>
		/// FindCallback
		///</Summary>
		public class FindCallback : IReplacingCallback
		{
			///<Summary>
			/// initialize document
			///</Summary>
			public Document Doc = new Document();

			///<Summary>
			/// get or set MatchesFound
			///</Summary>
			public int MatchesFound { get; private set; }

			/// <summary>
			/// Nodes that have been performed with highlighting results and Nodes for a result file
			/// </summary>
			public readonly Dictionary<Node, List<ReplacingArgs>> MatchedNodes = new Dictionary<Node, List<ReplacingArgs>>();

			///<Summary>
			/// Replacing
			///</Summary>
			public ReplaceAction Replacing(ReplacingArgs args)
			{
				MatchesFound += 1;
				if (MatchedNodes.ContainsKey(args.MatchNode))
				{
					MatchedNodes[args.MatchNode].Add(args);
					return ReplaceAction.Skip;
				}

				MatchedNodes.Add(args.MatchNode, new List<ReplacingArgs>() { args });
				return ReplaceAction.Skip;
			}

			/// <summary>
			/// Save the document with highlighted search results
			/// </summary>
			/// <param name="filename"></param>
			public void Save(string filename)
			{
				var builder = new DocumentBuilder(Doc);
				builder.MoveToSection(0);
				builder.Writeln("Matches found: " + MatchesFound);
				builder.Writeln();

				var nodeCount = 1;
				foreach (var kvp in MatchedNodes.Where(x => !(x.Key.PreviousSibling is FieldStart))) // e.g. not include "HYPERLINK" and other fields
        {
          var documentName = Path.GetFileName(((Document) kvp.Key.Document).OriginalFileName);
					builder.Writeln($"Document: {documentName}  Result{nodeCount}:");

					var node = Doc.ImportNode(kvp.Key, true, ImportFormatMode.KeepSourceFormatting);
					var runs = Highlight(node, kvp.Value);
					foreach (var run in runs)
						builder.InsertNode(run);

					nodeCount++;
					builder.Writeln();
					builder.Writeln();
				}

				Doc.Save(filename);
			}


			private List<Run> Highlight(Node node, List<ReplacingArgs> replacingArgs)
			{
				var runs = new List<Run>() { (Run)node };

				// This array is used to store all nodes of the match for further highlighting.
				var highlights = new List<Node>();

				// Sum of previous matchOffsets
				var cursor = 0;

				if (replacingArgs.Count > 1)
					Console.WriteLine();

				foreach (var args in replacingArgs)
				{
					node = runs.Last();
					var matchOffset = args.MatchOffset - cursor;
					cursor += matchOffset + args.Match.Value.Length;

					// The first (and may be the only) run can contain text before the match, 
					// In this case it is necessary to split the run.
					if (matchOffset > 0)
						node = SplitRun(runs, (Run)node, matchOffset);

					// Find all runs that contain parts of the match string.
					var remainingLength = args.Match.Value.Length;
					while (remainingLength > 0 &&
								 node != null &&
								 node.GetText().Length <= remainingLength)
					{
						highlights.Add(node);
						remainingLength = remainingLength - node.GetText().Length;

						// Select the next Run node. 
						// Have to loop because there could be other nodes such as BookmarkStart etc.
						do
						{
							node = node.NextSibling;
						} while (node != null && node.NodeType != NodeType.Run);
					}

					// Split the last run that contains the match if there is any text left.
					if (node != null && remainingLength > 0)
					{
						SplitRun(runs, (Run)node, remainingLength);
						highlights.Add(node);
					}
				}

				// Now highlight all runs in the sequence.
				foreach (Run run in highlights)
					run.Font.HighlightColor = HiglightedColor;

				return runs;
			}

			/// <summary>
			/// Splits text of the specified run into two runs.
			/// Inserts the new run into the list.
			/// </summary>
			private static Run SplitRun(List<Run> runs, Run run, int position)
			{
				var afterRun = (Run)run.Clone(true);
				afterRun.Text = run.Text.Substring(position);
				runs.Add(afterRun);
				run.Text = run.Text.Substring(0, position);
				return afterRun;
			}
		}
	}
}
