using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Services;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsSplitter class to convert words files to different formats
	///</Summary>
	public class AsposeWordsSplitter : AsposeWordsBase
	{
		///<Summary>
		/// Split method
		///</Summary>		
		public Response Split(Document[] docs, string sourceFolder, string outputType, int splitType, string pars = null)
		{
			
			if (docs.Length == 0 || docs.Length > MaximumUploadFiles)
				return MaximumFileLimitsResponse;

			SetDefaultOptions(docs, outputType);
			Opts.AppName = "Splitter";
			Opts.MethodName = "Split";
			Opts.CreateZip = true;
			Opts.ZipFileName = "Splitted documents";
			Opts.FolderName = sourceFolder;

			return  Process((inFilePath, outPath, zipOutFolder) =>
			{
				Action<Document, string, string> action;
				switch (splitType)
				{
					case 2:
						action = SplitOddEven;
						break;
					case 3:
						action = SplitPageNumber;
						break;
					case 4:
						action = SplitPageRange;
						break;
					default:
						action = SplitAllPages;
						break;
				}
				var tasks = docs.Select(x => Task.Factory.StartNew(() => action(x, zipOutFolder, pars))).ToArray();
				Task.WaitAll(tasks);
			});
		}

		private (DocumentPageSplitter, string, string) PrepareValues(Document doc)
		{
			var splitter = new DocumentPageSplitter(doc);
			var filename = Path.GetFileNameWithoutExtension(doc.OriginalFileName);
			var extension = Opts.OutputType == SaveAsOriginalName
			  ? Path.GetExtension(doc.OriginalFileName)
			  : Opts.OutputType;
			return (splitter, filename, extension);
		}

		private void SplitAllPages(Document doc, string outPath, string pars = null)
		{
			try
			{
				var (splitter, filename, extension) = PrepareValues(doc);
				for (var i = 1; i <= doc.PageCount; i++)
				{
					var pageDoc = splitter.GetDocumentOfPage(i);
					pageDoc.Save(Path.Combine(outPath, $"{filename} {i:00}{extension}"));
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		private void SplitOddEven(Document doc, string outPath, string pars = null)
		{
			try
			{
				var (splitter, filename, extension) = PrepareValues(doc);
				var oddDoc = new Document();
				var evenDoc = new Document();
				for (var i = 1; i <= doc.PageCount; i++)
				{
					var pageDoc = splitter.GetDocumentOfPage(i);
					if (i % 2 == 1)
						oddDoc.AppendDocument(pageDoc, ImportFormatMode.KeepSourceFormatting);
					else
						evenDoc.AppendDocument(pageDoc, ImportFormatMode.KeepSourceFormatting);
				}
				oddDoc.Sections[0].Remove();
				SaveDocument(oddDoc, Path.Combine(outPath, $"{filename} odd{extension}"));
				if (doc.PageCount >= 2)
				{
					evenDoc.Sections[0].Remove();
					SaveDocument(evenDoc, Path.Combine(outPath, $"{filename} even{extension}"));
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		private void SplitPageNumber(Document doc, string outPath, string pars)
		{
			try
			{
				var number = Convert.ToInt32(pars);
				if (number <= 0)
					return;
				var (splitter, filename, extension) = PrepareValues(doc);
				var last = number > 1
				  ? System.Math.Floor((double)doc.PageCount / number)
				  : doc.PageCount - 1;
				for (var i = 0; i <= last; i++)
				{
					var start = i * number + 1;
					var end = i < last ? (i + 1) * number : doc.PageCount;
					var pageDoc = splitter.GetDocumentOfPageRange(start, end);
					pageDoc.Save(Path.Combine(outPath, $"{filename} {i + 1:00}{extension}"));
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		private void SplitPageRange(Document doc, string outPath, string pars)
		{
			try
			{
				var (splitter, filename, extension) = PrepareValues(doc);
				var newDoc = new Document();
				var lst = new List<int>();
				var values = pars.Split(',');
				foreach (var value in values)
					if (!value.Contains("-"))
						lst.Add(Convert.ToInt32(value.Trim()));
					else
					{
						var v = value.Split('-').Select(x => Convert.ToInt32(x.Trim())).ToArray();
						if (v[0] <= v[1])
							for (var i = v[0]; i <= v[1]; i++)
								if (i > 0)
									lst.Add(i);
					}

				foreach (var i in lst)
					if (i <= doc.PageCount)
					{
						var pageDoc = splitter.GetDocumentOfPage(i);
						newDoc.AppendDocument(pageDoc, ImportFormatMode.KeepSourceFormatting);
					}
				newDoc.Sections[0].Remove();
				SaveDocument(newDoc, Path.Combine(outPath, $"{filename} splitted{extension}"));
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		/// <summary>
		/// Splits a document into multiple documents, one per page.
		/// </summary>
		internal class DocumentPageSplitter
		{
			private readonly PageNumberFinder pageNumberFinder;

			/// <summary>
			/// Initializes a new instance of the <see cref="DocumentPageSplitter"/> class.
			/// This method splits the document into sections so that each page begins and ends at a section boundary.
			/// It is recommended not to modify the document afterwards.
			/// </summary>
			/// <param name="source">source document</param>
			public DocumentPageSplitter(Document source)
			{
				this.pageNumberFinder = PageNumberFinderFactory.Create(source);
			}

			/// <summary>
			/// Gets the document this instance works with.
			/// </summary>
			private Document Document => this.pageNumberFinder.Document;

			/// <summary>
			/// Gets the document of a page.
			/// </summary>
			/// <param name="pageIndex">
			/// 1-based index of a page.
			/// </param>
			/// <returns>
			/// The <see cref="Document"/>.
			/// </returns>
			public Document GetDocumentOfPage(int pageIndex)
			{
				return this.GetDocumentOfPageRange(pageIndex, pageIndex);
			}

			/// <summary>
			/// Gets the document of a page range.
			/// </summary>
			/// <param name="startIndex">
			/// 1-based index of the start page.
			/// </param>
			/// <param name="endIndex">
			/// 1-based index of the end page.
			/// </param>
			/// <returns>
			/// The <see cref="Document"/>.
			/// </returns>
			public Document GetDocumentOfPageRange(int startIndex, int endIndex)
			{
				Document result = (Document)this.Document.Clone(false);
				foreach (var section in this.pageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex, NodeType.Section))
				{
					result.AppendChild(result.ImportNode(section, true));
				}

				return result;
			}
		}

		/// <summary>
		/// Provides methods for extracting nodes of a document which are rendered on a specified pages.
		/// </summary>
		public class PageNumberFinder
		{
			// Maps node to a start/end page numbers. This is used to override baseline page numbers provided by collector when document is split.
			private readonly IDictionary<Node, int> nodeStartPageLookup = new Dictionary<Node, int>();
			private readonly IDictionary<Node, int> nodeEndPageLookup = new Dictionary<Node, int>();
			private readonly LayoutCollector collector;

			// Maps page number to a list of nodes found on that page.
			private IDictionary<int, IList<Node>> reversePageLookup;

			/// <summary>
			/// Initializes a new instance of the <see cref="PageNumberFinder"/> class.
			/// </summary>
			/// <param name="collector">A collector instance which has layout model records for the document.</param>
			public PageNumberFinder(LayoutCollector collector)
			{
				this.collector = collector;
			}

			/// <summary>
			/// Gets the document this instance works with.
			/// </summary>
			public Document Document => this.collector.Document;

			/// <summary>
			/// Retrieves 1-based index of a page that the node begins on.
			/// </summary>
			/// <param name="node">
			/// The node.
			/// </param>
			/// <returns>
			/// Page index.
			/// </returns>
			public int GetPage(Node node)
			{
				return this.nodeStartPageLookup.ContainsKey(node)
						   ? this.nodeStartPageLookup[node]
						   : this.collector.GetStartPageIndex(node);
			}

			/// <summary>
			/// Retrieves 1-based index of a page that the node ends on.
			/// </summary>
			/// <param name="node">
			/// The node.
			/// </param>
			/// <returns>
			/// Page index.
			/// </returns>
			public int GetPageEnd(Node node)
			{
				return this.nodeEndPageLookup.ContainsKey(node)
						   ? this.nodeEndPageLookup[node]
						   : this.collector.GetEndPageIndex(node);
			}

			/// <summary>
			/// Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
			/// </summary>
			/// <param name="node">
			/// The node.
			/// </param>
			/// <returns>
			/// Page index.
			/// </returns>
			public int PageSpan(Node node)
			{
				return this.GetPageEnd(node) - this.GetPage(node) + 1;
			}

			/// <summary>
			/// Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
			/// </summary>
			/// <param name="startPage">
			/// The start Page.
			/// </param>
			/// <param name="endPage">
			/// The end Page.
			/// </param>
			/// <param name="nodeType">
			/// The node Type.
			/// </param>
			/// <returns>
			/// Collection of nodes.
			/// </returns>
			public IList<Node> RetrieveAllNodesOnPages(int startPage, int endPage, NodeType nodeType)
			{
				if (startPage < 1 || startPage > this.Document.PageCount)
				{
					throw new InvalidOperationException("'startPage' is out of range");
				}

				if (endPage < 1 || endPage > this.Document.PageCount || endPage < startPage)
				{
					throw new InvalidOperationException("'endPage' is out of range");
				}

				this.CheckPageListsPopulated();
				IList<Node> pageNodes = new List<Node>();
				for (int page = startPage; page <= endPage; page++)
				{
					// Some pages can be empty.
					if (!this.reversePageLookup.ContainsKey(page))
					{
						continue;
					}

					foreach (Node node in this.reversePageLookup[page])
					{
						if (node.ParentNode != null
							&& (nodeType == NodeType.Any || node.NodeType == nodeType)
							&& !pageNodes.Contains(node))
						{
							pageNodes.Add(node);
						}
					}
				}

				return pageNodes;
			}

			/// <summary>
			/// Splits nodes which appear over two or more pages into separate nodes so that they still appear in the same way
			/// but no longer appear across a page.
			/// </summary>
			public void SplitNodesAcrossPages()
			{
				foreach (Paragraph paragraph in this.Document.GetChildNodes(NodeType.Paragraph, true))
				{
					if (this.GetPage(paragraph) != this.GetPageEnd(paragraph))
					{
						this.SplitRunsByWords(paragraph);
					}
				}

				this.ClearCollector();

				// Visit any composites which are possibly split across pages and split them into separate nodes.
				this.Document.Accept(new SectionSplitter(this));
			}

			/// <summary>
			/// This is called by <see cref="SectionSplitter"/> to update page numbers of split nodes.
			/// </summary>
			/// <param name="node">
			/// The node.
			/// </param>
			/// <param name="startPage">
			/// The start Page.
			/// </param>
			/// <param name="endPage">
			/// The end Page.
			/// </param>
			internal void AddPageNumbersForNode(Node node, int startPage, int endPage)
			{
				if (startPage > 0)
				{
					this.nodeStartPageLookup[node] = startPage;
				}

				if (endPage > 0)
				{
					this.nodeEndPageLookup[node] = endPage;
				}
			}

			private static bool IsHeaderFooterType(Node node)
			{
				return node.NodeType == NodeType.HeaderFooter || node.GetAncestor(NodeType.HeaderFooter) != null;
			}

			private void CheckPageListsPopulated()
			{
				if (this.reversePageLookup != null)
				{
					return;
				}

				this.reversePageLookup = new Dictionary<int, IList<Node>>();

				// Add each node to a list which represent the nodes found on each page.
				foreach (Node node in this.Document.GetChildNodes(NodeType.Any, true))
				{
					// Headers/Footers follow sections. They are not split by themselves.
					if (IsHeaderFooterType(node))
					{
						continue;
					}

					int startPage = this.GetPage(node);
					int endPage = this.GetPageEnd(node);
					for (int page = startPage; page <= endPage; page++)
					{
						if (!this.reversePageLookup.ContainsKey(page))
						{
							this.reversePageLookup.Add(page, new List<Node>());
						}

						this.reversePageLookup[page].Add(node);
					}
				}
			}

			private void SplitRunsByWords(Paragraph paragraph)
			{
				foreach (Run run in paragraph.Runs)
				{
					if (this.GetPage(run) == this.GetPageEnd(run))
					{
						continue;
					}

					this.SplitRunByWords(run);
				}
			}

			private void SplitRunByWords(Run run)
			{
				var words = run.Text.Split(' ').Reverse();

				foreach (var word in words)
				{
					var pos = run.Text.Length - word.Length - 1;
					if (pos > 1)
					{
						SplitRun(run, run.Text.Length - word.Length - 1);
					}
				}
			}

			/// <summary>
			/// Splits text of the specified run into two runs.
			/// Inserts the new run just after the specified run.
			/// </summary>
			private static Run SplitRun(Run run, int position)
			{
				Run afterRun = (Run)run.Clone(true);
				afterRun.Text = run.Text.Substring(position);
				run.Text = run.Text.Substring(0, position);
				run.ParentNode.InsertAfter(afterRun, run);
				return afterRun;
			}

			private void ClearCollector()
			{
				this.collector.Clear();
				this.Document.UpdatePageLayout();

				this.nodeStartPageLookup.Clear();
				this.nodeEndPageLookup.Clear();
			}
		}

		internal static class PageNumberFinderFactory
		{
			public static PageNumberFinder Create(Document document)
			{
				LayoutCollector layoutCollector = new LayoutCollector(document);
				document.UpdatePageLayout();
				PageNumberFinder pageNumberFinder = new PageNumberFinder(layoutCollector);
				pageNumberFinder.SplitNodesAcrossPages();
				return pageNumberFinder;
			}
		}

		/// <summary>
		/// Splits a document into multiple sections so that each page begins and ends at a section boundary.
		/// </summary>
		internal class SectionSplitter : DocumentVisitor
		{
			private readonly PageNumberFinder pageNumberFinder;

			public SectionSplitter(PageNumberFinder pageNumberFinder)
			{
				this.pageNumberFinder = pageNumberFinder;
			}

			public override VisitorAction VisitParagraphStart(Paragraph paragraph)
			{
				return this.ContinueIfCompositeAcrossPageElseSkip(paragraph);
			}

			public override VisitorAction VisitTableStart(Table table)
			{
				return this.ContinueIfCompositeAcrossPageElseSkip(table);
			}

			public override VisitorAction VisitRowStart(Row row)
			{
				return this.ContinueIfCompositeAcrossPageElseSkip(row);
			}

			public override VisitorAction VisitCellStart(Cell cell)
			{
				return this.ContinueIfCompositeAcrossPageElseSkip(cell);
			}

			public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag sdt)
			{
				return this.ContinueIfCompositeAcrossPageElseSkip(sdt);
			}

			public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
			{
				return this.ContinueIfCompositeAcrossPageElseSkip(smartTag);
			}

			public override VisitorAction VisitSectionStart(Section section)
			{
				Section previousSection = (Section)section.PreviousSibling;

				// If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
				// extracted document if the previous section is missing.
				if (previousSection != null)
				{
					HeaderFooterCollection previousHeaderFooters = previousSection.HeadersFooters;
					if (!section.PageSetup.RestartPageNumbering)
					{
						section.PageSetup.RestartPageNumbering = true;
						section.PageSetup.PageStartingNumber = previousSection.PageSetup.PageStartingNumber + this.pageNumberFinder.PageSpan(previousSection);
					}

					foreach (HeaderFooter previousHeaderFooter in previousHeaderFooters)
					{
						if (section.HeadersFooters[previousHeaderFooter.HeaderFooterType] == null)
						{
							HeaderFooter newHeaderFooter = (HeaderFooter)previousHeaderFooters[previousHeaderFooter.HeaderFooterType].Clone(true);
							section.HeadersFooters.Add(newHeaderFooter);
						}
					}
				}

				return this.ContinueIfCompositeAcrossPageElseSkip(section);
			}

			public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
			{
				this.SplitComposite(smartTag);
				return VisitorAction.Continue;
			}

			public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
			{
				this.SplitComposite(sdt);
				return VisitorAction.Continue;
			}

			public override VisitorAction VisitCellEnd(Cell cell)
			{
				this.SplitComposite(cell);
				return VisitorAction.Continue;
			}

			public override VisitorAction VisitRowEnd(Row row)
			{
				this.SplitComposite(row);
				return VisitorAction.Continue;
			}

			public override VisitorAction VisitTableEnd(Table table)
			{
				this.SplitComposite(table);
				return VisitorAction.Continue;
			}

			public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
			{
				// If paragraph contains only section break, add fake run into 
				if (paragraph.IsEndOfSection && paragraph.ChildNodes.Count == 1 && paragraph.ChildNodes[0].GetText() == "\f")
				{
					var run = new Run(paragraph.Document);
					paragraph.AppendChild(run);
					var currentEndPageNum = this.pageNumberFinder.GetPageEnd(paragraph);
					this.pageNumberFinder.AddPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
				}

				foreach (Paragraph clonePara in SplitComposite(paragraph))
				{
					// Remove list numbering from the cloned paragraph but leave the indent the same 
					// as the paragraph is supposed to be part of the item before.
					if (paragraph.IsListItem)
					{
						var textPosition = clonePara.ListFormat.ListLevel.TextPosition;
						clonePara.ListFormat.RemoveNumbers();
						clonePara.ParagraphFormat.LeftIndent = textPosition;
					}
					// Reset spacing of split paragraphs in tables as additional spacing may cause them to look different.
					if (paragraph.IsInCell)
					{
						clonePara.ParagraphFormat.SpaceBefore = 0;
						paragraph.ParagraphFormat.SpaceAfter = 0;
					}
				}

				return VisitorAction.Continue;
			}

			public override VisitorAction VisitSectionEnd(Section section)
			{
				foreach (Section cloneSection in this.SplitComposite(section))
				{
					cloneSection.PageSetup.SectionStart = SectionStart.NewPage;
					cloneSection.PageSetup.RestartPageNumbering = true;
					cloneSection.PageSetup.PageStartingNumber = section.PageSetup.PageStartingNumber + (section.Document.IndexOf(cloneSection) - section.Document.IndexOf(section));
					cloneSection.PageSetup.DifferentFirstPageHeaderFooter = false;

					// corrects page break on end of the section
					SplitPageBreakCorrector.ProcessSection(cloneSection);
				}

				// corrects page break on end of the section
				SplitPageBreakCorrector.ProcessSection(section);

				// Add new page numbering for the body of the section as well.
				this.pageNumberFinder.AddPageNumbersForNode(section.Body, this.pageNumberFinder.GetPage(section), this.pageNumberFinder.GetPageEnd(section));
				return VisitorAction.Continue;
			}

			private VisitorAction ContinueIfCompositeAcrossPageElseSkip(CompositeNode composite)
			{
				return (this.pageNumberFinder.PageSpan(composite) > 1) ? VisitorAction.Continue : VisitorAction.SkipThisNode;
			}

			private List<Node> SplitComposite(CompositeNode composite)
			{
				List<Node> splitNodes = new List<Node>();
				foreach (Node splitNode in this.FindChildSplitPositions(composite))
				{
					splitNodes.Add(this.SplitCompositeAtNode(composite, splitNode));
				}

				return splitNodes;
			}

			private IEnumerable<Node> FindChildSplitPositions(CompositeNode node)
			{
				// A node may span across multiple pages so a list of split positions is returned.
				// The split node is the first node on the next page.
				var splitList = new List<Node>();
				int startingPage = this.pageNumberFinder.GetPage(node);
				Node[] childNodes = node.NodeType == NodeType.Section
										? ((Section)node).Body.ChildNodes.ToArray()
										: node.ChildNodes.ToArray();
				foreach (Node childNode in childNodes)
				{
					int pageNum = this.pageNumberFinder.GetPage(childNode);

					if (childNode is Run)
					{
						pageNum = this.pageNumberFinder.GetPageEnd(childNode);
					}

					// If the page of the child node has changed then this is the split position. Add
					// this to the list.
					if (pageNum > startingPage)
					{
						splitList.Add(childNode);
						startingPage = pageNum;
					}

					if (this.pageNumberFinder.PageSpan(childNode) > 1)
					{
						this.pageNumberFinder.AddPageNumbersForNode(childNode, pageNum, pageNum);
					}
				}

				// Split composites backward so the cloned nodes are inserted in the right order.
				splitList.Reverse();
				return splitList;
			}

			private CompositeNode SplitCompositeAtNode(CompositeNode baseNode, Node targetNode)
			{
				CompositeNode cloneNode = (CompositeNode)baseNode.Clone(false);
				Node node = targetNode;
				int currentPageNum = this.pageNumberFinder.GetPage(baseNode);

				// Move all nodes found on the next page into the copied node. Handle row nodes separately.
				if (baseNode.NodeType != NodeType.Row)
				{
					CompositeNode composite = cloneNode;
					if (baseNode.NodeType == NodeType.Section)
					{
						cloneNode = (CompositeNode)baseNode.Clone(true);
						Section section = (Section)cloneNode;
						section.Body.RemoveAllChildren();
						composite = section.Body;
					}

					while (node != null)
					{
						Node nextNode = node.NextSibling;
						composite.AppendChild(node);
						node = nextNode;
					}
				}
				else
				{
					// If we are dealing with a row then we need to add in dummy cells for the cloned row.
					int targetPageNum = this.pageNumberFinder.GetPage(targetNode);
					Node[] childNodes = baseNode.ChildNodes.ToArray();
					foreach (Node childNode in childNodes)
					{
						int pageNum = this.pageNumberFinder.GetPage(childNode);
						if (pageNum == targetPageNum)
						{
							if (cloneNode.NodeType == NodeType.Row)
								((Row)cloneNode).EnsureMinimum();

							if (cloneNode.NodeType == NodeType.Cell)
								((Cell)cloneNode).EnsureMinimum();

							cloneNode.LastChild.Remove();
							cloneNode.AppendChild(childNode);
						}
						else if (pageNum == currentPageNum)
						{
							cloneNode.AppendChild(childNode.Clone(false));
							if (cloneNode.LastChild.NodeType != NodeType.Cell)
							{
								((CompositeNode)cloneNode.LastChild).AppendChild(((CompositeNode)childNode).FirstChild.Clone(false));
							}
						}
					}
				}

				// Insert the split node after the original.
				baseNode.ParentNode.InsertAfter(cloneNode, baseNode);

				// Update the new page numbers of the base node and the clone node including its descendents.
				// This will only be a single page as the cloned composite is split to be on one page.
				int currentEndPageNum = this.pageNumberFinder.GetPageEnd(baseNode);
				this.pageNumberFinder.AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
				this.pageNumberFinder.AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);
				foreach (Node childNode in cloneNode.GetChildNodes(NodeType.Any, true))
				{
					this.pageNumberFinder.AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
				}

				return cloneNode;
			}
		}

		internal class SplitPageBreakCorrector
		{
			private const string PageBreakStr = "\f";
			private const char PageBreak = '\f';

			public static void ProcessSection(Section section)
			{
				if (section.ChildNodes.Count == 0)
				{
					return;
				}

				var lastBody = section.ChildNodes.OfType<Body>().LastOrDefault();
				if (lastBody == null)
				{
					return;
				}

				var run = lastBody.GetChildNodes(NodeType.Run, true).OfType<Run>().FirstOrDefault(p => p.Text.EndsWith(PageBreakStr));

				if (run != null)
				{
					RemovePageBreak(run);
				}

				return;
			}

			public static void RemovePageBreakFromParagraph(Paragraph paragraph)
			{
				Run run = (Run)paragraph.FirstChild;
				if (run.Text.Equals(PageBreakStr))
				{
					paragraph.RemoveChild(run);
				}
			}

			private static void ProcessLastParagraph(Paragraph paragraph)
			{
				Node lastNode = paragraph.ChildNodes[paragraph.ChildNodes.Count - 1];
				if (lastNode.NodeType != NodeType.Run)
				{
					return;
				}

				Run run = (Run)lastNode;
				RemovePageBreak(run);
			}

			private static void RemovePageBreak(Run run)
			{
				var paragraph = run.ParentParagraph;
				if (run.Text.Equals(PageBreakStr))
				{
					paragraph.RemoveChild(run);
				}
				else if (run.Text.EndsWith(PageBreakStr))
				{
					run.Text = run.Text.TrimEnd(PageBreak);
				}

				if (paragraph.ChildNodes.Count == 0)
				{
					CompositeNode parent = paragraph.ParentNode;
					parent.RemoveChild(paragraph);
				}
			}
		}
	}
}
