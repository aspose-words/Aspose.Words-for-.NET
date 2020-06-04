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
using Aspose.Words.Live.Demos.UI.Models.Common;
using Aspose.Words.Live.Demos.UI.Controllers;
using Newtonsoft.Json;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Replacing;
using Aspose.Words.Saving;
using Document = Aspose.Words.Document;
using File = System.IO.File;
using SaveFormat = Aspose.Words.SaveFormat;
using System.Net;
using System.Text.RegularExpressions;
using System.Drawing;
using System.Drawing.Imaging;


namespace Aspose.Words.Live.Demos.UI.Controllers
{
	///<Summary>
	/// AsposeWordsViewerController class to merge word document
	///</Summary>
	public class AsposeWordsViewerController : AsposeWordsBase
	{


		private const double ThumbnailsWidth = 150; // in pixels
		[HttpPost]		
		public HttpResponseMessage DocumentInfo(RequestData request)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = request.fileName;
			Opts.FolderName = request.folderName;
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				var doc = new Document(Opts.WorkingFileName);
				var lc = new LayoutCollector(doc);
				PrepareInternalLinks(doc, lc);

				var lst = new List<PageView>(doc.PageCount);
				for (int i = 0; i < doc.PageCount; i++)
				{
					var pageInfo = doc.GetPageInfo(i);
					var size = pageInfo.GetSizeInPixels(1, 72);
					lst.Add(new PageView()
					{
						width = size.Width,
						height = size.Height,
						angle = 0,
						number = i + 1
					});
				}
				
				return Request.CreateResponse(HttpStatusCode.OK, new PageParametersResponse(request.fileName, lst, PrepareNavigationPaneList(doc, lc)));
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}

		/// <summary>
		/// Change hyperlinks to _page{pageNumber}_{...}
		/// </summary>
		/// <param name="doc"></param>
		private void PrepareInternalLinks(Document doc, LayoutCollector lc)
		{
			foreach (var field in doc.Range.Fields)
				try
				{
					if (field.Type == FieldType.FieldHyperlink)
					{
						var link = (FieldHyperlink)field;
						if (!string.IsNullOrEmpty(link.SubAddress))
						{
							var bookmark = doc.Range.Bookmarks[link.SubAddress];
							if (bookmark == null)
								continue;
							var pageNumber = lc.GetStartPageIndex(bookmark.BookmarkStart);
							if (pageNumber == 0)
								continue;
							var name = $"_page{pageNumber}{link.SubAddress}";
							bookmark.Name = name;
							link.SubAddress = name;
						}
					}
					else if (field.Type == FieldType.FieldRef)
					{
						var link = (FieldRef)field;
						var bookmark = doc.Range.Bookmarks[link.BookmarkName];
						if (bookmark == null)
							continue;
						var pageNumber = lc.GetStartPageIndex(bookmark.BookmarkStart);
						if (pageNumber == 0)
							continue;
						var name = $"_page{pageNumber}{link.BookmarkName}";
						bookmark.Name = name;
						link.BookmarkName = name;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine(ex.Message);
				}
			var filename =  Config.Configuration.OutputDirectory + Opts.FolderName + "/" +
						   Path.GetFileNameWithoutExtension(Opts.FileName) + "_links.docx";
			doc.Save(filename);
		}

		private List<NavigationPaneItem> PrepareNavigationPaneList(Document doc, LayoutCollector lc)
		{
			try
			{
				var lst = new List<NavigationPaneItem>();
				foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
					switch (para.ParagraphFormat.StyleIdentifier)
					{
						case StyleIdentifier.Subtitle:
						case StyleIdentifier.Title:
						case StyleIdentifier.Heading3:
						case StyleIdentifier.Heading2:
						case StyleIdentifier.Heading1:
							var text = para.Range.Text;
							if (text.Count(char.IsLetterOrDigit) > 0)
								lst.Add(new NavigationPaneItem(text, para.ParagraphFormat.StyleIdentifier, lc.GetStartPageIndex(para)));
							break;
					}
				return lst;
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
				return null;
			}
		}

		/// <summary>
		/// Get page of the document with modified hyperlinks
		/// </summary>
		/// <param name="request"></param>
		/// <returns></returns>
		[HttpPost]		
		public HttpResponseMessage Page(RequestData request)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = request.fileName;
			Opts.FolderName = request.folderName;
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

			var filename = Config.Configuration.OutputDirectory + Opts.FolderName + "/" +
						   Path.GetFileNameWithoutExtension(Opts.FileName) + "_links.docx"; // This file was prepared by PrepareInternalLinks method
			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				var doc = new Document(filename);
				var page = PreparePageView(doc, request.pageNumber);
				return Request.CreateResponse(HttpStatusCode.OK, page);
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}
		/// <summary>
		/// Common method for 'Page' and 'Print'
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="pageNumber">Starts with one</param>
		/// <returns></returns>
		private PageView PreparePageView(Document doc, int pageNumber)
		{
			var so = new HtmlFixedSaveOptions()
			{
				ExportEmbeddedFonts = true,
				ExportEmbeddedImages = true,
				ExportEmbeddedCss = true,
				CssClassNamesPrefix = "pg" + pageNumber,
				ShowPageBorder = false,
				Encoding = UTF8WithoutBom,
				PageCount = 1,
				PageIndex = pageNumber - 1
			};

			using (var stream = new MemoryStream())
			{
				doc.Save(stream, so);
				var pageInfo = doc.GetPageInfo(pageNumber - 1);
				var size = pageInfo.GetSizeInPixels(1, 72);
				return new PageView()
				{
					width = size.Width,
					height = size.Height,
					angle = 0,
					number = pageNumber,
					data = UTF8WithoutBom.GetString(stream.ToArray()).Replace("letter-spacing:-107374182.4pt;", "")
				};
			}
		}

		///<Summary>
		/// Thumbnails method to get Thumbnails
		///</Summary>
		[HttpPost]

		public HttpResponseMessage Search(RequestData request)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = request.fileName;
			Opts.FolderName = request.folderName;
			Opts.MethodName = "Search";

			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				if (string.IsNullOrEmpty(request.searchQuery))
					return Request.CreateResponse(HttpStatusCode.OK, new int[] { });

				var doc = new Document(Opts.WorkingFileName);
				var lst = new HashSet<int>();
				var findings = new AsposeWordsSearch.FindCallback();
				var options = new FindReplaceOptions()
				{
					ReplacingCallback = findings,
					Direction = FindReplaceDirection.Forward,
					MatchCase = false
				};
				doc.Range.Replace(new Regex(request.searchQuery, RegexOptions.IgnoreCase), "", options);
				var lc = new LayoutCollector(doc);
				foreach (var mathchedNode in findings.MatchedNodes)
					foreach (var node in mathchedNode.Value.Select(x => x.MatchNode))
					{
						var pageNumber = lc.GetStartPageIndex(node);
						lst.Add(pageNumber);
					}
				return Request.CreateResponse(HttpStatusCode.OK, lst);
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}

		///<Summary>
		/// Thumbnails method to get Thumbnails
		///</Summary>
		[HttpPost]		
		public HttpResponseMessage Thumbnails(RequestData request)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = request.fileName;
			Opts.FolderName = request.folderName;
			Opts.MethodName = "Thumbnails";

			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				var doc = new Document(Opts.WorkingFileName);
				var so = new ImageSaveOptions(SaveFormat.Png) { PageCount = 1 };
				var lst = new List<PageView>();
				for (int i = 0; i < doc.PageCount; i++)
				{
					so.PageIndex = i;
					var stream = new MemoryStream();
					doc.Save(stream, so);
					var pageImage = Image.FromStream(stream);
					var zoom = ThumbnailsWidth / pageImage.Width;
					var image = (Image)ResizeImage(pageImage, zoom);
					var resizedStream = new MemoryStream();
					image.Save(resizedStream, ImageFormat.Png);
					var page = new PageView()
					{
						width = image.Width,
						height = image.Height,
						angle = 0,
						number = i + 1,
						data = Convert.ToBase64String(resizedStream.ToArray())
					};
					lst.Add(page);
				}
				return Request.CreateResponse(HttpStatusCode.OK, lst);
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}
		///<Summary>
		/// Download method to download document
		///</Summary>
		[HttpPost]
		[AcceptVerbs("GET", "POST")]
		public HttpResponseMessage Download(string fileName, string folderName, string outputType)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = fileName;
			Opts.FolderName = folderName;
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				if (string.IsNullOrEmpty(outputType))
					outputType = Path.GetExtension(Opts.FileName);
				var fn = Path.GetFileNameWithoutExtension(Opts.FileName) + outputType;
				var resultfile = Config.Configuration.OutputDirectory + Opts.FolderName + "/" + fn;
				if (!File.Exists(resultfile))
					if (string.IsNullOrEmpty(outputType) ||
						Path.GetExtension(Opts.FileName).ToLower() == Path.GetExtension(outputType).ToLower())
					{
						Directory.CreateDirectory(Path.GetDirectoryName(resultfile));
						File.Copy(Opts.WorkingFileName, resultfile);
					}
					else
					{
						var doc = new Document(Opts.WorkingFileName);
						switch (outputType.ToLower())
						{
							case ".html":
								var so = new HtmlFixedSaveOptions()
								{
									ExportEmbeddedFonts = true,
									ExportEmbeddedImages = true,
									ExportEmbeddedCss = true,
									ShowPageBorder = false,
									Encoding = UTF8WithoutBom
								};
								doc.Save(resultfile, so);
								break;
							default:
								doc.Save(resultfile);
								break;
						}
					}

				var response = new Response()
				{
					FileName = HttpUtility.UrlEncode(fn),
					FolderName = Opts.FolderName,
					StatusCode = 200
				};
				return Request.CreateResponse(HttpStatusCode.OK, response);
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}
		///<Summary>
		/// Download method to download document
		///</Summary>
		[HttpPost]
		[AcceptVerbs("GET", "POST")]		
		public HttpResponseMessage Download(RequestData request)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = request.fileName;
			Opts.FolderName = request.folderName;
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				if (string.IsNullOrEmpty(request.outputType))
					request.outputType = Path.GetExtension(Opts.FileName);
				var fn = Path.GetFileNameWithoutExtension(Opts.FileName) + request.outputType;
				var resultfile = Config.Configuration.OutputDirectory + Opts.FolderName + "/" + fn;
				if (!File.Exists(resultfile))
					if (string.IsNullOrEmpty(request.outputType) ||
						Path.GetExtension(Opts.FileName).ToLower() == Path.GetExtension(request.outputType).ToLower())
					{
						Directory.CreateDirectory(Path.GetDirectoryName(resultfile));
						File.Copy(Opts.WorkingFileName, resultfile);
					}
					else
					{
						var doc = new Document(Opts.WorkingFileName);
						switch (request.outputType.ToLower())
						{
							case ".html":
								var so = new HtmlFixedSaveOptions()
								{
									ExportEmbeddedFonts = true,
									ExportEmbeddedImages = true,
									ExportEmbeddedCss = true,
									ShowPageBorder = false,
									Encoding = UTF8WithoutBom
								};
								doc.Save(resultfile, so);
								break;
							default:
								doc.Save(resultfile);
								break;
						}
					}

				var response = new Response()
				{
					FileName = HttpUtility.UrlEncode(fn),
					FolderName = Opts.FolderName,
					StatusCode = 200
				};
				return Request.CreateResponse(HttpStatusCode.OK, response);
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}
		///<Summary>
		/// Print method to print
		///</Summary>
		[HttpPost]		
		public HttpResponseMessage Print(RequestData request)
		{
			Opts.AppName = "Viewer";
			Opts.FileName = request.fileName;
			Opts.FolderName = request.folderName;
			Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;

			try
			{
				if (Opts.FolderName.Contains(".."))
					throw new Exception("Break-in attempt");

				var doc = new Document(Opts.WorkingFileName);
				var lst = new PageView[doc.PageCount];
				for (int i = 0; i < doc.PageCount; i++)
					lst[i] = PreparePageView(doc, i + 1);
				return Request.CreateResponse(HttpStatusCode.OK, lst);
			}
			catch (Exception ex)
			{
				return ExceptionResponse(ex);
			}
		}
		///<Summary>
		/// DocumentInfoCORS method to 
		///</Summary>
		[HttpOptions]		
		public HttpResponseMessage DocumentInfoCORS()
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
		///<Summary>
		/// PageCORS method to 
		///</Summary>
		[HttpOptions]		
		public HttpResponseMessage PageCORS()
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
		///<Summary>
		/// ThumbnailsCORS method to 
		///</Summary>
		[HttpOptions]		
		public HttpResponseMessage ThumbnailsCORS()
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
		///<Summary>
		/// DownloadCORS method to 
		///</Summary>
		[HttpOptions]		
		public HttpResponseMessage DownloadCORS()
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
		///<Summary>
		/// PrintCORS method to 
		///</Summary>
		[HttpOptions]		
		public HttpResponseMessage PrintCORS()
		{
			return Request.CreateResponse(HttpStatusCode.OK);
		}
		///<Summary>
		/// ExceptionResponse method to 
		///</Summary>
		public HttpResponseMessage ExceptionResponse(Exception ex)
		{
			Console.Write(ex.Message);
			return Request.CreateResponse(HttpStatusCode.InternalServerError, new ExceptionEntity(ex));
		}
		///<Summary>
		/// PageView class to get or set page properties
		///</Summary>
		public class PageView
		{
			///<Summary>
			/// get or set width
			///</Summary>
			public double width { get; set; }
			///<Summary>
			/// get or set height
			///</Summary>
			public double height { get; set; }
			///<Summary>
			/// get or set number
			///</Summary>
			public int number { get; set; }
			///<Summary>
			/// get or set angle
			///</Summary>
			public int angle { get; set; }
			///<Summary>
			/// get or set data
			///</Summary>
			public string data { get; set; }
		}
		
		///<Summary>
		/// PageParametersResponse class to get or set PageParametersResponse properties
		///</Summary>
		public class PageParametersResponse
		{
			///<Summary>
			/// get or set guid
			///</Summary>
			[JsonProperty]
			public string guid;

			///<Summary>
			/// get or set pages
			///</Summary>
			[JsonProperty]
			public List<PageView> pages;

			///<Summary>
			/// get or set printAllowed
			///</Summary>
			[JsonProperty]
			public bool printAllowed = true;

			///<Summary>
			/// List of items for Navigation Pane
			///</Summary>
			[JsonProperty]
			public List<NavigationPaneItem> navigationPane;

			///<Summary>
			/// PageParametersResponse
			///</Summary>
			public PageParametersResponse(string guid, List<PageView> pages, List<NavigationPaneItem> navigationPane)
			{
				this.guid = guid;
				this.pages = pages;
				this.navigationPane = navigationPane;
			}
		}
		///<Summary>
		/// ExceptionEntity class to get or set ExceptionEntity properties
		///</Summary>
		public class ExceptionEntity
		{
			///<Summary>
			/// get or set message
			///</Summary>
			public string message { get; set; }
			///<Summary>
			/// init ExceptionEntity
			///</Summary>
			public ExceptionEntity(Exception ex)
			{
				message = ex.Message;
			}
		}

		/// <summary>
		/// NavigationPaneItem
		/// </summary>
		public class NavigationPaneItem
		{
			/// <summary>
			/// Name of the heading
			/// </summary>
			public string name { get; set; }

			/// <summary>
			/// Style: Heading1, Heading2, etc.
			/// </summary>
			public StyleIdentifier style { get; set; }

			/// <summary>
			/// Page, on which the heading is
			/// </summary>
			public int pageNumber { get; set; }

			/// <summary>
			/// Constructor
			/// </summary>
			/// <param name="name"></param>
			/// <param name="style"></param>
			/// <param name="pageNumber"></param>
			public NavigationPaneItem(string name, StyleIdentifier style, int pageNumber)
			{
				this.name = name;
				this.style = style;
				this.pageNumber = pageNumber;
			}

			/// <summary>
			/// ToString
			/// </summary>
			/// <returns></returns>
			public override string ToString()
			{
				return $"{name} | {style} | {pageNumber}";
			}
		}


	}
}