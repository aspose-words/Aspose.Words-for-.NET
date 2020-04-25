using Aspose.Words.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Runtime.Remoting.Channels;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using Aspose.Words.Live.Demos.UI.Helpers;
using Aspose.Words.Live.Demos.UI.Services;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Newtonsoft.Json.Linq;
using Aspose.Words.Live.Demos.UI.Config;
using System.Web.Http;
using Aspose.Words.Live.Demos.UI.Models.Common;

namespace Aspose.Words.Live.Demos.UI.Models
{
	///<Summary>
	/// AsposeWordsBase class 
	///</Summary>
	
  public class AsposeWordsBase : ModelBase
	{
		
		/// <summary>
		/// Maximum number of files which can be uploaded for MVC Aspose.Words apps
		/// </summary>
		protected const int MaximumUploadFiles = 10;

    /// <summary>
    /// Original file format SaveAs option for multiple files uploading. By default, "-"
    /// </summary>
    protected const string SaveAsOriginalName = ".-";
    
    /// <summary>
    /// Response when uploaded files exceed the limits
    /// </summary>
    protected Response MaximumFileLimitsResponse = new Response()
    {
      Status = $"Number of files should be less {MaximumUploadFiles}",
      StatusCode = 500
    };
		/// <summary>
		/// Response when uploaded files exceed the limits
		/// </summary>
		protected Response PasswordProtectedResponse = new Response()
		{
			Status = "Some of your documents are password protected",
			StatusCode = 500
		};
		/// <summary>
		/// Response when uploaded files exceed the limits
		/// </summary>
		protected Response BadDocumentResponse = new Response()
		{
			Status = "Some of your documents are corrupted",
			StatusCode = 500
		};


		///<Summary>
		/// Options class 
		///</Summary>
		public class Options
    {
      
      ///<Summary>
      /// FolderName
      ///</Summary>
      public string FolderName;
      ///<Summary>
      /// FileName
      ///</Summary>
      public string FileName;

      private string _outputType;

      /// <summary>
      /// By default, it is the extension of FileName - e.g. ".docx"
      /// </summary>
      public string OutputType
      {
        get => _outputType;
        set
        {
          if (!value.StartsWith("."))
            value = "." + value;
          _outputType = value.ToLower();
        }
      }

      /// <summary>
      /// Check if OuputType is a picture extension
      /// </summary>
      public bool IsPicture
      {
        get
        {
          switch (_outputType)
          {
            case ".bmp":
            case ".emf":
            case ".png":
            case ".jpg":
            case ".jpeg":
            case ".gif":
            case ".tiff":
              return true;
            default:
              return false;
          }
        }
      }

      ///<Summary>
      /// ResultFileName
      ///</Summary>
      public string ResultFileName = "";

  
			///<Summary>
			/// ModelName
			///</Summary>
			public string ModelName;
      ///<Summary>
      /// CreateZip
      ///</Summary>
      public bool CreateZip = false;
      ///<Summary>
      /// CheckNumberOfPages
      ///</Summary>
      public bool CheckNumberOfPages = false;
      ///<Summary>
      /// DeleteSourceFolder
      ///</Summary>
      public bool DeleteSourceFolder = false;

      /// <summary>
      /// Output zip filename (without '.zip'), if CreateZip property is true
      /// By default, FileName + AppName
      /// </summary>
      public string ZipFileName;

			///<Summary>
			/// AppName
			///</Summary>
			public string AppName;
			///<Summary>
			/// MethodName
			///</Summary>
			public string MethodName;

			/// <summary>
			/// AppSettings.WorkingDirectory + FolderName + "/" + FileName
			/// </summary>
			public string WorkingFileName
      {
        get
        {
          if (System.IO.File.Exists(Config.Configuration.WorkingDirectory + FolderName + "/" + FileName))
            return Config.Configuration.WorkingDirectory + FolderName + "/" + FileName;
          return Config.Configuration.OutputDirectory + FolderName + "/" + FileName;
        }
      }
    }
    /// <summary>
    /// init Options
    /// </summary>
    protected Options Opts = new Options();
    
    /// <summary>
    /// UTF8WithoutBom
    /// </summary>
    protected static readonly Encoding UTF8WithoutBom = new UTF8Encoding(false);

		/// <summary>
		/// Prepare upload files and return FileData
		/// </summary>
		protected async Task<Collection<MultipartFileData>> UploadFiles()
		{
			Opts.FolderName = Guid.NewGuid().ToString();
			var pathProcessor = new PathProcessor(Opts.FolderName);
			var uploadProvider = new MultipartFormDataStreamProviderSafe(pathProcessor.SourceFolder);
			await Request.Content.ReadAsMultipartAsync(uploadProvider);
			return uploadProvider.FileData;
		}

		/// <summary>
		/// AsposeWordsBase
		/// </summary>
		public AsposeWordsBase()
    {
    
      Opts.ModelName = GetType().Name;
    }

		/// <summary>
		/// AsposeWordsBase
		/// </summary>
		static AsposeWordsBase()
    {
			Aspose.Words.Live.Demos.UI.Models.License.SetAsposeWordsLicense();
      
    }

		/// <summary>
		/// Set default parameters into Opts
		/// </summary>
		/// <param name="docs"></param>
		protected void SetDefaultOptions(Document[] docs, string outputType)
		{
			if (docs.Length > 0)
			{
				SetDefaultOptions(Path.GetFileName(docs[0].OriginalFileName), outputType);
				Opts.CreateZip = docs.Length > 1 || Opts.IsPicture;
			}
		}

		/// <summary>
		/// Set default parameters into Opts
		/// </summary>
		/// <param name="InputFiles"></param>
		protected void SetDefaultOptions(InputFiles docs, string outputType)
		{
			if (docs.Count > 0)
			{
				SetDefaultOptions(docs[0].FileName, outputType);
				Opts.CreateZip = docs.Count > 1 || Opts.IsPicture;
			}
		}

		/// <summary>
		/// Set default parameters into Opts
		/// </summary>
		/// <param name="filename"></param>
		private void SetDefaultOptions(string filename, string outputType)
    {
			//Opts.FolderName = FolderName;
      Opts.ResultFileName = filename;
      Opts.FileName = Path.GetFileName(filename);

      //var query = Request.GetQueryNameValuePairs().ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
      
      //if (query.ContainsKey("outputType"))
        //outputType = query["outputType"];
      Opts.OutputType = !string.IsNullOrEmpty(outputType)
        ? outputType
        : Path.GetExtension(Opts.FileName);

      Opts.ResultFileName = Opts.OutputType == SaveAsOriginalName
        ? Opts.FileName
        : Path.GetFileNameWithoutExtension(Opts.FileName) + Opts.OutputType;
    }

		/// <summary>
		/// Check if the OutputType is a picture and saves the document
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="outPath"></param>
		/// <param name="zipOutFolder"></param>
		protected void SaveDocument(Document doc, string outPath, string zipOutFolder)
		{
			string filename;
			if (Opts.CreateZip)
				filename = zipOutFolder + "/" +
						   (Opts.OutputType == SaveAsOriginalName
							 ? Path.GetFileName(doc.OriginalFileName)
							 : Path.GetFileNameWithoutExtension(doc.OriginalFileName) + Opts.OutputType);
			else
				filename = outPath;
			SaveDocument(doc, filename);
		}

		/// <summary>
		/// Check if the OutputType is a picture and saves the document
		/// </summary>
		/// <param name="doc"></param>
		/// <param name="filename">Full FileName</param>
		protected void SaveDocument(Document doc, string filename)
		{
			if (!Opts.IsPicture)
			{
				switch (Opts.OutputType)
				{
					case ".pdf": // remove comments
						var comments = doc.GetChildNodes(NodeType.Comment, true);
						comments.Clear();
						break;
				}
				doc.Save(filename);
			}
			else
			{
				doc.UpdatePageLayout();
				var template = $"{Path.GetDirectoryName(filename)}\\{Path.GetFileNameWithoutExtension(filename)} ";
				var so = new ImageSaveOptions(FileFormatUtil.ExtensionToSaveFormat(Opts.OutputType))
				{
					PageCount = 1
				};

				for (var i = 1; i <= doc.PageCount; i++)
				{
					so.PageIndex = i - 1;
					var fname = template + i.ToString("D2") + Opts.OutputType;
					doc.Save(fname, so);
				}
			}
		}

		/// <summary>
		/// Process
		/// </summary>
		protected Response Process(ActionDelegate action)
    {
      if (string.IsNullOrEmpty(Opts.OutputType) && !string.IsNullOrEmpty(Opts.FileName))
        Opts.OutputType = Path.GetExtension(Opts.FileName);

      if (Opts.OutputType == ".html" || Opts.IsPicture)
        Opts.CreateZip = true;

      if (string.IsNullOrEmpty(Opts.ZipFileName))
        Opts.ZipFileName = Path.GetFileNameWithoutExtension(Opts.FileName) + Opts.AppName;
      
      var outputType = Opts.OutputType;
      if (outputType == SaveAsOriginalName && !string.IsNullOrEmpty(Opts.FileName))
        outputType = Path.GetExtension(Opts.FileName);

      return Process(Opts.ModelName, Opts.ResultFileName, Opts.FolderName, outputType, Opts.CreateZip, Opts.CheckNumberOfPages,
         Opts.MethodName, action, Opts.DeleteSourceFolder, Opts.ZipFileName);
    }

    

    

   

    

    #region Common
    /// <summary>
    /// IsValidRegex
    /// </summary>
    public static bool IsValidRegex(string pattern)
    {
      if (string.IsNullOrEmpty(pattern))
        return false;
      try
      {
        Regex.Match("", pattern);
      }
      catch (ArgumentException)
      {
        return false;
      }
      return true;
    }

    /// <summary>
    /// Prepare output folder for using when multiple files are uploaded
    /// Creates folder by filename without extension
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="path">Zip folder name</param>
    /// <returns>Tuple(original filename, output folder)</returns>
    protected static (string, string) PrepareFolder(Document doc, string path)
    {
      var filename = Path.GetFileNameWithoutExtension(doc.OriginalFileName);
      var folder = path + "/";
      folder += filename;
      while (Directory.Exists(folder))
        folder += "_";
      folder += "/";
      Directory.CreateDirectory(folder);
      return (Path.GetFileName(doc.OriginalFileName), folder);
    }

    ///<Summary>
    /// Extract images
    ///</Summary>
    protected void ExtractImages(Document doc, string outPath)
    {
			try
			{
				var shapes = doc.GetChildNodes(NodeType.Shape, true);
				var imageIndex = 0;
				foreach (var shape in shapes.OfType<Shape>().Where(x => x.HasImage))
					try
					{
						var imageFileName = $"Image_{imageIndex:00}{FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType)}";
						shape.ImageData.Save(outPath + "/" + imageFileName);
						imageIndex++;
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message);
					}
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
    }

    /// <summary>
    /// ResizeImage
    /// </summary>
    public static Bitmap ResizeImage(Image image, double zoom)
    {
      var width = (int)(image.Width * zoom);
      var height = (int)(image.Height * zoom);

      var destRect = new Rectangle(0, 0, width, height);
      var destImage = new Bitmap(width, height);

      destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

      using (var graphics = Graphics.FromImage(destImage))
      {
        graphics.CompositingMode = CompositingMode.SourceCopy;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.SmoothingMode = SmoothingMode.HighQuality;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

        using (var wrapMode = new ImageAttributes())
        {
          wrapMode.SetWrapMode(WrapMode.TileFlipXY);
          graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
        }
      }

      return destImage;
    }

    #endregion
  }
}
