using Aspose.Words.Live.Demos.UI.Helpers;
using Aspose.Words.Live.Demos.UI.Config;
using Aspose.Words.Live.Demos.UI.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Services
{
	/// <summary>
	/// Path processor.
	/// Manage paths for source and processed files.
	/// Have ability to zip folder.
	/// </summary>
	public class PathProcessor
	{
		/// <summary>
		/// Source upload id.
		/// </summary>
		public string id { get; private set; }
		/// <summary>
		/// Source uploaded file name
		/// </summary>
		public string File { get; private set; }

		/// <summary>
		/// Download id zip.
		/// Filled after GetResultZipped call.
		/// </summary>
		public string idZip { get; private set; }
		/// <summary>
		/// Download zip file name.
		/// Filled after GetResultZipped call.
		/// </summary>
		public string FileZip { get; private set; }

		/// <summary>
		/// Source folder.
		/// This place is where source (uploaded) file stores.
		/// </summary>
		public string SourceFolder => PathSafe.GenerateAndValidate(Configuration.WorkingDirectory, id);
		/// <summary>
		/// Output folder.
		/// This place is where processed file stores.
		/// </summary>
		public string OutFolder => PathSafe.GenerateAndValidate(Configuration.OutputDirectory, id);

		/// <summary>
		/// Generates path to file inside source folder.
		/// </summary>
		/// <param name="file">File name.</param>
		/// <returns>Path.</returns>
		public string GetSourceFilePath(string file) => PathSafe.GenerateAndValidate(Configuration.WorkingDirectory, id, file);
		/// <summary>
		/// Generates path to file inside output folder.
		/// </summary>
		/// <param name="file">File name.</param>
		/// <returns>Path.</returns>
		public string GetOutFilePath(string file) => PathSafe.GenerateAndValidate(Configuration.OutputDirectory, id, file);

		/// <summary>
		/// Default path to source file.
		/// Combined from id and file.
		/// </summary>
		public string DefaultSourceFile => GetSourceFilePath(File);
		/// <summary>
		/// Default path to output file.
		/// Combined from id and file.
		/// </summary>
		public string DefaultOutFile => GetOutFilePath(File);

		/// <summary>
		/// Path to folder where zip file stored.
		/// Filled after GetResultZipped call.
		/// </summary>
		public string ZipFolder => idZip == null ? null : PathSafe.GenerateAndValidate(Configuration.OutputDirectory, idZip);
		/// <summary>
		/// Path to zip file.
		/// Filled after GetResultZipped call.
		/// </summary>
		public string ZipFile => idZip == null ? null : PathSafe.GenerateAndValidate(Configuration.OutputDirectory, idZip, FileZip);

		/// <summary>
		/// Constructor used to prepare processing paths.
		/// Creates OutFolder.
		/// </summary>
		/// <param name="id">Upload id.</param>
		/// <param name="file">File name.</param>
		/// <param name="checkDefaultSourceFileExistence">When true forces to check DefaultSourceFile existence.</param>
		public PathProcessor(string id, string file, bool checkDefaultSourceFileExistence)
		{
			this.id = id;
			this.File = file;

			if (checkDefaultSourceFileExistence && !System.IO.File.Exists(DefaultSourceFile))
				throw HttpHelper.Http404();

			Directory.CreateDirectory(OutFolder);
		}

		/// <summary>
		/// Constructor used to prepare uploading paths.
		/// Creates SourceFolder.
		/// </summary>
		/// <param name="id"></param>
		public PathProcessor(string id)
		{
			this.id = id;
			Directory.CreateDirectory(SourceFolder);
		}

		/// <summary>
		/// Constructor used to prepare downloading.
		/// Checks DefaultOutFile existence.
		/// </summary>
		/// <param name="id"></param>
		/// <param name="file"></param>
		public PathProcessor(string id, string file)
		{
			this.id = id;
			this.File = file;

			if (!System.IO.File.Exists(DefaultOutFile))
				throw HttpHelper.Http404();
		}

		/// <summary>
		/// Returns FileSafeResult object. pointed to output file.
		/// Returns default file name if not specified file param.
		/// </summary>
		/// <param name="file">File name.</param>
		/// <returns>FileSafeResult.</returns>		
		public FileSafeResult GetResult(string file = null)
		{
			return new FileSafeResult(
				file != null
					? GetOutFilePath(file)
					: DefaultOutFile
			)
			{
				id = id,
				FileName = file != null ? file : File
			};
		}

		/// <summary>
		/// Zips output folder and returns FileSafeResult object pointed to zip file.
		/// Cleanups OutFolder and SourceFolder.
		/// </summary>
		/// <returns>FileSafeResult.</returns>
		public FileSafeResult GetResultZipped()
		{
			idZip = $"{Guid.NewGuid()}";
			FileZip = $"{Path.GetFileNameWithoutExtension(File)}.zip";

			Directory.CreateDirectory(ZipFolder);
			System.IO.Compression.ZipFile.CreateFromDirectory(OutFolder, ZipFile);
			Directory.Delete(OutFolder, true);
			Directory.Delete(SourceFolder, true);
			var result = new FileSafeResult(ZipFile)
			{
				id = idZip,
				FileName = FileZip
			};

			return result;
		}
	}
}
