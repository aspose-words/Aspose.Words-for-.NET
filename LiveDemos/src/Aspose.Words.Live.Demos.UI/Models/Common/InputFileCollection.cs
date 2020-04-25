using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Aspose.Words.Live.Demos.UI.Models.Common
{
	public class InputFile
	{
		private string _fileName = string.Empty;
		private string _folderName = string.Empty;
		private string _fullFileName = string.Empty;
		


		public InputFile(string fileName, string folderName, string fullFileName)
		{
			_fileName = fileName;
			_folderName = folderName;
			_fullFileName = fullFileName;
		}
		
		public string FileName
		{
			get { return _fileName; }
			
		}
		public string FolderName
		{
			get { return _folderName; }
			
		}
		public string FullFileName
		{
			get { return _fullFileName; }

		}

	}

	public class InputFiles : List<InputFile>
	{

	}
}