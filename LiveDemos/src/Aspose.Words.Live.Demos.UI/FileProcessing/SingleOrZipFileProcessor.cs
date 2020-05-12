using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words.Live.Demos.UI.Config;
using System;
using System.IO;
using System.IO.Compression;

namespace Aspose.Words.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// SingleOrZipFileProcessor class to check either to create zip file or not
	///</Summary>
	public abstract class SingleOrZipFileProcessor : FileProcessor
    {
		///<Summary>
		/// ProcessFileToResponce method 
		///</Summary>
		/// <param name="resp"></param>
		/// <param name="inputFolderName"></param>
		/// <param name="inputFileName"></param>
		/// <param name="outputFileName"></param>
		
		protected override void ProcessFileToResponce(Response resp, string inputFolderName, string inputFileName, string outputFileName = null)
        {
            var inputFolderPath = Path.Combine(Configuration.WorkingDirectory, inputFolderName);
            var inputFilePath = Path.Combine(inputFolderPath, inputFileName);

            var outputFolderName = Guid.NewGuid().ToString();
            var outputFolderPath = Path.Combine(Configuration.OutputDirectory, outputFolderName);

            Directory.CreateDirectory(outputFolderPath);

            ProcessFileToPath(inputFilePath, outputFolderPath);

            var files = Directory.GetFiles(outputFolderPath);

            // Dont create Zip if there is only one file in output
            if (files.Length == 1)
            {
                outputFileName = outputFileName ?? Path.GetFileName(files[0]);

                resp.FileName = outputFileName;
                resp.FolderName = outputFolderName;
            }
            else
            {
                var outputZipFolderName = Guid.NewGuid().ToString();
                var outputZipFolderPath = Path.Combine(Configuration.OutputDirectory, outputZipFolderName);
                Directory.CreateDirectory(outputZipFolderPath);

                outputFileName = (outputFileName ?? Guid.NewGuid().ToString()) + ".zip";

                var outputZipFilePath = Path.Combine(outputZipFolderPath, outputFileName);

                ZipFile.CreateFromDirectory(outputFolderPath, outputZipFilePath);
                Directory.Delete(outputFolderPath, true);

                resp.FileName = outputFileName;
                resp.FolderName = outputZipFolderName;
            }
        }
		///<Summary>
		/// ProcessFileToPath method to process file to path
		///</Summary>
		/// <param name="inputFilePath"></param>
		/// <param name="outDirectoryPath"></param>
		protected abstract void ProcessFileToPath(string inputFilePath, string outDirectoryPath);
    }
}
