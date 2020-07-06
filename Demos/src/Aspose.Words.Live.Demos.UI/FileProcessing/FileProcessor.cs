using  Aspose.Words.Live.Demos.UI.Models;

namespace Aspose.Words.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// FileProcessor class to process file
	///</Summary>
	public abstract class FileProcessor
    {
		///<Summary>
		/// FileProcessor class process method
		///</Summary>
		///<param name="inputFolderName"></param>
		///<param name="inputFileName"></param>
		///<param name="outputFileName"></param>
		public virtual Response Process(string inputFolderName, string inputFileName, string outputFileName = null)
        {
            var resp = new Response()
            {
                StatusCode = 200,
                Status = "OK"
            };

            ProcessFileToResponce(resp, inputFolderName, inputFileName, outputFileName);

            return resp;
        }
		///<Summary>
		/// FileProcessor class ProcessFileToResponce method
		///</Summary>
		///<param name="inputFolderName"></param>
		///<param name="inputFileName"></param>
		///<param name="outputFileName"></param>
		///<param name="resp"></param>
		protected abstract void ProcessFileToResponce(Response resp, string inputFolderName, string inputFileName, string outputFileName = null);
    }
}
