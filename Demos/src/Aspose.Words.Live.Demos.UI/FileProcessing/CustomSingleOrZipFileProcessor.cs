namespace Aspose.Words.Live.Demos.UI.FileProcessing
{
	///<Summary>
	/// CustomSingleOrZipFileProcessor class 
	///</Summary>
	public class CustomSingleOrZipFileProcessor : SingleOrZipFileProcessor
    {
		///<Summary>
		/// ProcessFileDelegate delegate
		///</Summary>
		public delegate void ProcessFileDelegate(string inputFilePath, string outputFolderPath);

		///<Summary>
		/// CustomProcessMethod method
		///</Summary>
		public ProcessFileDelegate CustomProcessMethod { get; set; }

		///<Summary>
		/// ProcessFileToPath method
		///</Summary>
		protected override void ProcessFileToPath(string inputFilePath, string outputFolderPath)
        {
            var action = CustomProcessMethod;

            action?.Invoke(inputFilePath, outputFolderPath);
        }
    }
}
