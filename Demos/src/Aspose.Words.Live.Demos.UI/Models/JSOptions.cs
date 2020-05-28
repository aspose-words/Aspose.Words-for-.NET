using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Words.Live.Demos.UI.Config;
using Newtonsoft.Json;

namespace Aspose.Words.Live.Demos.UI.Models
{
  public class JSOptions
  {
    private readonly ViewModel Parent;
		private Dictionary<string, string> Resources => Parent.Resources;

    public string AppURL => Parent.AppURL;
    public string AppName => Parent.AppName;

    //public string APIBasePath => $"{Configuration.AsposeToolsAPIBasePath}";
    public string UIBasePath => $"{Configuration.AsposeAppLiveDemosPath}";

		public string ViewerPathWF => $"{UIBasePath}/{Parent.Product}/viewer/";

    public string ViewerPath => $"{UIBasePath}/{Parent.Product}/view?";
    public string EditorPath => $"{UIBasePath}/{Parent.Product}/edit?";

    public string FileSelectMessage => Resources["FileSelectMessage"];

		public string Product => Parent.Product;

	public int MaximumUploadFiles => Parent.MaximumUploadFiles;
    
    public string FileAmountMessage => Resources["FileAmountMessageOne"];

    /// <summary>
    /// Apps like Viewer and Editor
    /// </summary>
    public bool UploadAndRedirect => Parent.UploadAndRedirect;
    public bool UseSorting => Parent.UseSorting;
		public bool ShowViewerButton => Parent.ShowViewerButton;
		public string FileWrongTypeMessage { get; }

    public Dictionary<int, string> FileProcessingErrorCodes => new Dictionary<int, string>()
    {
      { (int)FileProcessingErrorCode.NoSearchResults, Resources["NoSearchResultsMessage"] },
      { (int)FileProcessingErrorCode.WrongRegExp, Resources["WrongRegExpMessage"] }
    };

    /// <summary>
    /// ['DOCX', 'DOC', ...]
    /// </summary>
    public IEnumerable<string> UploadOptions =>
      Parent.ExtensionsString.Replace(".", "").ToUpper().Split('|');

    #region FileDrop
    public bool Multiple => !UploadAndRedirect;
    public string DropFilesPrompt => Parent.DropOrUploadFileLabel;
    public string Accept => Parent.ExtensionsString.Replace("|.", ",.");
    #endregion

    public JSOptions(ViewModel model)
    {
      Parent = model;
      if (string.IsNullOrEmpty(model.Extension) || model.IsCanonical)
        FileWrongTypeMessage = Resources["FileWrongTypeMessage"];
      else
        FileWrongTypeMessage = string.Format(Resources["FileWrongTypeMessage2"], $"<a href=\"/{Parent.Product}/{AppName.ToLower()}\">{AppName}</a>");
    }

    public override string ToString()
    {
      return JsonConvert.SerializeObject(this, Formatting.None);
    }
  }
}
