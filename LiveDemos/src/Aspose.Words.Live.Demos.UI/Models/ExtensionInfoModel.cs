using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Aspose.Words.Live.Demos.UI.Controllers;

namespace Aspose.Words.Live.Demos.UI.Models
{
  /// <summary>
  /// Prepares Extension and HowTo sections
  /// Changes Title and TitleSub in ViewModel
  /// </summary>
  public class ExtensionInfoModel
  {
    public ViewModel Parent;

    /// <summary>
    /// File extension without dot received by "fileformat" value in RouteData (e.g. docx)
    /// </summary>
    public string Extension { get; set; }
    public string Name { get; set; }
    public string Description { get; set; }
    public string URL { get; set; }

    public string AppName => Parent.AppName;

    //public GeneratedPage GeneratedPage => Parent.GeneratedPage;

    public ExtensionInfoModel(ViewModel parent, string extension)
    {
      Parent = parent;
      Extension = extension;

      // For Aspose.Words it is SEO bad
      //if (Parent.Product != "words" && Parent.Product != "pdf" && GeneratedPage != null)
      //{
      //  parent.Title = string.Format(GeneratedPage.MainHeadline, Extension.ToUpper());
      //  parent.TitleSub = string.Format(GeneratedPage.SubHeadline, Extension.ToUpper());
      //}

      
      Name = "";
      Description = "";
      URL = "";
    }
  }
}
