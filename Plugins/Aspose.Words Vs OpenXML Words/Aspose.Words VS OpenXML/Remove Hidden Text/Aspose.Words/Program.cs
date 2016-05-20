using Aspose.Words;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "Remove Hidden Text.docx";

            Document doc = new Document(fileName);
            foreach (Paragraph par in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                par.ParagraphBreakFont.Hidden = false;
                foreach (Run run in par.GetChildNodes(NodeType.Run, true))
                {
                    if (run.Font.Hidden)
                        run.Font.Hidden = false;
                }
            }
            doc.Save(fileName);
        }
    }
}
