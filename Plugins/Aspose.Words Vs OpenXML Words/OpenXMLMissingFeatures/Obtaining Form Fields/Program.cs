using Aspose.Words;
using Aspose.Words.Fields;
/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Words for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string FileName = FilePath + "Obtaining Form Fields.docx";
            
            //Shows how to get a collection of form fields.
            Document doc = new Document(FileName);
            FormFieldCollection formFields = doc.Range.FormFields;


            //Shows how to access form fields.
            Document myDoc = new Document(FileName);
            FormFieldCollection documentFormFields = myDoc.Range.FormFields;

            FormField formField1 = documentFormFields[3];
            FormField formField2 = documentFormFields["CustomerName"];

        }
    }
}
