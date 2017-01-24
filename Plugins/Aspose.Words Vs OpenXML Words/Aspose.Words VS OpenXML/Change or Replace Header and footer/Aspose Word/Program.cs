// Copyright (c) Aspose 2002-2014. All Rights Reserved.

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

            string path = FilePath + "Change or Replace Header and footer - Aspose.docx";
            ChangeHeader(path);

        }
        public static void ChangeHeader(string documentPath)
        {
            Document doc = new Document(documentPath);
            DocumentBuilder builder = new DocumentBuilder(doc);

            // --- Create header ---
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

            // Specify header title for the first page.
            builder.Write("Aspose.Words Header");

            // --- Create footer for pages other than first. ---
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // Specify Footer text.
            builder.Write("Aspose.Words Footer");

            // Save the resulting document.
            doc.Save(documentPath);
        }
    }
}
