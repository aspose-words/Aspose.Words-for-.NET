// Copyright (c) Aspose 2002-2021. All Rights Reserved.

/*
    This project uses NuGet's Automatic Package Restore feature to 
    resolve the Aspose.Words for .NET API reference when the project is built. 
    Please visit https://docs.nuget.org/consume/nuget-faq for more information. 

    If you do not wish to use NuGet, you can manually download Aspose.Words for .NET API 
    from http://www.aspose.com/downloads, install it, and then add a reference to it to this project. 

    For any issues, questions or suggestions, please visit the Aspose Forums: https://forum.aspose.com/
*/

using Aspose.Words;

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

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Aspose.Words Header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Aspose.Words Footer");

            doc.Save(documentPath);
        }
    }
}
