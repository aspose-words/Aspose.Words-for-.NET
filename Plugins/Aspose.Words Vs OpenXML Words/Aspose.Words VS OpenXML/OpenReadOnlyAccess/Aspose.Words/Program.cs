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
        private static string FilePath = @"..\..\..\..\Sample Files\";
        private static string fileName = FilePath + "OpenReadOnlyAccess.docx";
        
        static void Main(string[] args)
        {
            OpenWordprocessingDocumentReadonly(fileName);
        }

        private static void OpenWordprocessingDocumentReadonly(string fileName)
        {
            Document doc = new Document(fileName, new LoadOptions("1234"));
            DocumentBuilder db = new DocumentBuilder(doc);
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            db.Writeln(txt);
            doc.Save(fileName);
        }
    }
}
