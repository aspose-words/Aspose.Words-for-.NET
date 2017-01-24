// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;
using System.Text.RegularExpressions;
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
            string FileName = FilePath + "Find and Replace.docx";
            
            Document doc = new Document(FileName);
            ReplaceOneWordWithAnother(doc);
            doc.Save(FileName);
            ReplaceTwoSimilarWords(doc);
            doc.Save(FileName);

        }
        static void ReplaceOneWordWithAnother(Document doc)
        {
            doc.Range.Replace("sad", "bad", false, true);

        }
        static void ReplaceTwoSimilarWords(Document doc)
        {
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");

        }


    }
}
