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
            string fileName = FilePath + "Remove Section Breaks.docx";
            Document doc = new Document(fileName);
            RemoveSectionBreaks(doc);
        }

        private static void RemoveSectionBreaks(Document doc)
        {
            // Loop through all sections starting from the section that precedes the last one 
            // and moving to the first section.
            for (int i = doc.Sections.Count - 2; i >= 0; i--)
            {
                // Copy the content of the current section to the beginning of the last section.
                doc.LastSection.PrependContent(doc.Sections[i]);

                // Remove the copied section.
                doc.Sections[i].Remove();
            }
        }
    }
}
