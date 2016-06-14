
using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Styles
{
    class ChangeStyleOfTOCLevel
    {
        public static void Run()
        {
            //ExStart:ChangeStyleOfTOCLevel
            Document doc = new Document();            
            // Retrieve the style used for the first level of the TOC and change the formatting of the style.
            doc.Styles[StyleIdentifier.Toc1].Font.Bold = true;
            //ExEnd:ChangeStyleOfTOCLevel 
            Console.WriteLine("\nTOC level style changed successfully.");
        }        
    }
}
