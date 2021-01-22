// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string fileName = FilePath + "RemoveHeaderFooter.docx";

            RemoveHeadersAndFooters(fileName);

        }

        public static void RemoveHeadersAndFooters(string filename)
        {
            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(filename, true))
            {
                var mainDocumentPart = doc.MainDocumentPart;

                // Count the header and footer parts and continue if there 
                // are any.
                if (mainDocumentPart.HeaderParts.Any() || mainDocumentPart.FooterParts.Any())
                {
                    // Remove the header and footer parts.
                    mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                    mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

                    // Get a reference to the root element of the main
                    // document part.
                    Document document = mainDocumentPart.Document;

                    // Remove all references to the headers and footers.

                    // First, create a list of all descendants of type
                    // HeaderReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var headers = document.Descendants<HeaderReference>().ToList();

                    foreach (var header in headers)
                        header.Remove();

                    // First, create a list of all descendants of type
                    // FooterReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var footers = document.Descendants<FooterReference>().ToList();

                    foreach (var footer in footers)
                        footer.Remove();

                    document.Save();
                }
            }
        }
    }
}