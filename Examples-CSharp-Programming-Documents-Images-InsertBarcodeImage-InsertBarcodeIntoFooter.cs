// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
private static void InsertBarcodeIntoFooter(DocumentBuilder builder, Section section, int pageId, HeaderFooterType footerType)
{
    // Move to the footer type in the specific section.
    builder.MoveToSection(section.Document.IndexOf(section));
    builder.MoveToHeaderFooter(footerType);

    // Insert the barcode, then move to the next line and insert the ID along with the page number.
    // Use pageId if you need to insert a different barcode on each page. 0 = First page, 1 = Second page etc.    
    builder.InsertImage(System.Drawing.Image.FromFile( RunExamples.GetDataDir_WorkingWithImages() + "Barcode1.png"));
    builder.Writeln();
    builder.Write("1234567890");
    builder.InsertField("PAGE");

    // Create a right aligned tab at the right margin.
    double tabPos = section.PageSetup.PageWidth - section.PageSetup.RightMargin - section.PageSetup.LeftMargin;
    builder.CurrentParagraph.ParagraphFormat.TabStops.Add(new TabStop(tabPos, TabAlignment.Right, TabLeader.None));

    // Move to the right hand side of the page and insert the page and page total.
    builder.Write(ControlChar.Tab);
    builder.InsertField("PAGE");
    builder.Write(" of ");
    builder.InsertField("NUMPAGES");
}
