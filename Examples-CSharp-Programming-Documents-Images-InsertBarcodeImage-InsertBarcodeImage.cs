// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithImages();
// Create a blank documenet.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// The number of pages the document should have.
int numPages = 4;
// The document starts with one section, insert the barcode into this existing section.
InsertBarcodeIntoFooter(builder, doc.FirstSection, 1, HeaderFooterType.FooterPrimary);

for (int i = 1; i < numPages; i++)
{
    // Clone the first section and add it into the end of the document.
    Section cloneSection = (Section)doc.FirstSection.Clone(false);
    cloneSection.PageSetup.SectionStart = SectionStart.NewPage;
    doc.AppendChild(cloneSection);

    // Insert the barcode and other information into the footer of the section.
    InsertBarcodeIntoFooter(builder, cloneSection, i, HeaderFooterType.FooterPrimary);
}

dataDir  = dataDir + "Document_out_.docx";
// Save the document as a PDF to disk. You can also save this directly to a stream.
doc.Save(dataDir);
