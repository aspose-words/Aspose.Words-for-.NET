// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

Document doc = new Document(dataDir + "Rendering.doc");

// Aspose.Words embeds full fonts by default when EmbedFullFonts is set to true. The property below can be changed
// each time a document is rendered.
PdfSaveOptions options = new PdfSaveOptions();
options.EmbedFullFonts = true;

string outPath = dataDir + "Rendering.EmbedFullFonts_out_.pdf";
// The output PDF will be embedded with all fonts found in the document.
doc.Save(outPath, options);
