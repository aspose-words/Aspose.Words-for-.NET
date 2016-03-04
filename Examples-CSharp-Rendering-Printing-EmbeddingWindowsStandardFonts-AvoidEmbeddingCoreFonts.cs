// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

Document doc = new Document(dataDir + "Rendering.doc");

// To disable embedding of core fonts and subsuite PDF type 1 fonts set UseCoreFonts to true.
PdfSaveOptions options = new PdfSaveOptions();
options.UseCoreFonts = true;

string outPath = dataDir + "Rendering.DisableEmbedWindowsFonts_out_.pdf";
// The output PDF will not be embedded with core fonts such as Arial, Times New Roman etc.
doc.Save(outPath);
