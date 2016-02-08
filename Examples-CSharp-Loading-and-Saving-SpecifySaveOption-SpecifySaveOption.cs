// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

string fileName = "TestFile RenderShape.docx";

Document doc = new Document(dataDir + fileName);

// This is the directory we want the exported images to be saved to.
string imagesDir = Path.Combine(dataDir, "Images");

// The folder specified needs to exist and should be empty.
if (Directory.Exists(imagesDir))
    Directory.Delete(imagesDir, true);

Directory.CreateDirectory(imagesDir);

// Set an option to export form fields as plain text, not as HTML input elements.
HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Html);
options.ExportTextInputFormFieldAsText = true;
options.ImagesFolder = imagesDir;

dataDir = dataDir + "Document.SaveWithOptions_out_.html";
doc.Save(dataDir, options);

