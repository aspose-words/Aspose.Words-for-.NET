// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

// Open the document.
Document doc = new Document(dataDir + "TestFile Multipage TIFF.doc");

// Save the document as multipage TIFF.
doc.Save(dataDir + "TestFile Multipage TIFF_out_.tiff");
            
//Create an ImageSaveOptions object to pass to the Save method
ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
options.PageIndex = 0;
options.PageCount = 2;
options.TiffCompression = TiffCompression.Ccitt4;
options.Resolution = 160;
dataDir = dataDir + "TestFileWithOptions_out_.tiff";
doc.Save(dataDir, options);
