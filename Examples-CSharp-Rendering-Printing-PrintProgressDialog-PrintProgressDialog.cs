// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 
// Load the documents which store the shapes we want to render.
Document doc = new Document(dataDir + "TestFile RenderShape.doc");
// Obtain the settings of the default printer
System.Drawing.Printing.PrinterSettings settings = new System.Drawing.Printing.PrinterSettings();

// The standard print controller comes with no UI
System.Drawing.Printing.PrintController standardPrintController = new System.Drawing.Printing.StandardPrintController();

// Print the document using the custom print controller
AsposeWordsPrintDocument prntDoc = new AsposeWordsPrintDocument(doc);
prntDoc.PrinterSettings = settings;
prntDoc.PrintController = standardPrintController;
prntDoc.Print();            
