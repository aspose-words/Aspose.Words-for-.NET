// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithStyles();
string fileName = "TestFile.doc";
// Open the document.
Document doc = new Document(dataDir + fileName);

// Define style names as they are specified in the Word document.
const string paraStyle = "Heading 1";
const string runStyle = "Intense Emphasis";

// Collect paragraphs with defined styles. 
// Show the number of collected paragraphs and display the text of this paragraphs.
ArrayList paragraphs = ParagraphsByStyleName(doc, paraStyle);
Console.WriteLine(string.Format("Paragraphs with \"{0}\" styles ({1}):", paraStyle, paragraphs.Count));
foreach (Paragraph paragraph in paragraphs)
    Console.Write(paragraph.ToString(SaveFormat.Text));

// Collect runs with defined styles. 
// Show the number of collected runs and display the text of this runs.
ArrayList runs = RunsByStyleName(doc, runStyle);
Console.WriteLine(string.Format("\nRuns with \"{0}\" styles ({1}):", runStyle, runs.Count));
foreach (Run run in runs)
    Console.WriteLine(run.Range.Text);
