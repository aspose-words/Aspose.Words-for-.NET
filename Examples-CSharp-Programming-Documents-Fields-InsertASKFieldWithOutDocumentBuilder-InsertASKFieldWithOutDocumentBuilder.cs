// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();
Document doc = new Document(dataDir + "in.doc");
// Get paragraph you want to append this Ask field to
Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

// We want to insert an Ask field like this:
// { ASK \"Test 1\" Test2 \\d Test3 \\o }

// Create instance of FieldAsk class and lets build the above field code
FieldAsk field = (FieldAsk)para.AppendField(FieldType.FieldAsk, false);

// { ASK \"Test 1\" " }
field.BookmarkName = "Test 1";

// { ASK \"Test 1\" Test2 }
field.PromptText = "Test2";

// { ASK \"Test 1\" Test2 \\d Test3 }
field.DefaultResponse = "Test3";

// { ASK \"Test 1\" Test2 \\d Test3 \\o }
field.PromptOnceOnMailMerge = true;

// Finally update this Ask field
field.Update();

dataDir = dataDir + "InsertASKFieldWithOutDocumentBuilder_out_.doc";
doc.Save(dataDir);

