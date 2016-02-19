// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();
Document doc = new Document(dataDir + "in.doc");
// Get paragraph you want to append this AUTHOR field to
Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

// We want to insert an AUTHOR field like this:
// { AUTHOR Test1 }

// Create instance of FieldAuthor class and lets build the above field code
FieldAuthor field = (FieldAuthor)para.AppendField(FieldType.FieldAuthor, false);

// { AUTHOR Test1 }
field.AuthorName = "Test1";

// Finally update this AUTHOR field
field.Update();

dataDir = dataDir + "InsertAuthorField_out_.doc";
doc.Save(dataDir);
