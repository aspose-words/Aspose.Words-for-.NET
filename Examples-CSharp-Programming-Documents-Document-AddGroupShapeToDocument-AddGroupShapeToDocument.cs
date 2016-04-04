// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

Document doc = new Document();
doc.EnsureMinimum();
GroupShape gs = new GroupShape(doc);

Shape shape = new Shape(doc, Aspose.Words.Drawing.ShapeType.AccentBorderCallout1);
shape.Width = 100;
shape.Height = 100;
gs.AppendChild(shape);

Shape shape1 = new Shape(doc, Aspose.Words.Drawing.ShapeType.ActionButtonBeginning);
shape1.Left = 100;
shape1.Width = 100;
shape1.Height = 200;
gs.AppendChild(shape1);

gs.Width = 200;
gs.Height = 200;

gs.CoordSize = new System.Drawing.Size(200, 200);

DocumentBuilder builder = new DocumentBuilder(doc);
builder.InsertNode(gs);


dataDir = dataDir + "groupshape-doc_out_.doc";

// Save the document to disk.
doc.Save(dataDir);
