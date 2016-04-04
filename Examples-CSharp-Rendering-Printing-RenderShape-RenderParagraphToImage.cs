// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Shape shape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
Paragraph paragraph = (Paragraph)shape.LastParagraph;

// Save the node with a light pink background.
ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
options.PaperColor = Color.LightPink;
dataDir = dataDir + "TestFile.RenderParagraph_out_.png";
RenderNode(paragraph, dataDir, options);
