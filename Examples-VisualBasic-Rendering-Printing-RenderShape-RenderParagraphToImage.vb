' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim shape As Shape = CType(doc.GetChild(NodeType.Shape, 0, True), Shape)
Dim paragraph As Paragraph = CType(shape.LastParagraph, Paragraph)

' Save the node with a light pink background.
Dim options As New ImageSaveOptions(SaveFormat.Png)
options.PaperColor = Color.LightPink
dataDir = dataDir & "TestFile.RenderParagraph_out_.png"
RenderNode(paragraph, dataDir, options)
