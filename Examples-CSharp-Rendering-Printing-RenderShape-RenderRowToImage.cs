// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Row row = (Row)doc.GetChild(NodeType.Row, 0, true); // The first row in the first table.

dataDir = dataDir + "TestFile.RenderRow_out_.png";
RenderNode(row, dataDir, null);
