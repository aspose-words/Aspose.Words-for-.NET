// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Cell cell = (Cell)doc.GetChild(NodeType.Cell, 2, true); // The third cell in the first table.
dataDir = dataDir + "TestFile.RenderCell_out_.png";
RenderNode(cell, dataDir, null);
