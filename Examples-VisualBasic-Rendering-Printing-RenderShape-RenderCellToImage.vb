' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim cell As Cell = CType(doc.GetChild(NodeType.Cell, 2, True), Cell) ' The third cell in the first table.
dataDir = dataDir & "TestFile.RenderCell_out_.png"
RenderNode(cell, dataDir, Nothing)
