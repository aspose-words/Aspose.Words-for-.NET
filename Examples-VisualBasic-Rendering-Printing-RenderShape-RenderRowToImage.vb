' For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Dim row As Row = CType(doc.GetChild(NodeType.Row, 0, True), Row) ' The first row in the first table.
dataDir = dataDir & "TestFile.RenderRow_out_.png"
RenderNode(row, dataDir, Nothing)
