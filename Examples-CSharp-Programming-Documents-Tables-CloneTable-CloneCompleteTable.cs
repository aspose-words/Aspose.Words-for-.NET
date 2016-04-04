// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
Document doc = new Document(dataDir + "Table.SimpleTable.doc");

// Retrieve the first table in the document.
Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

// Create a clone of the table.
Table tableClone = (Table)table.Clone(true);

// Insert the cloned table into the document after the original
table.ParentNode.InsertAfter(tableClone, table);

// Insert an empty paragraph between the two tables or else they will be combined into one
// upon save. This has to do with document validation.
table.ParentNode.InsertAfter(new Paragraph(doc), table);
dataDir = dataDir + "Table.CloneTableAndInsert_out_.doc";
           
// Save the document to disk.
doc.Save(dataDir);
