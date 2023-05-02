// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document();
DocumentBuilder builder = new DocumentBuilder(doc);

// Note that AutoFitSettings does not apply to tables inserted from HTML.
builder.InsertHtml("<table>" +
                   "<tr>" +
                   "<td>Row 1, Cell 1</td>" +
                   "<td>Row 1, Cell 2</td>" +
                   "</tr>" +
                   "<tr>" +
                   "<td>Row 2, Cell 2</td>" +
                   "<td>Row 2, Cell 2</td>" +
                   "</tr>" +
                   "</table>");

doc.Save(ArtifactsDir + "WorkingWithTables.InsertTableFromHtml.docx");
