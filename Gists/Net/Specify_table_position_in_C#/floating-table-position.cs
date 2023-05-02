// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table wrapped by text.docx");

Table table = doc.FirstSection.Body.Tables[0];
table.AbsoluteHorizontalDistance = 10;
table.RelativeVerticalAlignment = VerticalAlignment.Center;

doc.Save(ArtifactsDir + "WorkingWithTables.FloatingTablePosition.docx");
