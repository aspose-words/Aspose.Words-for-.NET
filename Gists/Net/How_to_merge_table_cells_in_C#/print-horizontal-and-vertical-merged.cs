// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table with merged cells.docx");

SpanVisitor visitor = new SpanVisitor(doc);
doc.Accept(visitor);
