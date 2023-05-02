// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document srcDoc = new Document();

// Create a custom style for the source document.
Style srcStyle = srcDoc.Styles.Add(StyleType.Paragraph, "MyStyle");
srcStyle.Font.Color = Color.Red;

// Import the source document's custom style into the destination document.
Document dstDoc = new Document();
Style newStyle = dstDoc.Styles.AddCopy(srcStyle);

// The imported style has an appearance identical to its source style.
Assert.AreEqual("MyStyle", newStyle.Name);
Assert.AreEqual(Color.Red.ToArgb(), newStyle.Font.Color.ToArgb());
