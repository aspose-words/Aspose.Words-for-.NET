// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Table wrapped by text.docx");
            
foreach (Table table in doc.FirstSection.Body.Tables)
{
    // If the table is floating type, then print its positioning properties.
    if (table.TextWrapping == TextWrapping.Around)
    {
        Console.WriteLine(table.HorizontalAnchor);
        Console.WriteLine(table.VerticalAnchor);
        Console.WriteLine(table.AbsoluteHorizontalDistance);
        Console.WriteLine(table.AbsoluteVerticalDistance);
        Console.WriteLine(table.AllowOverlap);
        Console.WriteLine(table.AbsoluteHorizontalDistance);
        Console.WriteLine(table.RelativeVerticalAlignment);
        Console.WriteLine("..............................");
    }
}
