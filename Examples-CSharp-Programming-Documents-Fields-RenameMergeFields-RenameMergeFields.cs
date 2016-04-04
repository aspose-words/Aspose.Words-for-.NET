// For complete examples and data files, please go to https://github.com/asposewords/Aspose_Words_NET
// The path to the documents directory.
string dataDir = RunExamples.GetDataDir_WorkingWithFields();

// Specify your document name here.
Document doc = new Document(dataDir + "RenameMergeFields.doc");

// Select all field start nodes so we can find the merge fields.
NodeCollection fieldStarts = doc.GetChildNodes(NodeType.FieldStart, true);
foreach (FieldStart fieldStart in fieldStarts)
{
    if (fieldStart.FieldType.Equals(FieldType.FieldMergeField))
    {
        MergeField mergeField = new MergeField(fieldStart);
        mergeField.Name = mergeField.Name + "_Renamed";
    }
}

dataDir = dataDir + "RenameMergeFields_out_.doc";
doc.Save(dataDir);
