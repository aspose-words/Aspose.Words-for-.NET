// For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-.NET.git.
Document doc = new Document(MyDir + "Tables.docx");

Table table = (Table) doc.GetChild(NodeType.Table, 0, true);
table.Title = "Test title";
table.Description = "Test description";

OoxmlSaveOptions options = new OoxmlSaveOptions { Compliance = OoxmlCompliance.Iso29500_2008_Strict };

doc.CompatibilityOptions.OptimizeFor(Aspose.Words.Settings.MsWordVersion.Word2016);

doc.Save(ArtifactsDir + "WorkingWithTableStylesAndFormatting.TableTitleAndDescription.docx", options);
