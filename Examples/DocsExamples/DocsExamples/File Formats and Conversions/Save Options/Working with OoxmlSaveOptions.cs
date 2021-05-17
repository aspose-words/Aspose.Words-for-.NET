using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithOoxmlSaveOptions : DocsExamplesBase
    {
        [Test]
        public void EncryptDocxWithPassword()
        {
            //ExStart:EncryptDocxWithPassword
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "password" };

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.EncryptDocxWithPassword.docx", saveOptions);
            //ExEnd:EncryptDocxWithPassword
        }

        [Test]
        public void OoxmlComplianceIso29500_2008_Strict()
        {
            //ExStart:OoxmlComplianceIso29500_2008_Strict
            Document doc = new Document(MyDir + "Document.docx");

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);
            
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions() { Compliance = OoxmlCompliance.Iso29500_2008_Strict };

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.OoxmlComplianceIso29500_2008_Strict.docx", saveOptions);
            //ExEnd:OoxmlComplianceIso29500_2008_Strict
        }

        [Test]
        public void UpdateLastSavedTimeProperty()
        {
            //ExStart:UpdateLastSavedTimeProperty
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { UpdateLastSavedTimeProperty = true };

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.UpdateLastSavedTimeProperty.docx", saveOptions);
            //ExEnd:UpdateLastSavedTimeProperty
        }

        [Test]
        public void KeepLegacyControlChars()
        {
            //ExStart:KeepLegacyControlChars
            Document doc = new Document(MyDir + "Legacy control character.doc");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.FlatOpc) { KeepLegacyControlChars = true };

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.KeepLegacyControlChars.docx", saveOptions);
            //ExEnd:KeepLegacyControlChars
        }

        [Test]
        public void SetCompressionLevel()
        {
            //ExStart:SetCompressionLevel
            Document doc = new Document(MyDir + "Document.docx");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { CompressionLevel = CompressionLevel.SuperFast };

            doc.Save(ArtifactsDir + "WorkingWithOoxmlSaveOptions.SetCompressionLevel.docx", saveOptions);
            //ExEnd:SetCompressionLevel
        }
    }
}