// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExOoxmlSaveOptions : ApiExampleBase
    {
        [Test]
        public void Password()
        {
            //ExStart
            //ExFor:OoxmlSaveOptions.Password
            //ExSummary:Shows how to create a password encrypted Office Open XML document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "MyPassword";

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.Password.docx", saveOptions);

            // We will not be able to open this document with Microsoft Word or
            // Aspose.Words without providing the correct password.
            Assert.Throws<IncorrectPasswordException>(() =>
                doc = new Document(ArtifactsDir + "OoxmlSaveOptions.Password.docx"));

            // Open the encrypted document by passing the correct password in a LoadOptions object.
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.Password.docx", new LoadOptions("MyPassword"));

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void Iso29500Strict()
        {
            //ExStart
            //ExFor:CompatibilityOptions
            //ExFor:CompatibilityOptions.OptimizeFor(MsWordVersion)
            //ExFor:OoxmlSaveOptions
            //ExFor:OoxmlSaveOptions.#ctor
            //ExFor:OoxmlSaveOptions.SaveFormat
            //ExFor:OoxmlCompliance
            //ExFor:OoxmlSaveOptions.Compliance
            //ExFor:ShapeMarkupLanguage
            //ExSummary:Shows how to set an OOXML compliance specification for a saved document to adhere to.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // If we configure compatibility options to comply with Microsoft Word 2003,
            // inserting an image will define its shape using VML.
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            Assert.AreEqual(ShapeMarkupLanguage.Vml, ((Shape)doc.GetChild(NodeType.Shape, 0, true)).MarkupLanguage);

            // The "ISO/IEC 29500:2008" OOXML standard does not support VML shapes.
            // If we set the "Compliance" property of the SaveOptions object to "OoxmlCompliance.Iso29500_2008_Strict",
            // any document we save while passing this object will have to follow that standard. 
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Strict,
                SaveFormat = SaveFormat.Docx
            };

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

            // Our saved document defines the shape using DML to adhere to the "ISO/IEC 29500:2008" OOXML standard.
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.Iso29500Strict.docx");
            
            Assert.AreEqual(ShapeMarkupLanguage.Dml, ((Shape)doc.GetChild(NodeType.Shape, 0, true)).MarkupLanguage);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void RestartingDocumentList(bool restartListAtEachSection)
        {
            //ExStart
            //ExFor:List.IsRestartAtEachSection
            //ExFor:OoxmlCompliance
            //ExFor:OoxmlSaveOptions.Compliance
            //ExSummary:Shows how to configure a list to restart numbering at each section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.Lists.Add(ListTemplate.NumberDefault);

            Aspose.Words.Lists.List list = doc.Lists[0];
            list.IsRestartAtEachSection = restartListAtEachSection;

            // The "IsRestartAtEachSection" property will only be applicable when
            // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
            OoxmlSaveOptions options = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            builder.ListFormat.List = list;

            builder.Writeln("List item 1");
            builder.Writeln("List item 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("List item 3");
            builder.Writeln("List item 4");
            
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx", options);
            
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx");

            Assert.AreEqual(restartListAtEachSection, doc.Lists[0].IsRestartAtEachSection);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void LastSavedTime(bool updateLastSavedTimeProperty)
        {
            //ExStart
            //ExFor:SaveOptions.UpdateLastSavedTimeProperty
            //ExSummary:Shows how to determine whether to preserve the document's "Last saved time" property when saving.
            Document doc = new Document(MyDir + "Document.docx");

            Assert.AreEqual(new DateTime(2020, 7, 30, 5, 27, 0), 
                doc.BuiltInDocumentProperties.LastSavedTime);

            // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
            // and then pass it to the document's saving method to modify how we save the document.
            // Set the "UpdateLastSavedTimeProperty" property to "true" to
            // set the output document's "Last saved time" built-in property to the current date/time.
            // Set the "UpdateLastSavedTimeProperty" property to "false" to
            // preserve the original value of the input document's "Last saved time" built-in property.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateLastSavedTimeProperty = updateLastSavedTimeProperty;

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.LastSavedTime.docx", saveOptions);

            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.LastSavedTime.docx");
            DateTime lastSavedTimeNew = doc.BuiltInDocumentProperties.LastSavedTime;

            if (updateLastSavedTimeProperty)
                Assert.That(DateTime.Now, Is.EqualTo(lastSavedTimeNew).Within(1).Days);
            else
                Assert.AreEqual(new DateTime(2020, 7, 30, 5, 27, 0), 
                    lastSavedTimeNew);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void KeepLegacyControlChars(bool keepLegacyControlChars)
        {
            //ExStart
            //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
            //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
            //ExSummary:Shows how to support legacy control characters when converting to .docx.
            Document doc = new Document(MyDir + "Legacy control character.doc");

            // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
            // and then pass it to the document's saving method to modify how we save the document.
            // Set the "KeepLegacyControlChars" property to "true" to preserve
            // the "ShortDateTime" legacy character while saving.
            // Set the "KeepLegacyControlChars" property to "false" to remove
            // the "ShortDateTime" legacy character from the output document.
            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.Docx);
            so.KeepLegacyControlChars = keepLegacyControlChars;
 
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.KeepLegacyControlChars.docx", so);
            
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.KeepLegacyControlChars.docx");

            Assert.AreEqual(keepLegacyControlChars ? "\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f" : "\u001e\f",
                doc.FirstSection.Body.GetText());
            //ExEnd
        }

        [TestCase(CompressionLevel.Maximum)]
        [TestCase(CompressionLevel.Fast)]
        [TestCase(CompressionLevel.Normal)]
        [TestCase(CompressionLevel.SuperFast)]
        public void DocumentCompression(CompressionLevel compressionLevel)
        {
            //ExStart
            //ExFor:OoxmlSaveOptions.CompressionLevel
            //ExFor:CompressionLevel
            //ExSummary:Shows how to specify the compression level to use while saving an OOXML document.
            Document doc = new Document(MyDir + "Big document.docx");

            // When we save the document to an OOXML format, we can create an OoxmlSaveOptions object
            // and then pass it to the document's saving method to modify how we save the document.
            // Set the "CompressionLevel" property to "CompressionLevel.Maximum" to apply the strongest and slowest compression.
            // Set the "CompressionLevel" property to "CompressionLevel.Normal" to apply
            // the default compression that Aspose.Words uses while saving OOXML documents.
            // Set the "CompressionLevel" property to "CompressionLevel.Fast" to apply a faster and weaker compression.
            // Set the "CompressionLevel" property to "CompressionLevel.SuperFast" to apply
            // the default compression that Microsoft Word uses.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            saveOptions.CompressionLevel = compressionLevel;
            
            Stopwatch st = Stopwatch.StartNew();
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.DocumentCompression.docx", saveOptions);
            st.Stop();

            FileInfo fileInfo = new FileInfo(ArtifactsDir + "OoxmlSaveOptions.DocumentCompression.docx");

            Console.WriteLine($"Saving operation done using the \"{compressionLevel}\" compression level:");
            Console.WriteLine($"\tDuration:\t{st.ElapsedMilliseconds} ms");
            Console.WriteLine($"\tFile Size:\t{fileInfo.Length} bytes");
            //ExEnd

            switch (compressionLevel)
            {
                case CompressionLevel.Maximum:
                    Assert.That(1266000, Is.AtLeast(fileInfo.Length));
                    break;
                case CompressionLevel.Normal:
                    Assert.That(1267000, Is.LessThan(fileInfo.Length));
                    break;
                case CompressionLevel.Fast:
                    Assert.That(1269000, Is.LessThan(fileInfo.Length));
                    break;
                case CompressionLevel.SuperFast:
                    Assert.That(1271000, Is.LessThan(fileInfo.Length));
                    break;
            }
        }

        [Test]
        public void CheckFileSignatures()
        {
            CompressionLevel[] compressionLevels = {
                CompressionLevel.Maximum,
                CompressionLevel.Normal,
                CompressionLevel.Fast,
                CompressionLevel.SuperFast
            };

#if JAVA
            string[] fileSignatures = new string[]
            {
                "50 4B 03 04 14 00 08 08 08 00 ",
                "50 4B 03 04 14 00 08 08 08 00 ",
                "50 4B 03 04 14 00 08 08 08 00 ",
                "50 4B 03 04 14 00 08 08 08 00 "
            };
#else
            string[] fileSignatures = {
                "50 4B 03 04 14 00 02 00 08 00 ",
                "50 4B 03 04 14 00 00 00 08 00 ",
                "50 4B 03 04 14 00 04 00 08 00 ",
                "50 4B 03 04 14 00 06 00 08 00 "
            };
#endif

            Document doc = new Document();
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);

            long prevFileSize = 0;
            for (int i = 0; i < fileSignatures.Length; i++)
            {
                saveOptions.CompressionLevel = compressionLevels[i];
                doc.Save(ArtifactsDir + "OoxmlSaveOptions.CheckFileSignatures.docx", saveOptions);

                using (MemoryStream stream = new MemoryStream())
                using (FileStream outputFileStream = File.Open(ArtifactsDir + "OoxmlSaveOptions.CheckFileSignatures.docx", FileMode.Open))
                {
                    long fileSize = outputFileStream.Length;
                    Assert.That(prevFileSize < fileSize);

                    TestUtil.CopyStream(outputFileStream, stream);
                    Assert.AreEqual(fileSignatures[i], TestUtil.DumpArray(stream.ToArray(), 0, 10));

                    prevFileSize = fileSize;
                }
            }
        }
    }
}