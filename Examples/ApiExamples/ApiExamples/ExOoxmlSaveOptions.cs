﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.DigitalSignatures;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Loading;
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

            Assert.That(doc.GetText().Trim(), Is.EqualTo("Hello world!"));
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

            Assert.That(((Shape)doc.GetChild(NodeType.Shape, 0, true)).MarkupLanguage, Is.EqualTo(ShapeMarkupLanguage.Vml));

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
            
            Assert.That(((Shape)doc.GetChild(NodeType.Shape, 0, true)).MarkupLanguage, Is.EqualTo(ShapeMarkupLanguage.Dml));
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

            Aspose.Words.Lists.List docList = doc.Lists[0];
            docList.IsRestartAtEachSection = restartListAtEachSection;

            // The "IsRestartAtEachSection" property will only be applicable when
            // the document's OOXML compliance level is to a standard that is newer than "OoxmlComplianceCore.Ecma376".
            OoxmlSaveOptions options = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            builder.ListFormat.List = docList;

            builder.Writeln("List item 1");
            builder.Writeln("List item 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("List item 3");
            builder.Writeln("List item 4");
            
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx", options);
            
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx");

            Assert.That(doc.Lists[0].IsRestartAtEachSection, Is.EqualTo(restartListAtEachSection));
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

            Assert.That(doc.BuiltInDocumentProperties.LastSavedTime, Is.EqualTo(new DateTime(2021, 5, 11, 6, 32, 0)));

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
                Assert.That((DateTime.Now - lastSavedTimeNew).Days < 1, Is.True);
            else
                Assert.That(lastSavedTimeNew, Is.EqualTo(new DateTime(2021, 5, 11, 6, 32, 0)));
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

            Assert.That(doc.FirstSection.Body.GetText(), Is.EqualTo(keepLegacyControlChars ? "\u0013date \\@ \"MM/dd/yyyy\"\u0014\u0015\f" : "\u001e\f"));
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

            var testedFileLength = fileInfo.Length;

            switch (compressionLevel)
            {
                case CompressionLevel.Maximum:
                    Assert.That(testedFileLength < 1269000, Is.True);
                    break;
                case CompressionLevel.Normal:
                    Assert.That(testedFileLength < 1271000, Is.True);
                    break;
                case CompressionLevel.Fast:
                    Assert.That(testedFileLength < 1280000, Is.True);
                    break;
                case CompressionLevel.SuperFast:
                    Assert.That(testedFileLength < 1276000, Is.True);
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
                    Assert.That(prevFileSize < fileSize, Is.True);

                    TestUtil.CopyStream(outputFileStream, stream);
                    Assert.That(TestUtil.DumpArray(stream.ToArray(), 0, 10), Is.EqualTo(fileSignatures[i]));

                    prevFileSize = fileSize;
                }
            }
        }

        [Test]
        public void ExportGeneratorName()
        {
            //ExStart
            //ExFor:SaveOptions.ExportGeneratorName
            //ExSummary:Shows how to disable adding name and version of Aspose.Words into produced files.
            Document doc = new Document();

            // Use https://docs.aspose.com/words/net/generator-or-producer-name-included-in-output-documents/ to know how to check the result.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { ExportGeneratorName = false };

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.ExportGeneratorName.docx", saveOptions);
            //ExEnd
        }

        [TestCase(SaveFormat.Docx, "docx")]
        [TestCase(SaveFormat.Docm, "docm")]
        [TestCase(SaveFormat.Dotm, "dotm")]
        [TestCase(SaveFormat.Dotx, "dotx")]
        [TestCase(SaveFormat.FlatOpc, "flatopc")]
        //ExStart
        //ExFor:SaveOptions.ProgressCallback
        //ExFor:IDocumentSavingCallback
        //ExFor:IDocumentSavingCallback.Notify(DocumentSavingArgs)
        //ExFor:DocumentSavingArgs.EstimatedProgress
        //ExSummary:Shows how to manage a document while saving to docx.
        public void ProgressCallback(SaveFormat saveFormat, string ext)
        {
            Document doc = new Document(MyDir + "Big document.docx");

            // Following formats are supported: Docx, FlatOpc, Docm, Dotm, Dotx.
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(saveFormat)
            {
                ProgressCallback = new SavingProgressCallback()
            };

            var exception = Assert.Throws<OperationCanceledException>(() =>
                doc.Save(ArtifactsDir + $"OoxmlSaveOptions.ProgressCallback.{ext}", saveOptions));
            Assert.That(exception?.Message.Contains("EstimatedProgress"), Is.True);
        }

        /// <summary>
        /// Saving progress callback. Cancel a document saving after the "MaxDuration" seconds.
        /// </summary>
        public class SavingProgressCallback : IDocumentSavingCallback
        {
            /// <summary>
            /// Ctr.
            /// </summary>
            public SavingProgressCallback()
            {
                mSavingStartedAt = DateTime.Now;
            }

            /// <summary>
            /// Callback method which called during document saving.
            /// </summary>
            /// <param name="args">Saving arguments.</param>
            public void Notify(DocumentSavingArgs args)
            {
                DateTime canceledAt = DateTime.Now;
                double ellapsedSeconds = (canceledAt - mSavingStartedAt).TotalSeconds;
                if (ellapsedSeconds > MaxDuration)
                    throw new OperationCanceledException($"EstimatedProgress = {args.EstimatedProgress}; CanceledAt = {canceledAt}");
            }

            /// <summary>
            /// Date and time when document saving is started.
            /// </summary>
            private readonly DateTime mSavingStartedAt;

            /// <summary>
            /// Maximum allowed duration in sec.
            /// </summary>
            private const double MaxDuration = 0.01d;
        }
        //ExEnd

        [Test]
        public void Zip64ModeOption()
        {
            //ExStart:Zip64ModeOption
            //GistId:e386727403c2341ce4018bca370a5b41
            //ExFor:OoxmlSaveOptions.Zip64Mode
            //ExFor:Zip64Mode
            //ExSummary:Shows how to use ZIP64 format extensions.
            Random random = new Random();
            DocumentBuilder builder = new DocumentBuilder();

            for (int i = 0; i < 10000; i++)
            {
                using (Bitmap bmp = new Bitmap(5, 5))
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.FromArgb(random.Next(0, 254), random.Next(0, 254), random.Next(0, 254)));
                    using (MemoryStream ms = new MemoryStream())
                    {
                        bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                        builder.InsertImage(ms.ToArray());
                    }
                }
            }

            builder.Document.Save(ArtifactsDir + "OoxmlSaveOptions.Zip64ModeOption.docx", 
                new OoxmlSaveOptions { Zip64Mode = Zip64Mode.Always });
            //ExEnd:Zip64ModeOption
        }

        [Test]
        public void DigitalSignature()
        {
            //ExStart:DigitalSignature
            //GistId:5f20ac02cb42c6b08481aa1c5b0cd3db
            //ExFor:OoxmlSaveOptions.DigitalSignatureDetails
            //ExFor:DigitalSignatureDetails
            //ExFor:DigitalSignatureDetails.#ctor(CertificateHolder, SignOptions)
            //ExFor:DigitalSignatureDetails.CertificateHolder
            //ExFor:DigitalSignatureDetails.SignOptions
            //ExSummary:Shows how to sign OOXML document.
            Document doc = new Document(MyDir + "Document.docx");

            CertificateHolder certificateHolder = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            DigitalSignatureDetails digitalSignatureDetails = new DigitalSignatureDetails(
                certificateHolder,
                new SignOptions() { Comments = "Some comments", SignTime = DateTime.Now });

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.DigitalSignatureDetails = digitalSignatureDetails;

            Assert.That(digitalSignatureDetails.CertificateHolder, Is.EqualTo(certificateHolder));
            Assert.That(digitalSignatureDetails.SignOptions.Comments, Is.EqualTo("Some comments"));

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.DigitalSignature.docx", saveOptions);
            //ExEnd:DigitalSignature
        }

        [Test]
        public void UpdateAmbiguousTextFont()
        {
            //ExStart:UpdateAmbiguousTextFont
            //GistId:1a265b92fa0019b26277ecfef3c20330
            //ExFor:SaveOptions.UpdateAmbiguousTextFont
            //ExSummary:Shows how to update the font to match the character code being used.
            Document doc = new Document(MyDir + "Special symbol.docx");
            Run run = doc.FirstSection.Body.FirstParagraph.Runs[0];
            Console.WriteLine(run.Text); // ฿
            Console.WriteLine(run.Font.Name); // Arial

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateAmbiguousTextFont = true;
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.UpdateAmbiguousTextFont.docx", saveOptions);

            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.UpdateAmbiguousTextFont.docx");
            run = doc.FirstSection.Body.FirstParagraph.Runs[0];
            Console.WriteLine(run.Text); // ฿
            Console.WriteLine(run.Font.Name); // Angsana New
            //ExEnd:UpdateAmbiguousTextFont
        }
    }
}