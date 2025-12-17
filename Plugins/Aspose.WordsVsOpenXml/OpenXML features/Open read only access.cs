// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class OpenReadOnlyAccess : TestUtil
    {
        [Test]
        public void OpenReadOnlyOpenXml()
        {
            //ExStart:OpenReadOnlyOpenXml
            //GistId:702c287894827f3d4ddd2ca4b170ed45
            using WordprocessingDocument doc = WordprocessingDocument.Create(ArtifactsDir + "ReadOnly protection - OpenXml.docx", WordprocessingDocumentType.Document);
            
            // Add a main document part.
            MainDocumentPart mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = new Body();
            mainPart.Document.Append(body);

            // Add a paragraph with some text.
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            Text text = new Text("Open document as read-only");
            run.Append(text);
            paragraph.Append(run);
            body.Append(paragraph);

            // Add write protection settings.
            DocumentProtection documentProtection = new DocumentProtection
            {
                Edit = DocumentProtectionValues.ReadOnly,
                Enforcement = OnOffValue.FromBoolean(true),
                CryptographicProviderType = CryptProviderValues.RsaFull,
                CryptographicAlgorithmClass = CryptAlgorithmClassValues.Hash,
                CryptographicAlgorithmType = CryptAlgorithmValues.TypeAny,
                CryptographicAlgorithmSid = 4, // SHA-1
                Hash = "MyPassword", // Password hash (in real scenarios, you should hash the password).
                SpinCount = 100000, // Number of iterations for hashing.
                Salt = Convert.ToBase64String(Guid.NewGuid().ToByteArray()) // Random salt.
            };

            // Add the document protection settings to the document settings part.
            DocumentSettingsPart settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();
            settingsPart.Settings.Append(documentProtection);

            mainPart.Document.Save();
            //ExEnd:OpenReadOnlyOpenXml
        }
    }
}
