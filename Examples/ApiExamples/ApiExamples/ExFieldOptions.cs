﻿// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using NUnit.Framework;
#if !NETFRAMEWORK
using SkiaSharp;
using Image = SkiaSharp.SKBitmap;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExFieldOptions : ApiExampleBase
    {
        [Test]
        public void CurrentUser()
        {
            //ExStart
            //ExFor:Document.UpdateFields
            //ExFor:FieldOptions.CurrentUser
            //ExFor:UserInformation
            //ExFor:UserInformation.Name
            //ExFor:UserInformation.Initials
            //ExFor:UserInformation.Address
            //ExFor:UserInformation.DefaultUser
            //ExSummary:Shows how to set user details, and display them using fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a UserInformation object and set it as the data source for fields that display user information.
            UserInformation userInformation = new UserInformation
            {
                Name = "John Doe",
                Initials = "J. D.",
                Address = "123 Main Street"
            };
            doc.FieldOptions.CurrentUser = userInformation;

            // Insert USERNAME, USERINITIALS, and USERADDRESS fields, which display values of
            // the respective properties of the UserInformation object that we have created above. 
            Assert.That(builder.InsertField(" USERNAME ").Result, Is.EqualTo(userInformation.Name));
            Assert.That(builder.InsertField(" USERINITIALS ").Result, Is.EqualTo(userInformation.Initials));
            Assert.That(builder.InsertField(" USERADDRESS ").Result, Is.EqualTo(userInformation.Address));

            // The field options object also has a static default user that fields from all documents can refer to.
            UserInformation.DefaultUser.Name = "Default User";
            UserInformation.DefaultUser.Initials = "D. U.";
            UserInformation.DefaultUser.Address = "One Microsoft Way";
            doc.FieldOptions.CurrentUser = UserInformation.DefaultUser;

            Assert.That(builder.InsertField(" USERNAME ").Result, Is.EqualTo("Default User"));
            Assert.That(builder.InsertField(" USERINITIALS ").Result, Is.EqualTo("D. U."));
            Assert.That(builder.InsertField(" USERADDRESS ").Result, Is.EqualTo("One Microsoft Way"));

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FieldOptions.CurrentUser.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.CurrentUser.docx");

            Assert.That(doc.FieldOptions.CurrentUser, Is.Null);

            FieldUserName fieldUserName = (FieldUserName)doc.Range.Fields[0];

            Assert.That(fieldUserName.UserName, Is.Null);
            Assert.That(fieldUserName.Result, Is.EqualTo("Default User"));

            FieldUserInitials fieldUserInitials = (FieldUserInitials)doc.Range.Fields[1];

            Assert.That(fieldUserInitials.UserInitials, Is.Null);
            Assert.That(fieldUserInitials.Result, Is.EqualTo("D. U."));

            FieldUserAddress fieldUserAddress = (FieldUserAddress)doc.Range.Fields[2];

            Assert.That(fieldUserAddress.UserAddress, Is.Null);
            Assert.That(fieldUserAddress.Result, Is.EqualTo("One Microsoft Way"));
        }

        [Test]
        public void FileName()
        {
            //ExStart
            //ExFor:FieldOptions.FileName
            //ExFor:FieldFileName
            //ExFor:FieldFileName.IncludeFullPath
            //ExSummary:Shows how to use FieldOptions to override the default value for the FILENAME field.
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToDocumentEnd();
            builder.Writeln();

            // This FILENAME field will display the local system file name of the document we loaded.
            FieldFileName field = (FieldFileName)builder.InsertField(FieldType.FieldFileName, true);
            field.Update();

            Assert.That(field.GetFieldCode(), Is.EqualTo(" FILENAME "));
            Assert.That(field.Result, Is.EqualTo("Document.docx"));

            builder.Writeln();

            // By default, the FILENAME field shows the file's name, but not its full local file system path.
            // We can set a flag to make it show the full file path.
            field = (FieldFileName)builder.InsertField(FieldType.FieldFileName, true);
            field.IncludeFullPath = true;
            field.Update();

            Assert.That(field.Result, Is.EqualTo(MyDir + "Document.docx"));

            // We can also set a value for this property to
            // override the value that the FILENAME field displays.
            doc.FieldOptions.FileName = "FieldOptions.FILENAME.docx";
            field.Update();

            Assert.That(field.GetFieldCode(), Is.EqualTo(" FILENAME  \\p"));
            Assert.That(field.Result, Is.EqualTo("FieldOptions.FILENAME.docx"));

            doc.UpdateFields();
            doc.Save(ArtifactsDir + doc.FieldOptions.FileName);
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.FILENAME.docx");

            Assert.That(doc.FieldOptions.FileName, Is.Null);
            TestUtil.VerifyField(FieldType.FieldFileName, " FILENAME ", "FieldOptions.FILENAME.docx", doc.Range.Fields[0]);
        }

        [Test]
        public void Bidi()
        {
            //ExStart
            //ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
            //ExSummary:Shows how to use FieldOptions to ensure that field updating fully supports bi-directional text.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Ensure that any field operation involving right-to-left text is performs as expected. 
            doc.FieldOptions.IsBidiTextSupportedOnUpdate = true;

            // Use a document builder to insert a field that contains the right-to-left text.
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "עֶשְׂרִים", "שְׁלוֹשִׁים", "אַרְבָּעִים", "חֲמִשִּׁים", "שִׁשִּׁים" }, 0);
            comboBox.CalculateOnExit = true;

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FieldOptions.Bidi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.Bidi.docx");

            Assert.That(doc.FieldOptions.IsBidiTextSupportedOnUpdate, Is.False);

            comboBox = doc.Range.FormFields[0];

            Assert.That(comboBox.Result, Is.EqualTo("עֶשְׂרִים"));
        }

        [Test]
        public void LegacyNumberFormat()
        {
            //ExStart
            //ExFor:FieldOptions.LegacyNumberFormat
            //ExSummary:Shows how enable legacy number formatting for fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("= 2 + 3 \\# $##");

            Assert.That(field.Result, Is.EqualTo("$ 5"));

            doc.FieldOptions.LegacyNumberFormat = true;
            field.Update();

            Assert.That(field.Result, Is.EqualTo("$5"));
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.FieldOptions.LegacyNumberFormat, Is.False);
            TestUtil.VerifyField(FieldType.FieldFormula, "= 2 + 3 \\# $##", "$5", doc.Range.Fields[0]);
        }

        [Test]
        public void PreProcessCulture()
        {
            //ExStart
            //ExFor:FieldOptions.PreProcessCulture
            //ExSummary:Shows how to set the preprocess culture.
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set the culture according to which some fields will format their displayed values.
            doc.FieldOptions.PreProcessCulture = new CultureInfo("de-DE");

            Field field = builder.InsertField(" DOCPROPERTY CreateTime");

            // The DOCPROPERTY field will display its result formatted according to the preprocess culture
            // we have set to German. The field will display the date/time using the "dd.mm.yyyy hh:mm" format.
            Assert.That(Regex.Match(field.Result, @"\d{2}[.]\d{2}[.]\d{4} \d{2}[:]\d{2}").Success, Is.True);

            doc.FieldOptions.PreProcessCulture = CultureInfo.InvariantCulture;
            field.Update();

            // After switching to the invariant culture, the DOCPROPERTY field will use the "mm/dd/yyyy hh:mm" format.
            Assert.That(Regex.Match(field.Result, @"\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}").Success, Is.True);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.FieldOptions.PreProcessCulture, Is.Null);
            Assert.That(Regex.Match(doc.Range.Fields[0].Result, @"\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}").Success, Is.True);
        }

        [Test]
        public void TableOfAuthorityCategories()
        {
            //ExStart
            //ExFor:FieldOptions.ToaCategories
            //ExFor:ToaCategories
            //ExFor:ToaCategories.Item(Int32)
            //ExFor:ToaCategories.DefaultCategories
            //ExSummary:Shows how to specify a set of categories for TOA fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // TOA fields can filter their entries by categories defined in this collection.
            ToaCategories toaCategories = new ToaCategories();
            doc.FieldOptions.ToaCategories = toaCategories;

            // This collection of categories comes with default values, which we can overwrite with custom values.
            Assert.That(toaCategories[1], Is.EqualTo("Cases"));
            Assert.That(toaCategories[2], Is.EqualTo("Statutes"));

            toaCategories[1] = "My Category 1";
            toaCategories[2] = "My Category 2";

            // We can always access the default values via this collection.
            Assert.That(ToaCategories.DefaultCategories[1], Is.EqualTo("Cases"));
            Assert.That(ToaCategories.DefaultCategories[2], Is.EqualTo("Statutes"));

            // Insert 2 TOA fields. TOA fields create an entry for each TA field in the document.
            // Use the "\c" switch to select the index of a category from our collection.
            //  With this switch, a TOA field will only pick up entries from TA fields that
            // also have a "\c" switch with a matching category index. Each TOA field will also display
            // the name of the category that its "\c" switch points to.
            builder.InsertField("TOA \\c 1 \\h", null);
            builder.InsertField("TOA \\c 2 \\h", null);
            builder.InsertBreak(BreakType.PageBreak);

            // Insert TOA entries across 2 categories. Our first TOA field will receive one entry,
            // from the second TA field whose "\c" switch also points to the first category.
            // The second TOA field will have two entries from the other two TA fields.
            builder.InsertField("TA \\c 2 \\l \"entry 1\"");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertField("TA \\c 1 \\l \"entry 2\"");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertField("TA \\c 2 \\l \"entry 3\"");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FieldOptions.TOA.Categories.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.TOA.Categories.docx");

            Assert.That(doc.FieldOptions.ToaCategories, Is.Null);

            TestUtil.VerifyField(FieldType.FieldTOA, "TOA \\c 1 \\h", "My Category 1\rentry 2\t3\r", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldTOA, "TOA \\c 2 \\h",
                "My Category 2\r" +
                "entry 1\t2\r" +
                "entry 3\t4\r", doc.Range.Fields[1]);
        }

        [Test]
        public void UseInvariantCultureNumberFormat()
        {
            //ExStart
            //ExFor:FieldOptions.UseInvariantCultureNumberFormat
            //ExSummary:Shows how to format numbers according to the invariant culture.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            Field field = builder.InsertField(" = 1234567,89 \\# $#,###,###.##");
            field.Update();

            // Sometimes, fields may not format their numbers correctly under certain cultures. 
            Assert.That(doc.FieldOptions.UseInvariantCultureNumberFormat, Is.False);
            Assert.That(field.Result, Is.EqualTo("$1.234.567,89 ,     "));

            // To fix this, we could change the culture for the entire thread.
            // Another way to fix this is to set this flag,
            // which gets all fields to use the invariant culture when formatting numbers.
            // This way allows us to avoid changing the culture for the entire thread.
            doc.FieldOptions.UseInvariantCultureNumberFormat = true;
            field.Update();
            Assert.That(field.Result, Is.EqualTo("$1.234.567,89"));
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.That(doc.FieldOptions.UseInvariantCultureNumberFormat, Is.False);
            TestUtil.VerifyField(FieldType.FieldFormula, " = 1234567,89 \\# $#,###,###.##", "$1.234.567,89", doc.Range.Fields[0]);
        }

        //ExStart
        //ExFor:FieldOptions.FieldUpdateCultureProvider
        //ExFor:IFieldUpdateCultureProvider
        //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
        //ExSummary:Shows how to specify a culture which parses date/time formatting for each field.
        [Test]//ExSkip
        public void DefineDateTimeFormatting()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(FieldType.FieldTime, true);

            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;

            // Set a provider that returns a culture object specific to each field.
            doc.FieldOptions.FieldUpdateCultureProvider = new FieldUpdateCultureProvider();

            FieldTime fieldDate = (FieldTime)doc.Range.Fields[0];
            if (fieldDate.LocaleId != (int)EditingLanguage.Russian)
                fieldDate.LocaleId = (int)EditingLanguage.Russian;

            doc.Save(ArtifactsDir + "FieldOptions.UpdateDateTimeFormatting.pdf");
        }

        /// <summary>
        /// Provides a CultureInfo object that should be used during the update of a field.
        /// </summary>
        private class FieldUpdateCultureProvider : IFieldUpdateCultureProvider
        {
            /// <summary>
            /// Returns a CultureInfo object to be used during the field's update.
            /// </summary>
            public CultureInfo GetCulture(string name, Field field)
            {
                switch (name)
                {
                    case "ru-RU":
                        CultureInfo culture = new CultureInfo(name, false);
                        DateTimeFormatInfo format = culture.DateTimeFormat;

                        format.MonthNames = new[] { "месяц 1", "месяц 2", "месяц 3", "месяц 4", "месяц 5", "месяц 6", "месяц 7", "месяц 8", "месяц 9", "месяц 10", "месяц 11", "месяц 12", "" };
                        format.MonthGenitiveNames = format.MonthNames;
                        format.AbbreviatedMonthNames = new[] { "мес 1", "мес 2", "мес 3", "мес 4", "мес 5", "мес 6", "мес 7", "мес 8", "мес 9", "мес 10", "мес 11", "мес 12", "" };
                        format.AbbreviatedMonthGenitiveNames = format.AbbreviatedMonthNames;

                        format.DayNames = new[] { "день недели 7", "день недели 1", "день недели 2", "день недели 3", "день недели 4", "день недели 5", "день недели 6" };
                        format.AbbreviatedDayNames = new[] { "день 7", "день 1", "день 2", "день 3", "день 4", "день 5", "день 6" };
                        format.ShortestDayNames = new[] { "д7", "д1", "д2", "д3", "д4", "д5", "д6" };

                        format.AMDesignator = "До полудня";
                        format.PMDesignator = "После полудня";

                        const string pattern = "yyyy MM (MMMM) dd (dddd) hh:mm:ss tt";
                        format.LongDatePattern = pattern;
                        format.LongTimePattern = pattern;
                        format.ShortDatePattern = pattern;
                        format.ShortTimePattern = pattern;

                        return culture;
                    case "en-US":
                        return new CultureInfo(name, false);
                    default:
                        return null;
                }
            }
        }
        //ExEnd

        [Test]
        public void BarcodeGenerator()
        {
            //ExStart
            //ExFor:BarcodeParameters
            //ExFor:BarcodeParameters.AddStartStopChar
            //ExFor:BarcodeParameters.BackgroundColor
            //ExFor:BarcodeParameters.BarcodeType
            //ExFor:BarcodeParameters.BarcodeValue
            //ExFor:BarcodeParameters.CaseCodeStyle
            //ExFor:BarcodeParameters.DisplayText
            //ExFor:BarcodeParameters.ErrorCorrectionLevel
            //ExFor:BarcodeParameters.FacingIdentificationMark
            //ExFor:BarcodeParameters.FixCheckDigit
            //ExFor:BarcodeParameters.ForegroundColor
            //ExFor:BarcodeParameters.IsBookmark
            //ExFor:BarcodeParameters.IsUSPostalAddress
            //ExFor:BarcodeParameters.PosCodeStyle
            //ExFor:BarcodeParameters.PostalAddress
            //ExFor:BarcodeParameters.ScalingFactor
            //ExFor:BarcodeParameters.SymbolHeight
            //ExFor:BarcodeParameters.SymbolRotation
            //ExFor:IBarcodeGenerator
            //ExFor:IBarcodeGenerator.GetBarcodeImage(BarcodeParameters)
            //ExFor:IBarcodeGenerator.GetOldBarcodeImage(BarcodeParameters)
            //ExFor:FieldOptions.BarcodeGenerator
            //ExSummary:Shows how to use a barcode generator.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            Assert.That(doc.FieldOptions.BarcodeGenerator, Is.Null); //ExSkip

            // We can use a custom IBarcodeGenerator implementation to generate barcodes,
            // and then insert them into the document as images.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Below are four examples of different barcode types that we can create using our generator.
            // For each barcode, we specify a new set of barcode parameters, and then generate the image.
            // Afterwards, we can insert the image into the document, or save it to the local file system.
            // 1 -  QR code:
            BarcodeParameters barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "QR",
                BarcodeValue = "ABC123",
                BackgroundColor = "0xF8BD69",
                ForegroundColor = "0xB5413B",
                ErrorCorrectionLevel = "3",
                ScalingFactor = "250",
                SymbolHeight = "1000",
                SymbolRotation = "0"
            };

            Image img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
#if NETFRAMEWORK || JAVA
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.QR.jpg");
#elif NET6_0_OR_GREATER
            using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "FieldOptions.BarcodeGenerator.QR.jpg"))
            {
                img.Encode(fs, SKEncodedImageFormat.Jpeg, 100);
            }
#endif
            builder.InsertImage(img);

            // 2 -  EAN13 barcode:
            barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "EAN13",
                BarcodeValue = "501234567890",
                DisplayText = true,
                PosCodeStyle = "CASE",
                FixCheckDigit = true
            };

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
#if NETFRAMEWORK || JAVA
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.EAN13.jpg");
#elif NET6_0_OR_GREATER
            using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "FieldOptions.BarcodeGenerator.EAN13.jpg"))
            {
                img.Encode(fs, SKEncodedImageFormat.Jpeg, 100);
            }
#endif
            builder.InsertImage(img);

            // 3 -  CODE39 barcode:
            barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "CODE39",
                BarcodeValue = "12345ABCDE",
                AddStartStopChar = true
            };

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
#if NETFRAMEWORK || JAVA
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.CODE39.jpg");
#elif NET6_0_OR_GREATER
            using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "FieldOptions.BarcodeGenerator.CODE39.jpg"))
            {
                img.Encode(fs, SKEncodedImageFormat.Jpeg, 100);
            }
#endif
            builder.InsertImage(img);

            // 4 -  ITF14 barcode:
            barcodeParameters = new BarcodeParameters
            {
                BarcodeType = "ITF14",
                BarcodeValue = "09312345678907",
                CaseCodeStyle = "STD"
            };

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
#if NETFRAMEWORK || JAVA
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.ITF14.jpg");
#elif NET6_0_OR_GREATER
            using (SKFileWStream fs = new SKFileWStream(ArtifactsDir + "FieldOptions.BarcodeGenerator.ITF14.jpg"))
            {
                img.Encode(fs, SKEncodedImageFormat.Jpeg, 100);
            }
#endif            
            builder.InsertImage(img);

            doc.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.docx");
            //ExEnd

            TestUtil.VerifyImage(223, 223, ArtifactsDir + "FieldOptions.BarcodeGenerator.QR.jpg");
            TestUtil.VerifyImage(126, 118, ArtifactsDir + "FieldOptions.BarcodeGenerator.EAN13.jpg");
            TestUtil.VerifyImage(204, 70, ArtifactsDir + "FieldOptions.BarcodeGenerator.CODE39.jpg");
            TestUtil.VerifyImage(168, 134, ArtifactsDir + "FieldOptions.BarcodeGenerator.ITF14.jpg");

            doc = new Document(ArtifactsDir + "FieldOptions.BarcodeGenerator.docx");
            Shape barcode = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.That(barcode.HasImage, Is.True);
        }
    }
}
