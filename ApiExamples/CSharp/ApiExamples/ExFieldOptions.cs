// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExFieldOptions : ApiExampleBase
    {
        [Test]
        public void FieldOptionsCurrentUser()
        {
            //ExStart
            //ExFor:Document.UpdateFields
            //ExFor:FieldOptions.CurrentUser
            //ExFor:UserInformation
            //ExFor:UserInformation.Name
            //ExFor:UserInformation.Initials
            //ExFor:UserInformation.Address
            //ExFor:UserInformation.DefaultUser
            //ExSummary:Shows how to set user details and display them with fields.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set user information
            UserInformation userInformation = new UserInformation();
            userInformation.Name = "John Doe";
            userInformation.Initials = "J. D.";
            userInformation.Address = "123 Main Street";
            doc.FieldOptions.CurrentUser = userInformation;

            // Insert fields that reference our user information
            Assert.AreEqual(userInformation.Name, builder.InsertField(" USERNAME ").Result);
            Assert.AreEqual(userInformation.Initials, builder.InsertField(" USERINITIALS ").Result);
            Assert.AreEqual(userInformation.Address, builder.InsertField(" USERADDRESS ").Result);

            // The field options object also has a static default user value that fields from many documents can refer to
            UserInformation.DefaultUser.Name = "Default User";
            UserInformation.DefaultUser.Initials = "D. U.";
            UserInformation.DefaultUser.Address = "One Microsoft Way";
            doc.FieldOptions.CurrentUser = UserInformation.DefaultUser;

            Assert.AreEqual("Default User", builder.InsertField(" USERNAME ").Result);
            Assert.AreEqual("D. U.", builder.InsertField(" USERINITIALS ").Result);
            Assert.AreEqual("One Microsoft Way", builder.InsertField(" USERADDRESS ").Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FieldOptions.FieldOptionsCurrentUser.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.FieldOptionsCurrentUser.docx");

            Assert.Null(doc.FieldOptions.CurrentUser);

            FieldUserName fieldUserName = (FieldUserName)doc.Range.Fields[0];

            Assert.Null(fieldUserName.UserName);
            Assert.AreEqual("Default User", fieldUserName.Result);

            FieldUserInitials fieldUserInitials = (FieldUserInitials)doc.Range.Fields[1];

            Assert.Null(fieldUserInitials.UserInitials);
            Assert.AreEqual("D. U.", fieldUserInitials.Result);

            FieldUserAddress fieldUserAddress = (FieldUserAddress)doc.Range.Fields[2];

            Assert.Null(fieldUserAddress.UserAddress);
            Assert.AreEqual("One Microsoft Way", fieldUserAddress.Result);
        }

        [Test]
        public void FieldOptionsFileName()
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

            // This FILENAME field will display the file name of the document we opened
            FieldFileName field = (FieldFileName)builder.InsertField(FieldType.FieldFileName, true);
            field.Update();

            Assert.AreEqual(" FILENAME ", field.GetFieldCode());
            Assert.AreEqual("Document.docx", field.Result);

            builder.Writeln();

            // By default, the FILENAME field does not show the full path, and we can change this
            field = (FieldFileName)builder.InsertField(FieldType.FieldFileName, true);
            field.IncludeFullPath = true;

            // We can override the values displayed by our FILENAME fields by setting this attribute
            doc.FieldOptions.FileName = "FieldOptions.FILENAME.docx";
            field.Update();

            Assert.AreEqual(" FILENAME  \\p", field.GetFieldCode());
            Assert.AreEqual("FieldOptions.FILENAME.docx", field.Result);

            doc.UpdateFields();
            doc.Save(ArtifactsDir + doc.FieldOptions.FileName);
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.FILENAME.docx");

            Assert.IsNull(doc.FieldOptions.FileName);
            TestUtil.VerifyField(FieldType.FieldFileName, " FILENAME ", "FieldOptions.FILENAME.docx", doc.Range.Fields[0]);
        }

        [Test]
        public void FieldOptionsBidi()
        {
            //ExStart
            //ExFor:FieldOptions.IsBidiTextSupportedOnUpdate
            //ExSummary:Shows how to use FieldOptions to ensure that bi-directional text is properly supported during the field update.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Ensure that any field operation involving right-to-left text is performed correctly 
            doc.FieldOptions.IsBidiTextSupportedOnUpdate = true;

            // Use a document builder to insert a field which contains right-to-left text
            FormField comboBox = builder.InsertComboBox("MyComboBox", new[] { "עֶשְׂרִים", "שְׁלוֹשִׁים", "אַרְבָּעִים", "חֲמִשִּׁים", "שִׁשִּׁים" }, 0);
            comboBox.CalculateOnExit = true;

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FieldOptions.FieldOptionsBidi.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.FieldOptionsBidi.docx");

            Assert.False(doc.FieldOptions.IsBidiTextSupportedOnUpdate);

            comboBox = doc.Range.FormFields[0];

            Assert.AreEqual("עֶשְׂרִים", comboBox.Result);
        }

        [Test]
        public void FieldOptionsLegacyNumberFormat()
        {
            //ExStart
            //ExFor:FieldOptions.LegacyNumberFormat
            //ExSummary:Shows how use FieldOptions to change the number format.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Field field = builder.InsertField("= 2 + 3 \\# $##");

            Assert.AreEqual("$ 5", field.Result);

            doc.FieldOptions.LegacyNumberFormat = true;
            field.Update();

            Assert.AreEqual("$5", field.Result);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.False(doc.FieldOptions.LegacyNumberFormat);
            TestUtil.VerifyField(FieldType.FieldFormula, "= 2 + 3 \\# $##", "$5", doc.Range.Fields[0]);
        }

        [Test]
        public void FieldOptionsPreProcessCulture()
        {
            //ExStart
            //ExFor:FieldOptions.PreProcessCulture
            //ExSummary:Shows how to set the preprocess culture.
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.FieldOptions.PreProcessCulture = new CultureInfo("de-DE");

            Field field = builder.InsertField(" DOCPROPERTY CreateTime");

            // Conforming to the German culture, the date/time will be presented in the "dd.mm.yyyy hh:mm" format
            Assert.IsTrue(Regex.Match(field.Result, @"\d{2}[.]\d{2}[.]\d{4} \d{2}[:]\d{2}").Success);

            doc.FieldOptions.PreProcessCulture = CultureInfo.InvariantCulture;
            field.Update();

            // After switching to the invariant culture, the date/time will be presented in the "mm/dd/yyyy hh:mm" format
            Assert.IsTrue(Regex.Match(field.Result, @"\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}").Success);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.Null(doc.FieldOptions.PreProcessCulture);
            Assert.IsTrue(Regex.Match(doc.Range.Fields[0].Result, @"\d{2}[/]\d{2}[/]\d{4} \d{2}[:]\d{2}").Success);
        }

        [Test]
        public void FieldOptionsToaCategories()
        {
            //ExStart
            //ExFor:FieldOptions.ToaCategories
            //ExFor:ToaCategories
            //ExFor:ToaCategories.Item(Int32)
            //ExFor:ToaCategories.DefaultCategories
            //ExSummary:Shows how to specify a table of authorities categories for a document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // There are default category values we can use, or we can make our own like this
            ToaCategories toaCategories = new ToaCategories();
            doc.FieldOptions.ToaCategories = toaCategories;

            toaCategories[1] = "My Category 1"; // Replaces default value "Cases"
            toaCategories[2] = "My Category 2"; // Replaces default value "Statutes"

            // Even if we changed the categories in the FieldOptions object, the default categories are still available here
            Assert.AreEqual("Cases", ToaCategories.DefaultCategories[1]);
            Assert.AreEqual("Statutes", ToaCategories.DefaultCategories[2]);

            // Insert 2 tables of authorities, one per category
            builder.InsertField("TOA \\c 1 \\h", null);
            builder.InsertField("TOA \\c 2 \\h", null);
            builder.InsertBreak(BreakType.PageBreak);

            // Insert TOA entries across 2 categories
            builder.InsertField("TA \\c 2 \\l \"entry 1\"");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertField("TA \\c 1 \\l \"entry 2\"");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertField("TA \\c 2 \\l \"entry 3\"");

            doc.UpdateFields();
            doc.Save(ArtifactsDir + "FieldOptions.TOA.Categories.docx");
            //ExEnd

            doc = new Document(ArtifactsDir + "FieldOptions.TOA.Categories.docx");

            Assert.Null(doc.FieldOptions.ToaCategories);

            TestUtil.VerifyField(FieldType.FieldTOA, "TOA \\c 1 \\h", "My Category 1\rentry 2\t3\r", doc.Range.Fields[0]);
            TestUtil.VerifyField(FieldType.FieldTOA, "TOA \\c 2 \\h",
                "My Category 2\r" +
                "entry 1\t2\r" +
                "entry 3\t4\r", doc.Range.Fields[1]);
        }

        [Test]
        public void FieldOptionsUseInvariantCultureNumberFormat()
        {
            //ExStart
            //ExFor:FieldOptions.UseInvariantCultureNumberFormat
            //ExSummary:Shows how to format numbers according to the invariant culture.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            Field field = builder.InsertField(" = 1234567,89 \\# $#,###,###.##");
            field.Update();

            // The combination of field, number format and thread culture can sometimes produce an unsuitable result
            Assert.IsFalse(doc.FieldOptions.UseInvariantCultureNumberFormat);
            Assert.AreEqual("$1234567,89 .     ", field.Result);

            // We can set this attribute to avoid changing the whole thread culture just for numeric formats
            doc.FieldOptions.UseInvariantCultureNumberFormat = true;
            field.Update();
            Assert.AreEqual("$1.234.567,89", field.Result);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);

            Assert.False(doc.FieldOptions.UseInvariantCultureNumberFormat);
            TestUtil.VerifyField(FieldType.FieldFormula, " = 1234567,89 \\# $#,###,###.##", "$1.234.567,89", doc.Range.Fields[0]);
        }

        //ExStart
        //ExFor:FieldOptions.FieldUpdateCultureProvider
        //ExFor:IFieldUpdateCultureProvider
        //ExFor:IFieldUpdateCultureProvider.GetCulture(string, Field)
        //ExSummary:Shows how to specify a culture defining date/time formatting on per field basis.
        [Test]
        public void DefineDateTimeFormatting()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertField(FieldType.FieldTime, true);

            doc.FieldOptions.FieldUpdateCultureSource = FieldUpdateCultureSource.FieldCode;

            // Set a provider that returns a culture object specific for each field
            doc.FieldOptions.FieldUpdateCultureProvider = new FieldUpdateCultureProvider();

            FieldTime fieldDate = (FieldTime)doc.Range.Fields[0];
            if (fieldDate.LocaleId != (int)EditingLanguage.Russian)
                fieldDate.LocaleId = (int)EditingLanguage.Russian;

            doc.Save(ArtifactsDir + "FieldOptions.UpdateDateTimeFormatting.pdf");
        }

        /// <summary>
        /// Provides a CultureInfo object that should be used during the update of a particular field.
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

#if NET462 || JAVA
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
            Assert.IsNull(doc.FieldOptions.BarcodeGenerator); //ExSkip

            // Barcodes generated in this way will be in the form of images.
            // We can use a custom IBarcodeGenerator implementation to generate them.
            // This generator can be found here:
            // https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/ApiExamples/CSharp/ApiExamples/CustomBarcodeGenerator.cs
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Below are four of the types of barcodes that we can insert this way.
            // 1 -  QR code:
            BarcodeParameters barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "QR";
            barcodeParameters.BarcodeValue = "ABC123";
            barcodeParameters.BackgroundColor = "0xF8BD69";
            barcodeParameters.ForegroundColor = "0xB5413B";
            barcodeParameters.ErrorCorrectionLevel = "3";
            barcodeParameters.ScalingFactor = "250";
            barcodeParameters.SymbolHeight = "1000";
            barcodeParameters.SymbolRotation = "0";

            // We can generate the barcode after we have set the parameters for the barcode generator.
            // The barcode will be an image that we can insert into the document and save to the local file system.
            Image img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.QR.jpg");

            builder.InsertImage(img);

            // 2 -  EAN13 barcode:
            barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "EAN13";
            barcodeParameters.BarcodeValue = "501234567890";
            barcodeParameters.DisplayText = true;
            barcodeParameters.PosCodeStyle = "CASE";
            barcodeParameters.FixCheckDigit = true;

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.EAN13.jpg");
            builder.InsertImage(img);

            // 3 -  CODE39 barcode:
            barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "CODE39";
            barcodeParameters.BarcodeValue = "12345ABCDE";
            barcodeParameters.AddStartStopChar = true;

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.CODE39.jpg");
            builder.InsertImage(img);

            // 4 -  ITF14 barcode:
            barcodeParameters = new BarcodeParameters();
            barcodeParameters.BarcodeType = "ITF14";
            barcodeParameters.BarcodeValue = "09312345678907";
            barcodeParameters.CaseCodeStyle = "STD";

            img = doc.FieldOptions.BarcodeGenerator.GetBarcodeImage(barcodeParameters);
            img.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.ITF14.jpg");
            builder.InsertImage(img);

            doc.Save(ArtifactsDir + "FieldOptions.BarcodeGenerator.docx");
            //ExEnd

            TestUtil.VerifyImage(378, 378, ArtifactsDir + "FieldOptions.BarcodeGenerator.QR.jpg");
            TestUtil.VerifyImage(220, 78, ArtifactsDir + "FieldOptions.BarcodeGenerator.EAN13.jpg");
            TestUtil.VerifyImage(414, 65, ArtifactsDir + "FieldOptions.BarcodeGenerator.CODE39.jpg");
            TestUtil.VerifyImage(300, 65, ArtifactsDir + "FieldOptions.BarcodeGenerator.ITF14.jpg");

            doc = new Document(ArtifactsDir + "FieldOptions.BarcodeGenerator.docx");
            Shape barcode = (Shape)doc.GetChild(NodeType.Shape, 0, true);

            Assert.True(barcode.HasImage);

            TestUtil.VerifyWebResponseStatusCode(HttpStatusCode.OK,
                "https://github.com/aspose-words/Aspose.Words-for-.NET/blob/master/ApiExamples/CSharp/ApiExamples/CustomBarcodeGenerator.cs");
        }
#endif
    }
}
