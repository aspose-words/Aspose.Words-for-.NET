using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;

namespace QaTests.Tests
{
    [TestFixture]
    internal class QaDocumentBuilder : QaTestsBase
    {
        [Ignore]
        [Test]
        // Bug "trimmed name if you enter more than 20 characters"
        public void InsertCheckBox()
        {
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Insert checkboxes
            builder.InsertCheckBox("CheckBox_DefaultAndCheckedValue", false, true, 0);
            builder.InsertCheckBox("CheckBox_OnlyCheckedValue", true, 100);
            
            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            //Get checkboxes from the document
            FormFieldCollection formFields = doc.Range.FormFields;

            //Check that is the right checkbox
            Assert.AreEqual("CheckBox_DefaultAndCheckedValue", formFields[0].Name);

            //Assert that parameters sets correctly
            Assert.AreEqual(true, formFields[0].Checked);
            Assert.AreEqual(false, formFields[0].Default);
            Assert.AreEqual(10, formFields[0].CheckBoxSize);

            //Check that is the right checkbox
            Assert.AreEqual("CheckBox_OnlyCheckedValue", formFields[1].Name);

            //Assert that parameters sets correctly
            Assert.AreEqual(false, formFields[1].Checked);
            Assert.AreEqual(false, formFields[1].Default);
            Assert.AreEqual(100, formFields[1].CheckBoxSize);
        }

        [Test]
        public void InsertCheckBox_EmptyName()
        {
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Assert that empty string name working correctly
            builder.InsertCheckBox("", true, false, 1);
            builder.InsertCheckBox(string.Empty, false, 1);
        }
    }
}
