using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.WorkingWithDocument
{
    [TestClass, TestFixture]
    public class WorkingWithDocumentExamples
    {	
        [TestMethod, Test, Owner("WinForm")]
        public void DocumentInDB()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/DocumentInDB");

            DocumentInDBExample.MainForm.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void RemoveBreaks()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/RemoveBreaks");

            RemoveBreaksExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void PageNumbersOfNodes()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/PageNumbersOfNodes");

            PageNumbersOfNodesExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void ExtractContent()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/ExtractContent");

            ExtractContentExample.Program.Main();
        }

    }
}
