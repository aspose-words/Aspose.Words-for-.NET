using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.RenderingAndPrinting.PrintDocument
{
    [TestClass, TestFixture]
    public class PrintDocumentExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void XpsPrint()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/PrintDocument/XpsPrint");

            XpsPrintExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void MultiplePagesOnSheet()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/PrintDocument/MultiplePagesOnSheet");

            MultiplePagesOnSheetExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void DocumentPreviewAndPrint()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/PrintDocument/DocumentPreviewAndPrint");

            DocumentPreviewAndPrintExample.Program.Main();
        }

    }
}
