using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.RenderingAndPrinting
{
    [TestClass, TestFixture]
    public class RenderingAndPrintingExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void EnumerateLayoutElements()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/EnumerateLayoutElements");

            EnumerateLayoutElementsExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void DocumentLayoutHelper()
        {
            TestHelper.SetDataDir("RenderingAndPrinting/DocumentLayoutHelper");

            DocumentLayoutHelperExample.Program.Main();
        }

    }
}
