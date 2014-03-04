using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.UsingFindAndReplace
{
    [TestClass, TestFixture]
    public class UsingFindAndReplaceExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void FindAndHighlight()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/UsingFindAndReplace/FindAndHighlight");

            FindAndHighlightExample.Program.Main();
        }

    }
}
