using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.WorkingWithStyles
{
    [TestClass, TestFixture]
    public class WorkingWithStylesExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void ExtractContentBasedOnStyles()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithStyles/ExtractContentBasedOnStyles");

            ExtractContentBasedOnStylesExample.Program.Main();
        }

    }
}
