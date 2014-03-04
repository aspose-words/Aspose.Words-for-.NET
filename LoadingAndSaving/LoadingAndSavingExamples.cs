using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.LoadingAndSaving
{
    [TestClass, TestFixture]
    public class LoadingAndSavingExamples
    {	
        [TestMethod, Test, Owner("WinForm")]
        public void Excel2Word()
        {
            TestHelper.SetDataDir("LoadingAndSaving/Excel2Word");

            Excel2WordExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void CheckFormat()
        {
            TestHelper.SetDataDir("LoadingAndSaving/CheckFormat");

            CheckFormatExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void PageSplitter()
        {
            TestHelper.SetDataDir("LoadingAndSaving/PageSplitter");

            PageSplitterExample.Program.Main();
        }

    }
}
