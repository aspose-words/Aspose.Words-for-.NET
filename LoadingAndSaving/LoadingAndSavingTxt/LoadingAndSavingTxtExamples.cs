using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.LoadingAndSaving.LoadingAndSavingTxt
{
    [TestClass, TestFixture]
    public class LoadingAndSavingTxtExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void LoadTxt()
        {
            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingTxt/LoadTxt");

            LoadTxtExample.Program.Main();
        }

    }
}
