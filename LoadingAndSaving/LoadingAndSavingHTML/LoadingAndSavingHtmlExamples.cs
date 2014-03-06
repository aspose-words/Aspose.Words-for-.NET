using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.LoadingAndSaving.LoadingAndSavingHtml
{
    [TestClass, TestFixture]
    public class LoadingAndSavingHtmlExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void Word2Help()
        {
            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/Word2Help");

            Word2HelpExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void SplitIntoHtmlPages()
        {
            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/SplitIntoHtmlPages");

            SplitIntoHtmlPagesExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void SaveMhtmlAndEmail()
        {
            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/SaveMhtmlAndEmail");

            SaveMhtmlAndEmailExample.Program.Main();
        }

        [TestMethod, Test, Owner("WinForm")]
        public void SaveHtmlAndEmail()
        {
            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/SaveHtmlAndEmail");

            SaveHtmlAndEmailExample.MainForm.Main();
        }

    }
}
