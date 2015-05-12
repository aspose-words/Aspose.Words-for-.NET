using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.LoadingAndSaving.SavingToPdfFormat
{
    [TestClass, TestFixture]
    public class SavingToPdfFormatExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void ImageToPdf()
        {
            TestHelper.SetDataDir("LoadingAndSaving/SavingToPdfFormat/ImageToPdf");

            ImageToPdfExample.Program.Main();
        }

    }
}
