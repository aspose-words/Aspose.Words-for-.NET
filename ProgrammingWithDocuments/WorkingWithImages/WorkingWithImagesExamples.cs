using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.WorkingWithImages
{
    [TestClass, TestFixture]
    public class WorkingWithImagesExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void AddImageToEachPage()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithImages/AddImageToEachPage");

            AddImageToEachPageExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void AddWatermark()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithImages/AddWatermark");

            AddWatermarkExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void CompressImages()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithImages/CompressImages");

            CompressImagesExample.Program.Main();
        }

    }
}
