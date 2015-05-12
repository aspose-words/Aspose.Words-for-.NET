using Helpers;
using NUnit.Framework;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

[assembly: AssemblyTitle("Aspose.Words Examples")]
[assembly: AssemblyDescription("A collection of examples which demonstrate how to use the Aspose.Words for .NET API.")]
[assembly: AssemblyConfiguration("CSharp")]

namespace Examples.QuickStart
{
    [TestClass, TestFixture]
    public class QuickStartExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void HelloWorld()
        {
            TestHelper.SetDataDir("QuickStart/HelloWorld");

            HelloWorldExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void AppendDocuments()
        {
            TestHelper.SetDataDir("QuickStart/AppendDocuments");

            AppendDocumentsExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void ApplyLicense()
        {
            TestHelper.SetDataDir("QuickStart/ApplyLicense");

            ApplyLicenseExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void FindAndReplace()
        {
            TestHelper.SetDataDir("QuickStart/FindAndReplace");

            FindAndReplaceExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void SimpleMailMerge()
        {
            TestHelper.SetDataDir("QuickStart/SimpleMailMerge");

            SimpleMailMergeExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void LoadAndSaveToDisk()
        {
            TestHelper.SetDataDir("QuickStart/LoadAndSaveToDisk");

            LoadAndSaveToDiskExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void LoadAndSaveToStream()
        {
            TestHelper.SetDataDir("QuickStart/LoadAndSaveToStream");

            LoadAndSaveToStreamExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void Doc2Pdf()
        {
            TestHelper.SetDataDir("QuickStart/Doc2Pdf");

            Doc2PdfExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void UpdateFields()
        {
            TestHelper.SetDataDir("QuickStart/UpdateFields");

            UpdateFieldsExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void WorkingWithNodes()
        {
            TestHelper.SetDataDir("QuickStart/WorkingWithNodes");

            WorkingWithNodesExample.Program.Main();
        }

    }

    [TestClass, SetUpFixture]
    public class AsposeExamples
    {
        [AssemblyInitialize]
        public static void AssemblyInitialize(Microsoft.VisualStudio.TestTools.UnitTesting.TestContext context)
        {
            Main();
        }

        [SetUp]
        public static void AssemblySetup()
        {
            Main();
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            TestHelper.Cleanup();
        }

        public static void Main()
        {
            // Provides an introduction of how to use this example project.
            TestHelper.ShowIntroForm();
        }
    }
}
