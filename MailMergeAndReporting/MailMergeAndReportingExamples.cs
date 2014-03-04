using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.MailMergeAndReporting
{
    [TestClass, TestFixture]
    public class MailMergeAndReportingExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void ApplyCustomLogicToEmptyRegions()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/ApplyCustomLogicToEmptyRegions");

            ApplyCustomLogicToEmptyRegionsExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void XMLMailMerge()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/XMLMailMerge");

            XMLMailMergeExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void RemoveEmptyRegions()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/RemoveEmptyRegions");

            RemoveEmptyRegionsExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void NestedMailMerge()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/NestedMailMerge");

            NestedMailMergeExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void MultipleDocsInMailMerge()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/MultipleDocsInMailMerge");

            MultipleDocsInMailMergeExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void MailMergeFormFields()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/MailMergeFormFields");

            MailMergeFormFieldsExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void LINQtoXMLMailMerge()
        {
            TestHelper.SetDataDir("MailMergeAndReporting/LINQtoXMLMailMerge");

            LINQtoXMLMailMergeExample.Program.Main();
        }

    }
}
