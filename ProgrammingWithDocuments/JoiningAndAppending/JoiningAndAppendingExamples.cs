using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.JoiningAndAppending
{
    [TestClass, TestFixture]
    public class JoiningAndAppendingExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void AppendDocument()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/JoiningAndAppending/AppendDocument");

            AppendDocumentExample.Program.Main();
        }

    }
}
