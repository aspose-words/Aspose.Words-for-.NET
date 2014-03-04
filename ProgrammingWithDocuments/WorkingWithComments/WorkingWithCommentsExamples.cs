using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.WorkingWithComments
{
    [TestClass, TestFixture]
    public class WorkingWithCommentsExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void ProcessComments()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithComments/ProcessComments");

            ProcessCommentsExample.Program.Main();
        }

    }
}
