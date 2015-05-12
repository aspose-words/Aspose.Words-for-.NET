using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.WorkingWithBookmarks
{
    [TestClass, TestFixture]
    public class WorkingWithBookmarksExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void CopyBookmarkedText()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithBookmarks/CopyBookmarkedText");

            CopyBookmarkedTextExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void UntangleRowBookmarks()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithBookmarks/UntangleRowBookmarks");

            UntangleRowBookmarksExample.Program.Main();
        }

    }
}
