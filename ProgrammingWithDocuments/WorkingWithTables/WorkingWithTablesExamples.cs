using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ProgrammingWithDocuments.WorkingWithTables
{
    [TestClass, TestFixture]
    public class WorkingWithTablesExamples
    {	
        [TestMethod, Test, Owner("Console")]
        public void AutoFitTables()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithTables/AutoFitTables");

            AutoFitTablesExample.Program.Main();
        }

        [TestMethod, Test, Owner("Console")]
        public void ImportTableFromDataTable()
        {
            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithTables/ImportTableFromDataTable");

            ImportTableFromDataTableExample.Program.Main();
        }

    }
}
