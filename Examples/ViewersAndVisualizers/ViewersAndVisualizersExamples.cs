using Helpers;
using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Assert = NUnit.Framework.Assert;
using Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute;

namespace Examples.ViewersAndVisualizers
{
    [TestClass, TestFixture]
    public class ViewersAndVisualizersExamples
    {	
        [TestMethod, Test, Owner("WinForm")]
        public void DocumentExplorer()
        {
            TestHelper.SetDataDir("ViewersAndVisualizers/DocumentExplorer");

            DocumentExplorerExample.MainForm.Main();
        }

    }
}
