Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.UsingFindAndReplace
	<TestClass, TestFixture> _
	Public Class UsingFindAndReplaceExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub FindAndHighlight()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/UsingFindAndReplace/FindAndHighlight")
	
	            FindAndHighlightExample.Program.Main()
        End Sub

	End Class
End Namespace