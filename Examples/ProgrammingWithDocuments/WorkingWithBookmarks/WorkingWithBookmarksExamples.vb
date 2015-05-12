Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithBookmarks
	<TestClass, TestFixture> _
	Public Class WorkingWithBookmarksExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub CopyBookmarkedText()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithBookmarks/CopyBookmarkedText")
	
	            CopyBookmarkedTextExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub UntangleRowBookmarks()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithBookmarks/UntangleRowBookmarks")
	
	            UntangleRowBookmarksExample.Program.Main()
        End Sub

	End Class
End Namespace