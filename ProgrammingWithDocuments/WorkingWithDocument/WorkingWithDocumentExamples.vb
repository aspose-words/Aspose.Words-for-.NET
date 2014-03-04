Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithDocument
	<TestClass, TestFixture> _
	Public Class WorkingWithDocumentExamples
        <TestMethod(), Test(), Owner("WinForm")> _
	        Public Sub DocumentInDB()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/DocumentInDB")
	
	            DocumentInDBExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub RemoveBreaks()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/RemoveBreaks")
	
	            RemoveBreaksExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub PageNumbersOfNodes()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/PageNumbersOfNodes")
	
	            PageNumbersOfNodesExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ExtractContent()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithDocument/ExtractContent")
	
	            ExtractContentExample.Program.Main()
        End Sub

	End Class
End Namespace