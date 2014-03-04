Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.JoiningAndAppending
	<TestClass, TestFixture> _
	Public Class JoiningAndAppendingExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub AppendDocument()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/JoiningAndAppending/AppendDocument")
	
	            AppendDocumentExample.Program.Main()
        End Sub

	End Class
End Namespace