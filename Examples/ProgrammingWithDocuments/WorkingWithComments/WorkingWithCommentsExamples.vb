Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithComments
	<TestClass, TestFixture> _
	Public Class WorkingWithCommentsExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ProcessComments()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithComments/ProcessComments")
	
	            ProcessCommentsExample.Program.Main()
        End Sub

	End Class
End Namespace