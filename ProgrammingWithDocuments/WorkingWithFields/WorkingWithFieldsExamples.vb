Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithFields
	<TestClass, TestFixture> _
	Public Class WorkingWithFieldsExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ReplaceFieldsWithStaticText()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithFields/ReplaceFieldsWithStaticText")
	
	            ReplaceFieldsWithStaticTextExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub InsertNestedFields()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithFields/InsertNestedFields")
	
	            InsertNestedFieldsExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub RemoveField()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithFields/RemoveField")
	
	            RemoveFieldExample.Program.Main()
        End Sub

	End Class
End Namespace