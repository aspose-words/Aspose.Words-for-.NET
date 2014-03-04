Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithTables
	<TestClass, TestFixture> _
	Public Class WorkingWithTablesExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub AutoFitTables()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithTables/AutoFitTables")
	
	            AutoFitTablesExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ImportTableFromDataTable()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithTables/ImportTableFromDataTable")
	
	            ImportTableFromDataTableExample.Program.Main()
        End Sub

	End Class
End Namespace