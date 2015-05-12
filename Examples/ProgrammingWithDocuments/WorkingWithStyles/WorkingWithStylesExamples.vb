Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ProgrammingWithDocuments.WorkingWithStyles
	<TestClass, TestFixture> _
	Public Class WorkingWithStylesExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub ExtractContentBasedOnStyles()
	            TestHelper.SetDataDir("ProgrammingWithDocuments/WorkingWithStyles/ExtractContentBasedOnStyles")
	
	            ExtractContentBasedOnStylesExample.Program.Main()
        End Sub

	End Class
End Namespace