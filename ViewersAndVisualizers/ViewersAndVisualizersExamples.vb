Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.ViewersAndVisualizers
	<TestClass, TestFixture> _
	Public Class ViewersAndVisualizersExamples
        <TestMethod(), Test(), Owner("WinForm")> _
	        Public Sub DocumentExplorer()
	            TestHelper.SetDataDir("ViewersAndVisualizers/DocumentExplorer")
	
	            DocumentExplorerExample.Program.Main()
        End Sub

	End Class
End Namespace