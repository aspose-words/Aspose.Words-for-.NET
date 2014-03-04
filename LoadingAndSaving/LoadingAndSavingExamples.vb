Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.LoadingAndSaving
	<TestClass, TestFixture> _
	Public Class LoadingAndSavingExamples
        <TestMethod(), Test(), Owner("WinForm")> _
	        Public Sub Excel2Word()
	            TestHelper.SetDataDir("LoadingAndSaving/Excel2Word")
	
	            Excel2WordExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub CheckFormat()
	            TestHelper.SetDataDir("LoadingAndSaving/CheckFormat")
	
	            CheckFormatExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub PageSplitter()
	            TestHelper.SetDataDir("LoadingAndSaving/PageSplitter")
	
	            PageSplitterExample.Program.Main()
        End Sub

	End Class
End Namespace