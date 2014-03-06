Imports Helpers
Imports NUnit.Framework
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Assert = NUnit.Framework.Assert
Imports Description = Microsoft.VisualStudio.TestTools.UnitTesting.DescriptionAttribute

Namespace Examples.LoadingAndSaving.LoadingAndSavingHtml
	<TestClass, TestFixture> _
	Public Class LoadingAndSavingHtmlExamples
        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub Word2Help()
	            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/Word2Help")
	
	            Word2HelpExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub SplitIntoHtmlPages()
	            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/SplitIntoHtmlPages")
	
	            SplitIntoHtmlPagesExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("Console")> _
	        Public Sub SaveMhtmlAndEmail()
	            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/SaveMhtmlAndEmail")
	
	            SaveMhtmlAndEmailExample.Program.Main()
        End Sub

        <TestMethod(), Test(), Owner("WinForm")> _
	        Public Sub SaveHtmlAndEmail()
	            TestHelper.SetDataDir("LoadingAndSaving/LoadingAndSavingHtml/SaveHtmlAndEmail")
	
	            SaveHtmlAndEmailExample.MainForm.Main()
        End Sub

	End Class
End Namespace