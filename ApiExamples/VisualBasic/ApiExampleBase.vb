' Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.IO
Imports System.Reflection

Imports Aspose.Words

Imports NUnit.Framework

Namespace ApiExamples
	''' <summary>
	''' Provides common infrastructure for all API examples that are implemented as unit tests.
	''' </summary>
	Public Class ApiExampleBase
		Private ReadOnly dirPath As String = MyDir & "\Artifacts\"

		<TestFixtureSetUp> _
		Public Sub SetUp()
			SetUnlimitedLicense()

			If (Not Directory.Exists(dirPath)) Then
				Directory.CreateDirectory(dirPath)
			End If
		End Sub

		<TestFixtureTearDown> _
		Public Sub TearDown()
			'Delete all files from dir
            For Each file In Directory.GetFiles(dirPath)
                IO.File.Delete(file)
            Next

			'Delete empty folder
			Directory.Delete(dirPath)
		End Sub

		Friend Shared Sub SetUnlimitedLicense()
			If File.Exists(TestLicenseFileName) Then
				' This shows how to use an Aspose.Words license when you have purchased one.
				' You don't have to specify full path as shown here. You can specify just the 
				' file name if you copy the license file into the same folder as your application
				' binaries or you add the license to your project as an embedded resource.
				Dim license As New License()
				license.SetLicense(TestLicenseFileName)
			End If
		End Sub

	   Friend Shared Sub RemoveLicense()
			Dim license As New License()
			license.SetLicense("")
	   End Sub

		''' <summary>
		''' Returns the assembly directory correctly even if the assembly is shadow-copied.
		''' </summary>
		Private Shared Function GetAssemblyDir(ByVal [assembly] As System.Reflection.Assembly) As String
			' CodeBase is a full URI, such as file:///x:\blahblah.
			Dim uri As New Uri([assembly].CodeBase)
			Return Path.GetDirectoryName(uri.LocalPath) + Path.DirectorySeparatorChar
		End Function

		''' <summary>
		''' Gets the path to the currently running executable.
		''' </summary>
		Friend Shared ReadOnly Property AssemblyDir() As String
			Get
				Return gAssemblyDir
			End Get
		End Property

		''' <summary>
		''' Gets the path to the documents used by the code examples. Ends with a back slash.
		''' </summary>
		Friend Shared ReadOnly Property MyDir() As String
			Get
				Return gMyDir
			End Get
		End Property

		''' <summary>
		''' Gets the path of the demo database. Ends with a back slash.
		''' </summary>
		Friend Shared ReadOnly Property DatabaseDir() As String
			Get
				Return gDatabaseDir
			End Get
		End Property

		Shared Sub New()
			gAssemblyDir = GetAssemblyDir(System.Reflection.Assembly.GetExecutingAssembly())
			gMyDir = New Uri(New Uri(gAssemblyDir), "../../../Data/").LocalPath
			gDatabaseDir = New Uri(New Uri(gAssemblyDir), "../../../Data/Database/").LocalPath
		End Sub

		Private Shared ReadOnly gAssemblyDir As String
		Private Shared ReadOnly gMyDir As String
		Private Shared ReadOnly gDatabaseDir As String

		''' <summary>
		''' This is where the test license is on my development machine.
		''' </summary>
		Friend Const TestLicenseFileName As String = "X:\awuex\Licenses\Aspose.Total.lic"
	End Class
End Namespace
