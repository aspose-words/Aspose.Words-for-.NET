'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Runtime.InteropServices

Imports Aspose.Words

Namespace XpsPrintExample
	''' <summary>
	''' A utility class that converts a document to XPS using Aspose.Words and then sends to the XpsPrint API.
	''' </summary>
	Public Class XpsPrintHelper
		''' <summary>
		''' No ctor.
		''' </summary>
		Private Sub New()
		End Sub

'ExStart
'ExId:XpsPrint_PrintDocument
'ExSummary:Convert an Aspose.Words document into an XPS stream and print.
		''' <summary>
		''' Sends an Aspose.Words document to a printer using the XpsPrint API.
		''' </summary>
		''' <param name="document"></param>
		''' <param name="printerName"></param>
		''' <param name="jobName">Job name. Can be null.</param>
		''' <param name="isWait">True to wait for the job to complete. False to return immediately after submitting the job.</param>
		''' <exception cref="Exception">Thrown if any error occurs.</exception>
		Public Shared Sub Print(ByVal document As Aspose.Words.Document, ByVal printerName As String, ByVal jobName As String, ByVal isWait As Boolean)
			If document Is Nothing Then
				Throw New ArgumentNullException("document")
			End If

			' Use Aspose.Words to convert the document to XPS and store in a memory stream.
			Dim stream As New MemoryStream()
			document.Save(stream, SaveFormat.Xps)
			stream.Position = 0

			Print(stream, printerName, jobName, isWait)
		End Sub
'ExEnd

'ExStart
'ExId:XpsPrint_PrintStream
'ExSummary:Prints an XPS document using the XpsPrint API.
		''' <summary>
		''' Sends a stream that contains a document in the XPS format to a printer using the XpsPrint API.
		''' Has no dependency on Aspose.Words, can be used in any project.
		''' </summary>
		''' <param name="stream"></param>
		''' <param name="printerName"></param>
		''' <param name="jobName">Job name. Can be null.</param>
		''' <param name="isWait">True to wait for the job to complete. False to return immediately after submitting the job.</param>
		''' <exception cref="Exception">Thrown if any error occurs.</exception>
		Public Shared Sub Print(ByVal stream As Stream, ByVal printerName As String, ByVal jobName As String, ByVal isWait As Boolean)
			If stream Is Nothing Then
				Throw New ArgumentNullException("stream")
			End If
			If printerName Is Nothing Then
				Throw New ArgumentNullException("printerName")
			End If

			' Create an event that we will wait on until the job is complete.
			Dim completionEvent As IntPtr = CreateEvent(IntPtr.Zero, True, False, Nothing)
			If completionEvent = IntPtr.Zero Then
				Throw New Win32Exception()
			End If

			Try
				Dim job As IXpsPrintJob
				Dim jobStream As IXpsPrintJobStream
				StartJob(printerName, jobName, completionEvent, job, jobStream)

				CopyJob(stream, job, jobStream)

				If isWait Then
					WaitForJob(completionEvent)
					CheckJobStatus(job)
				End If
			Finally
				If completionEvent <> IntPtr.Zero Then
					CloseHandle(completionEvent)
				End If
			End Try
		End Sub
'ExEnd

		Private Shared Sub StartJob(ByVal printerName As String, ByVal jobName As String, ByVal completionEvent As IntPtr, <System.Runtime.InteropServices.Out()> ByRef job As IXpsPrintJob, <System.Runtime.InteropServices.Out()> ByRef jobStream As IXpsPrintJobStream)
			Dim result As Integer = StartXpsPrintJob(printerName, jobName, Nothing, IntPtr.Zero, completionEvent, Nothing, 0, job, jobStream, IntPtr.Zero)
			If result <> 0 Then
				Throw New Win32Exception(result)
			End If
		End Sub

		Private Shared Sub CopyJob(ByVal stream As Stream, ByVal job As IXpsPrintJob, ByVal jobStream As IXpsPrintJobStream)
			Try
				Dim buff(4095) As Byte
				Do
					Dim read As UInteger = CUInt(stream.Read(buff, 0, buff.Length))
					If read = 0 Then
						Exit Do
					End If

					Dim written As UInteger
					jobStream.Write(buff, read, written)

					If read <> written Then
						Throw New Exception("Failed to copy data to the print job stream.")
					End If
				Loop

				' Indicate that the entire document has been copied.
				jobStream.Close()
			Catch e1 As Exception
				' Cancel the job if we had any trouble submitting it.
				job.Cancel()
				Throw
			End Try
		End Sub

		Private Shared Sub WaitForJob(ByVal completionEvent As IntPtr)
			Const INFINITE As Integer = -1
			Select Case WaitForSingleObject(completionEvent, INFINITE)
				Case WAIT_RESULT.WAIT_OBJECT_0
					' Expected result, do nothing.
				Case WAIT_RESULT.WAIT_FAILED
					Throw New Win32Exception()
				Case Else
					Throw New Exception("Unexpected result when waiting for the print job.")
			End Select
		End Sub

		Private Shared Sub CheckJobStatus(ByVal job As IXpsPrintJob)
			Dim jobStatus As XPS_JOB_STATUS
			job.GetJobStatus(jobStatus)
			Select Case jobStatus.completion
				Case XPS_JOB_COMPLETION.XPS_JOB_COMPLETED
					' Expected result, do nothing.
				Case XPS_JOB_COMPLETION.XPS_JOB_FAILED
					Throw New Win32Exception(jobStatus.jobStatus)
				Case Else
					Throw New Exception("Unexpected print job status.")
			End Select
		End Sub

		<DllImport("XpsPrint.dll", EntryPoint := "StartXpsPrintJob")> _
		Private Shared Function StartXpsPrintJob(<MarshalAs(UnmanagedType.LPWStr)> ByVal printerName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal jobName As String, <MarshalAs(UnmanagedType.LPWStr)> ByVal outputFileName As String, ByVal progressEvent As IntPtr, ByVal completionEvent As IntPtr, <MarshalAs(UnmanagedType.LPArray)> ByVal printablePagesOn() As Byte, ByVal printablePagesOnCount As UInt32, <System.Runtime.InteropServices.Out()> ByRef xpsPrintJob As IXpsPrintJob, <System.Runtime.InteropServices.Out()> ByRef documentStream As IXpsPrintJobStream, ByVal printTicketStream As IntPtr) As Integer ' This is actually "out IXpsPrintJobStream", but we don't use it and just want to pass null, hence IntPtr.
		End Function

		<DllImport("Kernel32.dll", SetLastError := True)> _
		Private Shared Function CreateEvent(ByVal lpEventAttributes As IntPtr, ByVal bManualReset As Boolean, ByVal bInitialState As Boolean, ByVal lpName As String) As IntPtr
		End Function

		<DllImport("Kernel32.dll", SetLastError := True, ExactSpelling := True)> _
		Private Shared Function WaitForSingleObject(ByVal handle As IntPtr, ByVal milliseconds As Int32) As WAIT_RESULT
		End Function

		<DllImport("Kernel32.dll", SetLastError := True)> _
		Private Shared Function CloseHandle(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
		End Function
	End Class

	''' <summary>
	''' This interface definition is HACKED.
	''' 
	''' It appears that the IID for IXpsPrintJobStream specified in XpsPrint.h as 
	''' MIDL_INTERFACE("7a77dc5f-45d6-4dff-9307-d8cb846347ca") is not correct and the RCW cannot return it.
	''' But the returned object returns the parent ISequentialStream inteface successfully.
	''' 
	''' So the hack is that we obtain the ISequentialStream interface but work with it as 
	''' with the IXpsPrintJobStream interface. 
	''' </summary>
	<Guid("0C733A30-2A1C-11CE-ADE5-00AA0044773D"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
	Friend Interface IXpsPrintJobStream
		' ISequentualStream methods.
		Sub Read(<MarshalAs(UnmanagedType.LPArray)> ByVal pv() As Byte, ByVal cb As UInteger, <System.Runtime.InteropServices.Out()> ByRef pcbRead As UInteger)
		Sub Write(<MarshalAs(UnmanagedType.LPArray)> ByVal pv() As Byte, ByVal cb As UInteger, <System.Runtime.InteropServices.Out()> ByRef pcbWritten As UInteger)
		' IXpsPrintJobStream methods.
		Sub Close()
	End Interface

	<Guid("5ab89b06-8194-425f-ab3b-d7a96e350161"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)> _
	Friend Interface IXpsPrintJob
		Sub Cancel()
		Sub GetJobStatus(<System.Runtime.InteropServices.Out()> ByRef jobStatus As XPS_JOB_STATUS)
	End Interface

	<StructLayout(LayoutKind.Sequential)> _
	Friend Structure XPS_JOB_STATUS
		Public jobId As UInt32
		Public currentDocument As Int32
		Public currentPage As Int32
		Public currentPageTotal As Int32
		Public completion As XPS_JOB_COMPLETION
		Public jobStatus As Int32 ' UInt32
	End Structure

	Friend Enum XPS_JOB_COMPLETION
		XPS_JOB_IN_PROGRESS = 0
		XPS_JOB_COMPLETED = 1
		XPS_JOB_CANCELLED = 2
		XPS_JOB_FAILED = 3
	End Enum

	Friend Enum WAIT_RESULT
		WAIT_OBJECT_0 = 0
		WAIT_ABANDONED = &H80
		WAIT_TIMEOUT = &H102
		WAIT_FAILED = -1 ' 0xFFFFFFFF
	End Enum
End Namespace