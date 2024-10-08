﻿Option Strict On

Imports Newtonsoft.Json

Module Module1

	Function Main() As Integer

		Console.WriteLine("FitIsoView starting...")

		Dim ExitCode As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.  Edit as required.

		Dim SEApp As SolidEdgeFramework.Application
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument

		Dim ErrorMessageList As New List(Of String)

		Dim Settings As New Dictionary(Of String, String)
		Settings = GetSettings()

		Try
			SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
			SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)

			If SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument Then
				Try
					SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
					SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))
				Catch ex As Exception
					ExitCode = 1
					ErrorMessageList.Add("Error fitting view")
				End Try


			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igPartDocument Then
				Try
					SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
					SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))
				Catch ex As Exception
					ExitCode = 1
					ErrorMessageList.Add("Error fitting view")
				End Try


			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument Then
				Try
					SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
					SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))
				Catch ex As Exception
					ExitCode = 1
					ErrorMessageList.Add("Error fitting view")
				End Try


			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
				Try
					Dim SheetWindow As SolidEdgeDraft.SheetWindow = CType(SEApp.ActiveWindow, SolidEdgeDraft.SheetWindow)
					SheetWindow.FitEx(SolidEdgeDraft.SheetFitConstants.igFitSheet)
				Catch ex As Exception
					ExitCode = 1
					ErrorMessageList.Add("Error fitting view")
				End Try

			Else
				ExitCode = 1
				ErrorMessageList.Add("Unrecognized document type")
			End If

			If ExitCode = 0 Then
				If SEDoc.ReadOnly Then
					ExitCode = 1
					ErrorMessageList.Add("Cannot save read-only document")
				Else
					SEDoc.Save()
					SEApp.DoIdle()
				End If
			End If

		Catch ex As Exception
			ExitCode = 1
			ErrorMessageList.Add("Error connecting to Solid Edge")
		End Try

		If ExitCode <> 0 Then
			SaveErrorMessages(ErrorMessageList)
		End If
		Console.WriteLine("FitIsoView complete")

		Return ExitCode
	End Function


	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		Dim ErrorFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

	End Sub


	Private Function GetSettings() As Dictionary(Of String, String)
		Dim SettingsFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory  ' Contains trailing '\'

		SettingsFilename = String.Format("{0}form_main_settings.json", StartupPath)
		Dim tmpJSONDict As New Dictionary(Of String, String)
		Dim JSONString As String

		Try
			JSONString = IO.File.ReadAllText(SettingsFilename)

			tmpJSONDict = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(JSONString)

		Catch ex As Exception
		End Try

		Return tmpJSONDict
	End Function


End Module
