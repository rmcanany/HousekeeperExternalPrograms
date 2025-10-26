Option Strict On
Module Module1

	Dim ExitCode As Integer
	Dim ErrorMessageList As List(Of String)

	Function Main() As Integer

		Console.WriteLine(String.Format("{0} starting...", System.AppDomain.CurrentDomain.FriendlyName))

		ExitCode = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgeDraft.DraftDocument = Nothing
		Dim TemplateDoc As SolidEdgeDraft.DraftDocument = Nothing

		Dim Configuration As New Dictionary(Of String, String)
		ErrorMessageList = New List(Of String)

		Dim ProgramSettings As New Dictionary(Of String, String)

		' Sets ExitCode = 1 on error.
		ProgramSettings = GetProgramSettings()

		' The variable names are: BackgroundSheetNames, TemplateFilename, LinkFilenames, BlockNames

		Dim BackgroundSheetNames As List(Of String) = GetSettingsList(ProgramSettings, "BackgroundSheetNames")
		If BackgroundSheetNames Is Nothing Then ExitCode = 1 : ErrorMessageList.Add("Background sheet names not found")

		Dim TemplateFilename As String = GetSettingsList(ProgramSettings, "TemplateFilename")(0)
		If TemplateFilename Is Nothing OrElse Not IO.File.Exists(TemplateFilename) Then ExitCode = 1 : ErrorMessageList.Add("Template not found")

		Dim LinkFilenames As List(Of String) = GetSettingsList(ProgramSettings, "LinkFilenames")
		If LinkFilenames Is Nothing Then ExitCode = 1 : ErrorMessageList.Add("Link file names not found")

		Dim BlockNames As List(Of String) = GetSettingsList(ProgramSettings, "BlockNames")
		If BlockNames Is Nothing Then ExitCode = 1 : ErrorMessageList.Add("Block names not found")
		If BlockNames IsNot Nothing And LinkFilenames IsNot Nothing Then
			If Not BlockNames.Count = LinkFilenames.Count Then
				ExitCode = 1
				ErrorMessageList.Add("Link file names and block names do not have the same number of entries")
			End If
		End If


		If ExitCode = 0 Then
			Try
				SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
				SEApp.DisplayAlerts = False
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Could not connect to Solid Edge")
			End Try
		End If

		If ExitCode = 0 Then
			Try
				SEDoc = CType(SEApp.ActiveDocument, SolidEdgeDraft.DraftDocument)
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Draft file not active in Solid Edge")
			End Try
		End If

		If ExitCode = 0 Then
			If SEDoc.FullName = TemplateFilename Then
				ExitCode = 1
				ErrorMessageList.Add("Cannot process the block library itself")
			End If
		End If

		If ExitCode = 0 Then
			Try
				TemplateDoc = CType(SEApp.Documents.Open(TemplateFilename), SolidEdgeDraft.DraftDocument)
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add($"Could not open template '{TemplateFilename}'")
			End Try
		End If

		If ExitCode = 0 Then
			Dim FoundLinkFilenames As New List(Of String)
			For Each Sheet As SolidEdgeDraft.Sheet In SEDoc.Sheets
				If BackgroundSheetNames.Contains(Sheet.Name) Or BackgroundSheetNames.Contains("*") Then
					Dim SmartFrames2d As SolidEdgeFrameworkSupport.SmartFrames2d = CType(Sheet.SmartFrames2d, SolidEdgeFrameworkSupport.SmartFrames2d)
					If SmartFrames2d IsNot Nothing Then
						Dim SmartFrame2d As SolidEdgeFrameworkSupport.SmartFrame2d
						Dim LinkFilename As String = ""
						For i As Integer = SmartFrames2d.Count - 1 To 0 Step -1
							Try
								' Check the SmartFrame2d link name and add it to the list of found names.
								SmartFrame2d = CType(SmartFrames2d(i), SolidEdgeFrameworkSupport.SmartFrame2d)
								LinkFilename = SmartFrame2d.LinkMoniker
								'If LinkFilename = "" Then LinkFilename = "PUTOLERANCE.doc"
								LinkFilename = IO.Path.GetFileName(LinkFilename)
								If Not FoundLinkFilenames.Contains(LinkFilename) Then FoundLinkFilenames.Add(LinkFilename)

								' Get the index of the link file name in LinkFilenames
								Dim idx As Integer = -1
								For j As Integer = 0 To LinkFilenames.Count - 1
									If LinkFilenames(j) = LinkFilename Then
										idx = j
										Exit For
									End If
								Next

								' Get the template block name corresponding to the link name.
								' Skip if the link file name was not in LinkFilenames
								Dim TemplateBlockName As String
								If idx = -1 Then
									Continue For
								Else
									TemplateBlockName = BlockNames(idx)
								End If

								' Find the block in the template.  Skip if not found.
								Dim TemplateBlock As SolidEdgeDraft.Block = Nothing
								For Each tmpTemplateBlock As SolidEdgeDraft.Block In TemplateDoc.Blocks
									If tmpTemplateBlock.Name = TemplateBlockName Then
										TemplateBlock = tmpTemplateBlock
										Exit For
									End If
								Next
								If TemplateBlock Is Nothing Then
									ErrorMessageList.Add($"Template block not found '{TemplateBlockName}'")
									Continue For
								End If

								Dim DocBlock As SolidEdgeDraft.Block = Nothing
								For Each tmpDocBlock As SolidEdgeDraft.Block In SEDoc.Blocks
									If tmpDocBlock.Name = TemplateBlockName Then DocBlock = tmpDocBlock
									Exit For
								Next

								If DocBlock Is Nothing Then
									Dim x As Double
									Dim y As Double
									SmartFrame2d.GetOrigin(x, y)

									DocBlock = SEDoc.Blocks.CopyBlock(TemplateBlock)
									Dim BlockOccurrence As SolidEdgeDraft.BlockOccurrence = Sheet.BlockOccurrences.Add(DocBlock.Name, x, y)
									SmartFrame2d.Delete()
									SEApp.DoIdle()
								Else
									ErrorMessageList.Add($"Sheet already has a block named '{TemplateBlockName}'")
								End If
							Catch ex As Exception
								ErrorMessageList.Add($"Could not process OLE Link '{LinkFilename}'")
								ErrorMessageList.Add($"Error was {ex.Message}")
							End Try
						Next

					End If
				End If
			Next
			' Report if a link file name was not found in the file.
			For Each LinkFilename As String In LinkFilenames
				If Not FoundLinkFilenames.Contains(LinkFilename) Then
					ErrorMessageList.Add($"Link file name not found '{LinkFilename}'")
				End If
			Next
		End If


		If TemplateDoc IsNot Nothing Then
			TemplateDoc.Close()
			SEApp.DoIdle()
		End If

		If Not ExitCode = 0 Then
			SaveErrorMessages(ErrorMessageList)
		Else
			SEDoc.Save()
			SEApp.DoIdle()
		End If

		Console.WriteLine(String.Format("{0} complete", System.AppDomain.CurrentDomain.FriendlyName))

		Return ExitCode
	End Function


	Private Function GetProgramSettings() As Dictionary(Of String, String)

		Dim ProgramSettings As New Dictionary(Of String, String)
		Dim Settings As String() = Nothing
		Dim Key As String
		Dim Value As String
		Dim ProgramSettingsFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		ProgramSettingsFilename = String.Format("{0}program_settings.txt", StartupPath)

		Try
			Settings = IO.File.ReadAllLines(ProgramSettingsFilename)

			For Each KVPair As String In Settings

				Dim s As String = KVPair.Trim()

				If (s = "") Or (s = "'") Then
					Continue For
				End If

				If Not s.Contains("="c) Then
					Continue For
				End If

				Key = s.Split("="c)(0)
				Value = s.Split("="c)(1)

				ProgramSettings(Key.Trim()) = Value.Trim()
			Next

		Catch ex As Exception
			ExitCode = 1
			ErrorMessageList.Add(String.Format("Problem reading {0}", ProgramSettingsFilename))
		End Try

		Return ProgramSettings

	End Function


	Private Function GetSettingsList(
		ProgramSettings As Dictionary(Of String, String),
		VariableName As String
		) As List(Of String)

		Dim OutList As New List(Of String)

		If Not ProgramSettings.Keys.Contains(VariableName) Then
			OutList = Nothing
		Else
			Dim tmpOutList As List(Of String) = ProgramSettings(VariableName).Split(","c).ToList

			For Each s As String In tmpOutList
				If Not s.Trim = "" Then OutList.Add(s.Trim)
			Next

			If OutList.Count = 0 Then OutList = Nothing
		End If

		Return OutList
	End Function


	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		Dim ErrorFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		Dim HousekeeperRunning As Boolean = False
		Dim msg As String = ""

		' Save to file
		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)
		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

		' See if Housekeeper is running.  If not, show errors in a MsgBox.
		For Each P As Process In Process.GetProcesses()
			If P.ProcessName.ToLower = "housekeeper" Then
				HousekeeperRunning = True
				Exit For
			End If
		Next

		If Not HousekeeperRunning Then
			For Each s As String In ErrorMessageList
				msg = String.Format("{0}{1}{2}", msg, Chr(13), s)
			Next
			MsgBox(msg)
		End If

	End Sub

	'Private Function GetConfiguration() As Dictionary(Of String, String)
	'	Dim Configuration As New Dictionary(Of String, String)
	'	Dim Defaults As String() = Nothing
	'	Dim Key As String
	'	Dim Value As String
	'	Dim DefaultsFilename As String
	'	Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

	'	DefaultsFilename = String.Format("{0}defaults.txt", StartupPath)

	'	Try
	'		Defaults = IO.File.ReadAllLines(DefaultsFilename)

	'		For Each KVPair As String In Defaults
	'			If Not KVPair.Contains("=") Then
	'				Continue For
	'			End If

	'			Key = KVPair.Split("="c)(0)
	'			Value = KVPair.Split("="c)(1)

	'			Configuration(Key) = Value
	'		Next

	'	Catch ex As Exception
	'		ExitCode = 1
	'		ErrorMessageList.Add(String.Format("Problem reading {0}", DefaultsFilename))
	'	End Try


	'	Return Configuration
	'End Function



End Module
