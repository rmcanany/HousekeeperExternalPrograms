Option Strict On

Module UpdatePartsListStyleFromTemplate

	Dim ExitCode As Integer
	Dim ErrorMessageList As List(Of String)

	Function Main() As Integer

		Console.WriteLine(String.Format("{0} starting...", System.AppDomain.CurrentDomain.FriendlyName))

		ExitCode = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing

		ErrorMessageList = New List(Of String)

		Dim ProgramSettings As New Dictionary(Of String, String)

		' The variable names are: DraftListPath, BackgroundSheetNames, PartsListStyleNames
		Dim DraftListPath As String = ""
		Dim BackgroundSheetNames As List(Of String) = Nothing
		Dim PartsListStyleNames As List(Of String) = Nothing
		Dim LUT As New Dictionary(Of String, String)

		' Sets ExitCode = 1 on error.
		ProgramSettings = GetProgramSettings()

		'DraftListPath = GetSettingsList(ProgramSettings, "DraftListPath")(0).ToLower
		'If DraftListPath Is Nothing Then ExitCode = 1 : ErrorMessageList.Add("DraftListPath variable not found")

		If ExitCode = 0 Then
			Dim tmpList As List(Of String) = GetSettingsList(ProgramSettings, "DraftListPath")
			If tmpList Is Nothing Then
				ExitCode = 1
				ErrorMessageList.Add("DraftListPath variable not found")
			Else
				DraftListPath = tmpList(0)
			End If

			BackgroundSheetNames = GetSettingsList(ProgramSettings, "BackgroundSheetNames")
			If BackgroundSheetNames Is Nothing Then ExitCode = 1 : ErrorMessageList.Add("BackgroundSheetNames variable not found")

			PartsListStyleNames = GetSettingsList(ProgramSettings, "PartsListStyleNames")
			If PartsListStyleNames Is Nothing Then ExitCode = 1 : ErrorMessageList.Add("PartsListStyleNames variable not found")

			If Not BackgroundSheetNames.Count = PartsListStyleNames.Count Then ExitCode = 1 : ErrorMessageList.Add("Background names and style name count mismatch")

			If Not IO.File.Exists(DraftListPath) Then
				ExitCode = 1
				ErrorMessageList.Add(String.Format("File not found '{0}'", DraftListPath))
			End If

		End If

		If ExitCode = 0 Then
			For i As Integer = 0 To BackgroundSheetNames.Count - 1
				LUT(BackgroundSheetNames(i)) = PartsListStyleNames(i)
			Next
		End If

		If ExitCode = 0 Then
			Try
				SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
				SEApp.DisplayAlerts = False
				SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Could not connect to Solid Edge, or no active document found")
			End Try
		End If

		If ExitCode = 0 Then
			Dim DocType As String = IO.Path.GetExtension(SEDoc.FullName)
			If DocType = ".dft" Then

				'#### VERIFY PARTS LIST STYLE NAMES ####

				' DraftList.txt format
				'BEGINSET
				' CONFIGNAME = ANSI
				' TextStyle = Normal
				' ...
				'BEGINSET
				' CONFIGNAME = ISO
				' TextStyle = Normal
				' ...

				' Parse DraftList.txt file and store style names it contains.

				Dim KnownStyles As New List(Of String)
				Dim DraftListContents As List(Of String) = IO.File.ReadAllLines(DraftListPath).ToList
				Dim SearchingHeader As Boolean = True
				Dim SearchingName As Boolean = False

				For Each Line As String In DraftListContents
					If SearchingHeader Then
						If Line.ToUpper.Trim = "BEGINSET" Then
							SearchingHeader = False
							SearchingName = True
							Continue For
						End If
					End If
					If SearchingName Then
						If Line.ToUpper.Trim.StartsWith("CONFIGNAME") Then
							Dim KnownStyle As String = Line.Split("="c)(1).Trim
							If Not KnownStyles.Contains(KnownStyle) Then
								KnownStyles.Add(KnownStyle)
							Else
								ExitCode = 1
								ErrorMessageList.Add(String.Format("DraftList.txt duplicate style name found `{0}`", KnownStyle))
							End If
							SearchingHeader = True
							SearchingName = False
							Continue For
						End If
					End If
				Next

				' Check the lookup table for valid parts list style names

				For Each BackgroundName As String In LUT.Keys
					If Not KnownStyles.Contains(LUT(BackgroundName)) Then
						ExitCode = 1
						ErrorMessageList.Add(String.Format("DraftList.txt style name not found '{0}'", LUT(BackgroundName)))
					End If
				Next

				' #### UPDATE PARTS LIST STYLES ####

				If ExitCode = 0 Then
					Dim tmpSEDoc As SolidEdgeDraft.DraftDocument = CType(SEDoc, SolidEdgeDraft.DraftDocument)

					If tmpSEDoc.Sheets Is Nothing Then
						ExitCode = 1
						ErrorMessageList.Add("Could not process Sheets")
					Else
						For Each Sheet As SolidEdgeDraft.Sheet In tmpSEDoc.Sheets
							If Sheet.Section.Type = 0 Then
								For Each Item As Object In Sheet.DrawingObjects
									Dim PartsList As SolidEdgeDraft.PartsList = TryCast(Item, SolidEdgeDraft.PartsList)
									If PartsList IsNot Nothing Then
										Dim IsAlreadyPresent As Boolean = False
										For Each Key As String In LUT.Keys
											If Key = Sheet.Background.Name Then
												IsAlreadyPresent = True
												Exit For
											End If
										Next
										If Not LUT.Keys.Contains(Sheet.Background.Name) Then
											ExitCode = 1
											ErrorMessageList.Add(String.Format("Not in lookup table: Sheet '{0}' background '{1}'", Sheet.Name, Sheet.Background.Name))

										Else
											Try
												PartsList.SavedSettings = LUT(Sheet.Background.Name)
												PartsList.Update()
												SEApp.DoIdle()
											Catch ex As Exception
												ExitCode = 1
												Dim s As String = "Could not update parts list style: "
												s = String.Format("{0}Sheet '{1}' Style '{2}'{3}", s, Sheet.Name, LUT(Sheet.Background.Name), vbCrLf)
												s = String.Format("{0}Error message was: {1}", s, ex.Message)
												ErrorMessageList.Add(s)
											End Try

										End If
									End If
								Next
							End If
						Next

						Try
							SEDoc.Save()
							SEApp.DoIdle()
						Catch ex As Exception
							ExitCode = 1
							ErrorMessageList.Add(String.Format("Could not save file: '{0}'", ex.Message))
						End Try

					End If
				End If

			End If
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

				If (s = "") Or (s.StartsWith("'")) Then
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
			Dim tmpVariableName As String = ProgramSettings(VariableName).Replace("\,", "LITERALCOMMA")
			Dim tmpOutList As List(Of String) = tmpVariableName.Split(","c).ToList

			For Each s As String In tmpOutList
				OutList.Add(s.Replace("LITERALCOMMA", ",").Trim)
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

		' If Housekeeper is not running, the program is running stand-alone.  If so, show errors in a MsgBox.
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


End Module
