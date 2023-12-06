Option Strict On

Module Module1

	Function Main() As Integer

		Console.WriteLine(String.Format("{0} starting...", System.AppDomain.CurrentDomain.FriendlyName))

		Dim ExitCode As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.  
		Dim ErrorMessageList As New List(Of String)

		Dim ProgramSettings As New Dictionary(Of String, String)
		Dim Threshold As Double

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgePart.SheetMetalDocument = Nothing

		Dim Models As SolidEdgePart.Models = Nothing
		Dim Model As SolidEdgePart.Model = Nothing
		Dim ModelBody As SolidEdgeGeometry.Body = Nothing
		Dim FlatPatternModels As SolidEdgePart.FlatPatternModels = Nothing
		Dim FlatPatternModel As SolidEdgePart.FlatPatternModel = Nothing
		Dim FlatPatternModelBody As SolidEdgeGeometry.Body = Nothing

		Dim NormalizedVolumeDifference As Double

		ProgramSettings = GetProgramSettings()

		If ProgramSettings.Keys.Contains("Threshold") Then
			Try
				Threshold = CDbl(ProgramSettings("Threshold"))
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Could not convert program setting 'Threshold' to a floating point number")
			End Try
		Else
			ExitCode = 1
			ErrorMessageList.Add("Could not find 'Threshold' in program_settings.txt")
		End If

		' Dim Configuration As New Dictionary(Of String, String)

		' Configuration = GetConfiguration()

		If ExitCode = 0 Then
			Try
				SEApp = CType(MarshalHelper.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
				SEDoc = CType(SEApp.ActiveDocument, SolidEdgePart.SheetMetalDocument)
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Could not connect to Solid Edge, or Sheetmetal document not found")
			End Try
		End If


		If ExitCode = 0 Then
			Try
				Models = SEDoc.Models
				FlatPatternModels = SEDoc.FlatPatternModels
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Models or FlatPatternModels not found")
			End Try

			If Models.Count <> 1 Then
				ExitCode = 1
				ErrorMessageList.Add(String.Format("Only 1 model allowed.  {0} found", Models.Count))
			End If

			If FlatPatternModels.Count <> 1 Then
				ExitCode = 1
				ErrorMessageList.Add(String.Format("Only 1 flat pattern model allowed.  {0} found", FlatPatternModels.Count))
			End If

		End If

		If ExitCode = 0 Then
			Try
				Model = Models.Item(1)
				ModelBody = CType(Model.Body, SolidEdgeGeometry.Body)

				FlatPatternModel = FlatPatternModels.Item(1)
				FlatPatternModelBody = CType(FlatPatternModel.Body, SolidEdgeGeometry.Body)
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Model body or Flat pattern body not found")
			End Try

			NormalizedVolumeDifference = Math.Abs(ModelBody.Volume - FlatPatternModelBody.Volume) / ModelBody.Volume

			If NormalizedVolumeDifference > Threshold Then
				ExitCode = 1
				ErrorMessageList.Add(
					String.Format(
						"Volume difference is {0}, higher than threshold, {1}",
						NormalizedVolumeDifference,
						Threshold))
			End If
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

		If ExitCode <> 0 Then
			SaveErrorMessages(ErrorMessageList)
		End If

		Console.WriteLine(String.Format("{0} complete", System.AppDomain.CurrentDomain.FriendlyName))

		Return ExitCode
	End Function


	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		Dim ErrorFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

	End Sub

	Private Function GetProgramSettings() As Dictionary(Of String, String)
		Dim ProgramSettings As New Dictionary(Of String, String)
		Dim Settings As String() = Nothing
		Dim Key As String
		Dim Value As String
		Dim ProgramSettingsFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		ProgramSettingsFilename = String.Format("{0}\program_settings.txt", StartupPath)

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
		End Try


		Return ProgramSettings

	End Function


	Private Function GetConfiguration() As Dictionary(Of String, String)
		Dim Configuration As New Dictionary(Of String, String)
		Dim Defaults As String() = Nothing
		Dim Key As String
		Dim Value As String
		Dim DefaultsFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		DefaultsFilename = String.Format("{0}\defaults.txt", StartupPath)

		Try
			Defaults = IO.File.ReadAllLines(DefaultsFilename)

			For Each KVPair As String In Defaults
				If Not KVPair.Contains("=") Then
					Continue For
				End If

				Key = KVPair.Split("="c)(0)
				Value = KVPair.Split("="c)(1)

				Configuration(Key) = Value
			Next

		Catch ex As Exception
		End Try


		Return Configuration
	End Function


End Module
