Option Strict On
Module Module1

	Dim ExitCode As Integer
	Dim ErrorMessageList As List(Of String)

	Function Main() As Integer

		Console.WriteLine(String.Format("{0} starting...", System.AppDomain.CurrentDomain.FriendlyName))

		ExitCode = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgeAssembly.AssemblyDocument = Nothing

		Dim Configuration As New Dictionary(Of String, String)
		ErrorMessageList = New List(Of String)

		Dim ProgramSettings As New Dictionary(Of String, String)

		Dim BomDict As New Dictionary(Of String, Double)
		Dim PropDict As New Dictionary(Of String, String)
		'Dim SourceAssyfilename As String

		' Sets ExitCode = 1 on error.
		ProgramSettings = GetProgramSettings()

		'Configuration = GetConfiguration()

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
				SEDoc = CType(SEApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)
			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add("Assembly file not active in Solid Edge")
			End Try
		End If

		If ExitCode = 0 Then
			PropDict(ProgramSettings("QuantityPropertyName")) = "0"
			PropDict(ProgramSettings("SourceAssemblyPropertyName")) = System.IO.Path.GetFileNameWithoutExtension(SEDoc.FullName)

			' Sets ExitCode = 1 if an occurrence could not be found.
			BomDict = GetOccurrences(SEDoc, BomDict, ProgramSettings)

		End If

		If ExitCode = 0 Then

			'' Save and close assembly
			'SourceAssyfilename = SEDoc.FullName
			'Console.WriteLine("Closing assembly")
			'SEDoc.Save()
			'SEApp.DoIdle()
			'SEDoc.Close()
			'SEApp.DoIdle()

			' Add the properties to each file.
			PopulateProps(BomDict, PropDict, ProgramSettings("QuantityPropertyName"))

			'' Reopen assembly
			'Console.WriteLine("Opening assembly")
			'SEApp.Documents.Open(SourceAssyfilename)
			'SEApp.DoIdle()
			'SEApp.DisplayAlerts = True

			'SEDoc = CType(SEApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)

		End If

		If ExitCode <> 0 Then
			SaveErrorMessages(ErrorMessageList)
		End If

		Console.WriteLine(String.Format("{0} complete", System.AppDomain.CurrentDomain.FriendlyName))

		Return ExitCode
	End Function

	Private Sub PopulateProps(
		BomDict As Dictionary(Of String, Double),
		PropDict As Dictionary(Of String, String),
		QuantityPropertyName As String)

		Dim PropSets As SolidEdgeFileProperties.PropertySets = New SolidEdgeFileProperties.PropertySets
		Dim Prop As SolidEdgeFileProperties.Property = Nothing
		Dim Props As SolidEdgeFileProperties.Properties = Nothing

		Dim Filename As String
		Dim msg As String
		Dim FoundProp As Boolean

		For Each Filename In BomDict.Keys
			msg = String.Format("Updating properties: {0}", IO.Path.GetFileName(Filename))
			Console.WriteLine(msg)

			Try
				PropDict(QuantityPropertyName) = CType(BomDict(Filename), String)

				PropSets.Open(Filename, False)

				For Each Props In PropSets
					If Props.Name = "Custom" Then
						For Each Propname As String In PropDict.Keys
							FoundProp = False
							For Each Prop In Props
								If Prop.Name = Propname Then
									FoundProp = True
									Prop.Value = PropDict(Propname)
								End If
							Next
							If Not FoundProp Then
								Props.Add(Propname, PropDict(Propname))
							End If
						Next
					End If
				Next

				'Props.Save()
				PropSets.Save()
				PropSets.Close()

			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add(String.Format("Problem updating properties in {0}", Filename))
			End Try

		Next

		PropSets.Close()

	End Sub

	Private Function GetOccurrences(
		SEDoc As SolidEdgeAssembly.AssemblyDocument,
		BomDict As Dictionary(Of String, Double),
		ProgramSettings As Dictionary(Of String, String)
		) As Dictionary(Of String, Double)

		'Dim tf As Boolean

		Dim Occurrences As SolidEdgeAssembly.Occurrences
		Dim Occurrence As SolidEdgeAssembly.Occurrence
		Dim OccurrenceDoc As SolidEdgeFramework.SolidEdgeDocument
		Dim SubDoc As SolidEdgeAssembly.AssemblyDocument

		Dim QtyMultiplier As Integer = CInt(ProgramSettings("QuantityMultiplier"))
		Dim IncludeWeldmentParts As Boolean = CBool(ProgramSettings("IncludeWeldmentParts"))

		Dim NewQtyMultiplier As Double

		Dim msg As String

		Occurrences = SEDoc.Occurrences

		For Each Occurrence In Occurrences
			msg = String.Format("Getting quantities: {0}", IO.Path.GetFileName(Occurrence.OccurrenceFileName))
			Console.WriteLine(msg)
			Try
				If Occurrence.IncludeInBom Then
					If Occurrence.FileMissing Then
						ExitCode = 1
						msg = String.Format("File not found {0}", Occurrence.OccurrenceFileName)
						If Not ErrorMessageList.Contains(msg) Then
							ErrorMessageList.Add(msg)
						End If
						Continue For
					End If

					If Occurrence.HasNongraphicQuantity Then
						NewQtyMultiplier = QtyMultiplier * Occurrence.NongraphicQuantity
					Else
						NewQtyMultiplier = QtyMultiplier * Occurrence.Quantity
					End If

					OccurrenceDoc = CType(Occurrence.OccurrenceDocument, SolidEdgeFramework.SolidEdgeDocument)
					If Not BomDict.Keys.Contains(OccurrenceDoc.FullName.ToLower) Then
						BomDict(OccurrenceDoc.FullName.ToLower) = NewQtyMultiplier
					Else
						BomDict(OccurrenceDoc.FullName.ToLower) += NewQtyMultiplier
					End If
					If Occurrence.Subassembly Then
						SubDoc = CType(Occurrence.OccurrenceDocument, SolidEdgeAssembly.AssemblyDocument)

						'If Not SubDoc.WeldmentAssembly Then
						'	If Not Occurrence.FileMissing Then
						'		BomDict = GetOccurrences(SubDoc, BomDict, NewQtyMultiplier)
						'	End If
						'End If
						If Not Occurrence.FileMissing Then
							If SubDoc.WeldmentAssembly Then
								If IncludeWeldmentParts Then
									BomDict = GetOccurrences(SubDoc, BomDict, ProgramSettings)
								End If
							Else
								BomDict = GetOccurrences(SubDoc, BomDict, ProgramSettings)
							End If
						Else
							ExitCode = 1
							ErrorMessageList.Add(String.Format("Occurrence '{0}' not found", Occurrence.Name))
						End If

					End If
				End If

			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add(String.Format("Problem with occurrence {0}", Occurrence.Name))
			End Try

		Next

		Return BomDict
	End Function

	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		Dim ErrorFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		Dim HousekeeperRunning As Boolean = False
		Dim msg As String = ""

		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)
		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

		' See if Housekeeper is running
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

End Module
