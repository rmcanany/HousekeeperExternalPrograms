Option Strict On
Module Module1

	Dim ExitCode As Integer
	Dim ErrorMessageList As List(Of String)

	Function Main() As Integer

		Console.WriteLine("QtyFromAssy starting...")

		ExitCode = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

		Dim SEApp As SolidEdgeFramework.Application
		Dim SEDoc As SolidEdgeAssembly.AssemblyDocument

		Dim Configuration As New Dictionary(Of String, String)
		ErrorMessageList = New List(Of String)

		Dim BomDict As New Dictionary(Of String, Double)
		Dim PropDict As New Dictionary(Of String, String)
		Dim SourceAssyfilename As String

		Configuration = GetConfiguration()

		SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
		'SEApp.DisplayAlerts = False

		SEDoc = CType(SEApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)

		If SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument Then
			BomDict = GetOccurrences(SEDoc, BomDict, 1)

			PropDict = BuildPropDict(SEDoc)

			SourceAssyfilename = SEDoc.FullName

			'Console.WriteLine("Closing assembly")
			'SEDoc.Close()
			'SEApp.DoIdle()

			PopulateProps(BomDict, PropDict)

			'Console.WriteLine("Opening assembly")
			'SEApp.Documents.Open(SourceAssyfilename)
			'SEApp.DoIdle()

			'SEDoc = CType(SEApp.ActiveDocument, SolidEdgeAssembly.AssemblyDocument)

		Else
			ExitCode = 1
			ErrorMessageList.Add("Assembly not found.  Assembly document must be open in Solid Edge for this command.")
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
		Console.WriteLine("QtyFromAssy complete")

		Return ExitCode
	End Function

	Private Sub PopulateProps(BomDict As Dictionary(Of String, Double), PropDict As Dictionary(Of String, String))

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
				PropDict("QtyFromAssy_Qty") = CType(BomDict(Filename), String)

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

				PropSets.Save()

			Catch ex As Exception
				ExitCode = 1
				ErrorMessageList.Add(String.Format("Problem updating properties in {0}", Filename))
			End Try

		Next

		PropSets.Close()

	End Sub


	Private Function BuildPropDict(SEDoc As SolidEdgeAssembly.AssemblyDocument) As Dictionary(Of String, String)
		Dim PropDict As New Dictionary(Of String, String)

		PropDict("QtyFromAssy_Qty") = "0"
		PropDict("QtyFromAssy_Assy") = SEDoc.Name


		Return PropDict
	End Function
	Private Function GetOccurrences(
		SEDoc As SolidEdgeAssembly.AssemblyDocument,
		BomDict As Dictionary(Of String, Double),
		QtyMultiplier As Double
		) As Dictionary(Of String, Double)

		'Dim tf As Boolean

		Dim Occurrences As SolidEdgeAssembly.Occurrences
		Dim Occurrence As SolidEdgeAssembly.Occurrence
		Dim OccurrenceDoc As SolidEdgeFramework.SolidEdgeDocument
		Dim SubDoc As SolidEdgeAssembly.AssemblyDocument

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
					NewQtyMultiplier = QtyMultiplier * Occurrence.Quantity
					OccurrenceDoc = CType(Occurrence.OccurrenceDocument, SolidEdgeFramework.SolidEdgeDocument)
					If Not BomDict.Keys.Contains(OccurrenceDoc.FullName.ToLower) Then
						BomDict(OccurrenceDoc.FullName.ToLower) = NewQtyMultiplier
					Else
						BomDict(OccurrenceDoc.FullName.ToLower) += NewQtyMultiplier
					End If
					If Occurrence.Subassembly Then
						SubDoc = CType(Occurrence.OccurrenceDocument, SolidEdgeAssembly.AssemblyDocument)
						If Not SubDoc.WeldmentAssembly Then
							If Not Occurrence.FileMissing Then
								BomDict = GetOccurrences(CType(OccurrenceDoc, SolidEdgeAssembly.AssemblyDocument), BomDict, NewQtyMultiplier)
							End If
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

		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

	End Sub


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
