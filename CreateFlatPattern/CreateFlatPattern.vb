' https://community.sw.siemens.com/s/question/0D54O000061xpllSAA/add-flat-pattern-through-the-api-woes
' https://community.sw.siemens.com/s/question/0D54O000061xs5ISAQ/applicationstartcommandcommandid
' https://community.sw.siemens.com/s/question/0D54O000061x3leSAA/from-design-model-to-flat-pattern

Option Strict On

Imports Newtonsoft.Json

Module CreateFlatPattern

	Function Main() As Integer

		Console.WriteLine("CreateFlatPattern starting...")

		Dim ExitStatus As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing

		Dim ErrorMessageList As New List(Of String)

		Dim Settings As New Dictionary(Of String, String)
		Settings = GetSettings()

		Try
			SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
			SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
		Catch ex As Exception
			ExitStatus = 1
			ErrorMessageList.Add("Error connecting to Solid Edge, or no file open")
		End Try

		Dim ActiveEnvironment As String = ""
		Dim Models As SolidEdgePart.Models = Nothing
		Dim FlatPatternModels As SolidEdgePart.FlatPatternModels = Nothing

		If ExitStatus = 0 Then

			ActiveEnvironment = SEApp.ActiveEnvironment ' Ordered: "SheetMetal", Sync: "DMSheetMetal"
			If Not ActiveEnvironment.ToLower.Contains("sheetmetal") Then
				ExitStatus = 1
				ErrorMessageList.Add(String.Format("File active environment was '{0}', not 'SheetMetal'", ActiveEnvironment))
			End If

			Select Case IO.Path.GetExtension(SEDoc.FullName)
				Case ".psm"
					Dim tmpSEDoc As SolidEdgePart.SheetMetalDocument = CType(SEDoc, SolidEdgePart.SheetMetalDocument)
					Models = tmpSEDoc.Models
					FlatPatternModels = tmpSEDoc.FlatPatternModels
				Case ".par"
					Dim tmpSEDoc As SolidEdgePart.PartDocument = CType(SEDoc, SolidEdgePart.PartDocument)
					Models = tmpSEDoc.Models
					FlatPatternModels = tmpSEDoc.FlatPatternModels
				Case Else
					ExitStatus = 1
					ErrorMessageList.Add(String.Format("Unrecognized file type '{0}'", IO.Path.GetExtension(SEDoc.FullName)))
			End Select
		End If

		If ExitStatus = 0 Then

			Dim Model As SolidEdgePart.Model = Nothing
			Dim FlatPatternModel As SolidEdgePart.FlatPatternModel = Nothing
			Dim FlatPatterns As SolidEdgePart.FlatPatterns = Nothing
			Dim FlatPattern As SolidEdgePart.FlatPattern
			Dim Body As SolidEdgeGeometry.Body
			Dim Faces As SolidEdgeGeometry.Faces = Nothing
			Dim Face As SolidEdgeGeometry.Face
			Dim LargestFace As SolidEdgeGeometry.Face = Nothing
			Dim Edges As SolidEdgeGeometry.Edges = Nothing
			Dim LongestLinearEdge As SolidEdgeGeometry.Edge = Nothing
			Dim MaxArea As Double

			If Models Is Nothing OrElse Models.Count = 0 Then
				ExitStatus = 1
				ErrorMessageList.Add("No models detected")
			End If

			If FlatPatternModels Is Nothing Then
				ExitStatus = 1
				ErrorMessageList.Add("Unable to access the flat pattern model collection")

			ElseIf FlatPatternModels.Count > 0 Then
				For Each FlatPatternModel In FlatPatternModels
					FlatPatterns = FlatPatternModel.FlatPatterns
					If FlatPatterns.Count > 0 Then
						ExitStatus = 1
						ErrorMessageList.Add("Flat pattern already present")
						Exit For
					End If
				Next
				FlatPatternModel = Nothing
				FlatPatterns = Nothing
			End If

			If ExitStatus = 0 Then
				'For Each Model In Models
				'	If Model IsNot Nothing Then Exit For
				'Next
				Model = Models.Item(1)

				Try
					Body = CType(Model.Body, SolidEdgeGeometry.Body)
					Faces = CType(Body.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryPlane), SolidEdgeGeometry.Faces)
				Catch ex As Exception
					ExitStatus = 1
					ErrorMessageList.Add("Could not process model geometry")
				End Try

				If Faces IsNot Nothing AndAlso Faces.Count = 0 Then
					ExitStatus = 1
					ErrorMessageList.Add("No planar faces found")
				End If
			End If

			If ExitStatus = 0 Then

				Try
					MaxArea = 0
					For Each Face In Faces
						If Face.Area > MaxArea Then
							LargestFace = Face
							MaxArea = Face.Area
						End If
					Next

					Edges = CType(LargestFace.Edges, SolidEdgeGeometry.Edges)

					LongestLinearEdge = GetLongestLinearEdge(Edges)

				Catch ex As Exception
					ExitStatus = 1
					ErrorMessageList.Add("Unable to process face edges")
				End Try

				If LongestLinearEdge Is Nothing Then
					ExitStatus = 1
					ErrorMessageList.Add("No linear edges found")
				End If

			End If

			If ExitStatus = 0 Then
				Try
					FlatPatternModel = FlatPatternModels.Add(Model)
					FlatPatterns = FlatPatternModel.FlatPatterns
					'FlatPattern = FlatPatterns.Add(Edge)
					FlatPattern = FlatPatterns.Add(LongestLinearEdge, LargestFace, LargestFace)
				Catch ex As Exception
					ExitStatus = 1
					ErrorMessageList.Add("Unable to create flat pattern")
				End Try
			End If
		End If

		If ExitStatus = 0 Then
			If SEDoc.ReadOnly Then
				ExitStatus = 1
				ErrorMessageList.Add("Cannot save read-only document")
			Else
				Try
					' Ordered: "SheetMetal", Sync: "DMSheetMetal"
					' Exit flat pattern environment Command ID 10768 for Ordered, 10767 for Sync.
					If ActiveEnvironment = "SheetMetal" Then SEApp.StartCommand(CType(10768, SolidEdgeFramework.SolidEdgeCommandConstants))
					If ActiveEnvironment = "DMSheetMetal" Then SEApp.StartCommand(CType(10767, SolidEdgeFramework.SolidEdgeCommandConstants))
					SEApp.DoIdle()
				Catch ex As Exception
				End Try

				SEDoc.Save()
				SEApp.DoIdle()
			End If
		End If

		If ExitStatus <> 0 Then
			SaveErrorMessages(ErrorMessageList)
		End If
		Console.WriteLine("CreateFlatPattern complete")

		Return ExitStatus
	End Function

	Private Function GetLongestLinearEdge(Edges As SolidEdgeGeometry.Edges) As SolidEdgeGeometry.Edge

		Dim LongestEdge As SolidEdgeGeometry.Edge = Nothing

		Dim Edge As SolidEdgeGeometry.Edge
		Dim minParam As Double
		Dim maxParam As Double
		Dim Length As Double
		Dim MaxLength As Double = 0.0

		For Each Edge In Edges

			Dim Geometry = Edge.Geometry
			Dim Line = TryCast(Geometry, SolidEdgeGeometry.Line)
			If Line Is Nothing Then Continue For

			Edge.GetParamExtents(MinParam:=minParam, MaxParam:=maxParam)
			Edge.GetLengthAtParam(FromParam:=minParam, ToParam:=maxParam, Length:=Length)

			If Length > MaxLength Then
				MaxLength = Length
				LongestEdge = Edge
			End If
		Next

		Return LongestEdge
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
