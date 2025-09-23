' https://community.sw.siemens.com/s/question/0D54O000061xpllSAA/add-flat-pattern-through-the-api-woes
' https://community.sw.siemens.com/s/question/0D54O000061xs5ISAQ/applicationstartcommandcommandid
' https://community.sw.siemens.com/s/question/0D54O000061x3leSAA/from-design-model-to-flat-pattern
' https://community.sw.siemens.com/s/question/0D54O00006rjisUSAQ/programicatcially-created-flat-pattern-cut-size-variables-Not-working

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

			Select Case GetDocType(SEDoc)
				Case "psm"
					Dim tmpSEDoc As SolidEdgePart.SheetMetalDocument = CType(SEDoc, SolidEdgePart.SheetMetalDocument)
					Models = tmpSEDoc.Models
					FlatPatternModels = tmpSEDoc.FlatPatternModels
				Case "par"
					Dim tmpSEDoc As SolidEdgePart.PartDocument = CType(SEDoc, SolidEdgePart.PartDocument)
					Models = tmpSEDoc.Models
					FlatPatternModels = tmpSEDoc.FlatPatternModels
				Case Else
					ExitStatus = 1
					ErrorMessageList.Add(String.Format("Unrecognized file type '{0}'", GetDocType(SEDoc)))
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
			'Dim Edges As SolidEdgeGeometry.Edges = Nothing
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

					LongestLinearEdge = GetLongestLinearEdge(LargestFace)

				Catch ex As Exception
					ExitStatus = 1
					ErrorMessageList.Add("Unable to process face edges")
				End Try

				If LongestLinearEdge Is Nothing Then
					ExitStatus = 1
					ErrorMessageList.Add("No linear edges found")
				End If

			End If

			'Dim X As Double
			'Dim Y As Double

			'Dim DocDimensionDict As Dictionary(Of String, SolidEdgeFrameworkSupport.Dimension)
			'DocDimensionDict = GetDocDimensions(SEDoc)

			'If ExitStatus = 0 Then
			'	For Each DimensionName As String In DocDimensionDict.Keys
			'		If DimensionName = "Flat_Pattern_Model_CutSizeX" Then
			'			'Dim tf = DocDimensionDict(DimensionName).IsReadOnly
			'			DocDimensionDict(DimensionName).Formula = CStr(X)
			'		End If
			'		If DimensionName = "Flat_Pattern_Model_CutSizeY" Then
			'			DocDimensionDict(DimensionName).Value = Y
			'		End If
			'	Next
			'End If

			If ExitStatus = 0 Then
				Try
					FlatPatternModel = FlatPatternModels.Add(Model)
					FlatPatterns = FlatPatternModel.FlatPatterns
					'FlatPattern = FlatPatterns.Add(Edge)  ' Needs a face to get the flatpattern oriented to the top view
					FlatPattern = FlatPatterns.Add(LongestLinearEdge, LargestFace, LargestFace)

					FlatPattern.SetCutSizeValues(
						MaxCutSizeX:=0, MaxCutSizeY:=0, ShowRangeBox:=True, AlarmOnX:=False, AlarmOnY:=False, UseDefaultValues:=True)

					'FlatPatternModel.UpdateCutSize()
					'FlatPatternModel.GetCutSize(X, Y)
					'Dim i = 0
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

	'Public Function GetDocDimensions(SEDoc As SolidEdgeFramework.SolidEdgeDocument
	') As Dictionary(Of String, SolidEdgeFrameworkSupport.Dimension)
	'	Dim DocDimensionDict As New Dictionary(Of String, SolidEdgeFrameworkSupport.Dimension)

	'	Dim Variables As SolidEdgeFramework.Variables = Nothing
	'	Dim VariableListObject As SolidEdgeFramework.VariableList = Nothing
	'	Dim Variable As SolidEdgeFramework.variable = Nothing
	'	Dim Dimension As SolidEdgeFrameworkSupport.Dimension = Nothing
	'	Dim VariableTypeName As String

	'	Try
	'		Variables = DirectCast(SEDoc.Variables, SolidEdgeFramework.Variables)

	'		VariableListObject = DirectCast(Variables.Query(pFindCriterium:="*",
	'							  NamedBy:=SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
	'							  VarType:=SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth),
	'							  SolidEdgeFramework.VariableList)

	'		' Populate dictionary
	'		For Each VariableListItem In VariableListObject.OfType(Of Object)()
	'			VariableTypeName = Microsoft.VisualBasic.Information.TypeName(VariableListItem)

	'			If VariableTypeName.ToLower() = "dimension" Then
	'				Dimension = CType(VariableListItem, SolidEdgeFrameworkSupport.Dimension)
	'				DocDimensionDict(Dimension.DisplayName) = Dimension
	'			End If
	'		Next

	'	Catch ex As Exception
	'	End Try

	'	Return DocDimensionDict
	'End Function


	Private Function GetLongestLinearEdge(LargestFace As SolidEdgeGeometry.Face) As SolidEdgeGeometry.Edge

		Dim LongestEdge As SolidEdgeGeometry.Edge = Nothing

		Dim Edges As SolidEdgeGeometry.Edges = CType(LargestFace.Edges, SolidEdgeGeometry.Edges)

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


	Public Function GetDocType(SEDoc As SolidEdgeFramework.SolidEdgeDocument) As String
		' See SolidEdgeFramework.DocumentTypeConstants

		' If the type is not recognized, the empty string is returned.
		Dim DocType As String = ""

		If Not IsNothing(SEDoc) Then
			Select Case SEDoc.Type

				Case SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument
					DocType = "asm"
				Case SolidEdgeFramework.DocumentTypeConstants.igWeldmentAssemblyDocument
					DocType = "asm"
				Case SolidEdgeFramework.DocumentTypeConstants.igSyncAssemblyDocument
					DocType = "asm"
				Case SolidEdgeFramework.DocumentTypeConstants.igPartDocument
					DocType = "par"
				Case SolidEdgeFramework.DocumentTypeConstants.igSyncPartDocument
					DocType = "par"
				Case SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument
					DocType = "psm"
				Case SolidEdgeFramework.DocumentTypeConstants.igSyncSheetMetalDocument
					DocType = "psm"
				Case SolidEdgeFramework.DocumentTypeConstants.igDraftDocument
					DocType = "dft"

				Case Else
					MsgBox(String.Format("DocType '{0}' not recognized", SEDoc.Type.ToString))
			End Select
		End If

		Return DocType
	End Function

	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		' Saves error_messages.txt to the directory of the external program

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
