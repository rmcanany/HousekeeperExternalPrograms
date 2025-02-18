Option Strict On

Imports Newtonsoft.Json

Module Module1

	Function Main() As Integer

		Console.WriteLine("Recognize Holes starting...")

		Dim ExitStatus As Integer = 0  ' 0 means success

		Dim SEApp As SolidEdgeFramework.Application
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument

		Dim ErrorMessageList As New List(Of String)

		Dim Settings As New Dictionary(Of String, String)
		Settings = GetSettings()

		Try
			SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
			SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)

			Dim SupplementalErrorMessage As New Dictionary(Of Integer, List(Of String))

			Dim Proceed As Boolean = True

			Dim WasOrdered As Boolean = False

			Dim DocType = GetDocType(SEDoc)

			If SEDoc.ReadOnly Then
				Proceed = False
				ExitStatus = 1
				ErrorMessageList.Add("Cannot save document marked 'Read Only'")
			End If

			Dim Models As SolidEdgePart.Models = Nothing
			Dim Model As SolidEdgePart.Model

			If Proceed Then
				Select Case DocType

					Case "par"
						Dim tmpSEDoc As SolidEdgePart.PartDocument
						tmpSEDoc = CType(SEDoc, SolidEdgePart.PartDocument)

						Models = tmpSEDoc.Models
						If Models.Count = 0 Then Proceed = False  ' Not an error, but nothing to do.
						If Models.Count > 1 Then
							Proceed = False
							ExitStatus = 1
							ErrorMessageList.Add("Cannot process files with more than one model")
						End If

						If Proceed Then
							If tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered Then
								WasOrdered = True
							End If

							'Determine if the first body is in Synchronous Mode
							If WasOrdered Then
								Try
									tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous
								Catch ex As Exception
									'The first body are in Ordered Mode, move to Synchronous is needed
									SupplementalErrorMessage = MoveToSync(tmpSEDoc, Models.Item(1))
									AddSupplementalErrorMessage(ExitStatus, ErrorMessageList, SupplementalErrorMessage)
								End Try
							End If

						End If

					Case "psm"
						Dim tmpSEDoc As SolidEdgePart.SheetMetalDocument
						tmpSEDoc = CType(SEApp.ActiveDocument, SolidEdgePart.SheetMetalDocument)

						Models = tmpSEDoc.Models
						If Models.Count = 0 Then Proceed = False  ' Not an error, but nothing to do.
						If Models.Count > 1 Then
							Proceed = False
							ExitStatus = 1
							ErrorMessageList.Add("Cannot process files with more than one model")
						End If

						If Proceed Then
							If tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered Then
								WasOrdered = True
							End If

							'Determine if the first body is in Synchronous Mode
							If WasOrdered Then
								Try
									tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous
								Catch ex As Exception
									'The first body are in Ordered Mode, move to Synchronous is needed
									SupplementalErrorMessage = MoveToSync(tmpSEDoc, Models.Item(1))
									AddSupplementalErrorMessage(ExitStatus, ErrorMessageList, SupplementalErrorMessage)
								End Try
							End If

						End If

				End Select

			End If

			If Proceed Then
				'Heal and optimize the body
				Model = Models.Item(1)
				Try
					Model.HealAndOptimizeBody(False, True)
					SEApp.DoIdle()
				Catch ex As Exception
					Proceed = False
					ExitStatus = 1
					ErrorMessageList.Add("Geometry optimization did not succeed")
				End Try

			End If

			If Proceed Then
				'Recognize holes
				Model = Models.Item(1)
				Dim numBodies As Integer = 1
				Dim Body As SolidEdgeGeometry.Body
				Body = CType(Model.Body, SolidEdgeGeometry.Body)
				Dim Bodies As Array
				Bodies = New SolidEdgeGeometry.Body(0) {Body}
				Dim numHoles As Integer = 1
				Dim RecognizedHoles As Array
				RecognizedHoles = New Object() {}
				Try
					Model.Holes.RecognizeAndCreateHoleGroups(numBodies, Bodies, numHoles, RecognizedHoles)
					SEApp.DoIdle()
					Model.Recompute()
				Catch ex As Exception
					' Holes not found.  Not an error.
				End Try

			End If

			If Proceed And WasOrdered Then
				'Finish in Ordered Mode ready to work (This could be an option)
				Select Case DocType
					Case "par"
						Dim tmpSEDoc As SolidEdgePart.PartDocument
						tmpSEDoc = CType(SEDoc, SolidEdgePart.PartDocument)
						tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered
					Case "psm"
						Dim tmpSEDoc As SolidEdgePart.SheetMetalDocument
						tmpSEDoc = CType(SEDoc, SolidEdgePart.SheetMetalDocument)
						tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeOrdered
				End Select

			End If

			If Proceed Then
				'Save file
				SEDoc.Save()
				SEApp.DoIdle()
			End If

			If ExitStatus = 0 Then
				If SEDoc.ReadOnly Then
					ExitStatus = 1
					ErrorMessageList.Add("Cannot save read-only document")
				Else
					SEDoc.Save()
					SEApp.DoIdle()
				End If
			End If

		Catch ex As Exception
			ExitStatus = 1
			ErrorMessageList.Add("Error connecting to Solid Edge")
		End Try

		If ExitStatus <> 0 Then
			SaveErrorMessages(ErrorMessageList)
		End If
		Console.WriteLine("Recognize Holes complete")

		Return ExitStatus
	End Function

	Private Function MoveToSync(
		ByRef tmpSEDoc As SolidEdgePart.PartDocument,
		ByRef Model As SolidEdgePart.Model
		) As Dictionary(Of Integer, List(Of String))

		Dim ErrorMessageList As New List(Of String)
		Dim ExitStatus As Integer = 0
		Dim ErrorMessage As New Dictionary(Of Integer, List(Of String))

		Dim Features As SolidEdgePart.Features = Nothing
		Dim Feature As Object = Nothing
		Features = Model.Features

		Dim bIgnoreWarnings As Boolean = True
		Dim bExtentSelection As Boolean = True
		Dim aErrorMessages As Array
		Dim aWarningMessages As Array
		Dim lNumberOfFeaturesCausingError As Integer
		Dim lNumberOfFeaturesCausingWarning As Integer

		For Each Feature In Features
			aErrorMessages = Array.CreateInstance(GetType(String), 0)
			aWarningMessages = Array.CreateInstance(GetType(String), 0)
			Dim dVolumeDifference As Double = 0
			'MoveToSynchronous in Part Mode have 8 arguments
			tmpSEDoc.MoveToSynchronous(Feature,
									   bIgnoreWarnings,
									   bExtentSelection,
									   lNumberOfFeaturesCausingError,
									   aErrorMessages,
									   lNumberOfFeaturesCausingWarning,
									   aWarningMessages,
									   dVolumeDifference)

			If aErrorMessages.Length > 0 Then
				ExitStatus = 1
				For Each s As String In aErrorMessages
					ErrorMessageList.Add(s)
				Next
			End If

		Next
		tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous

		ErrorMessage(ExitStatus) = ErrorMessageList
		Return ErrorMessage
	End Function

	Private Function MoveToSync(
		ByRef tmpSEDoc As SolidEdgePart.SheetMetalDocument,
		ByRef Model As SolidEdgePart.Model
		) As Dictionary(Of Integer, List(Of String))

		Dim ErrorMessageList As New List(Of String)
		Dim ExitStatus As Integer = 0
		Dim ErrorMessage As New Dictionary(Of Integer, List(Of String))

		Dim Features As SolidEdgePart.Features = Nothing
		Dim Feature As Object = Nothing
		Features = Model.Features

		Dim bIgnoreWarnings As Boolean = True
		Dim bExtentSelection As Boolean = True
		Dim aErrorMessages As Array
		Dim aWarningMessages As Array
		Dim lNumberOfFeaturesCausingError As Integer
		Dim lNumberOfFeaturesCausingWarning As Integer

		For Each Feature In Features
			aErrorMessages = Array.CreateInstance(GetType(String), 0)
			aWarningMessages = Array.CreateInstance(GetType(String), 0)
			Dim dVolumeDifference As Double = 0
			'MoveToSynchronous in Part Mode have 8 arguments
			tmpSEDoc.MoveToSynchronous(Feature,
									   bIgnoreWarnings,
									   bExtentSelection,
									   lNumberOfFeaturesCausingError,
									   aErrorMessages,
									   lNumberOfFeaturesCausingWarning,
									   aWarningMessages)

			If aErrorMessages.Length > 0 Then
				ExitStatus = 1
				For Each s As String In aErrorMessages
					ErrorMessageList.Add(s)
				Next
			End If

		Next
		tmpSEDoc.ModelingMode = SolidEdgePart.ModelingModeConstants.seModelingModeSynchronous

		ErrorMessage(ExitStatus) = ErrorMessageList
		Return ErrorMessage
	End Function

	Public Sub AddSupplementalErrorMessage(
		ByRef ExitStatus As Integer,
		ErrorMessageList As List(Of String),
		SupplementalErrorMessage As Dictionary(Of Integer, List(Of String))
		)

		Dim SupplementalExitStatus As Integer = SupplementalErrorMessage.Keys(0)

		If Not SupplementalExitStatus = 0 Then
			If SupplementalExitStatus > ExitStatus Then
				ExitStatus = SupplementalExitStatus
			End If
			For Each s As String In SupplementalErrorMessage(SupplementalExitStatus)
				ErrorMessageList.Add(s)
			Next
		End If
	End Sub


	Public Function GetDocType(SEDoc As SolidEdgeFramework.SolidEdgeDocument) As String
		' See SolidEdgeFramework.DocumentTypeConstants

		' If the type is not recognized, the empty string is returned.
		Dim DocType As String = ""

		If Not IsNothing(SEDoc) Then
			Select Case SEDoc.Type

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument
					DocType = "asm"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igWeldmentAssemblyDocument
					DocType = "asm"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igSyncAssemblyDocument
					DocType = "asm"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igPartDocument
					DocType = "par"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igSyncPartDocument
					DocType = "par"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument
					DocType = "psm"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igSyncSheetMetalDocument
					DocType = "psm"

				Case Is = SolidEdgeFramework.DocumentTypeConstants.igDraftDocument
					DocType = "dft"

				Case Else
					MsgBox(String.Format("{0} DocType '{1}' not recognized", "Task_Common", SEDoc.Type.ToString))
			End Select
		End If

		Return DocType
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
