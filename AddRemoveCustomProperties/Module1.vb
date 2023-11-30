Option Strict On
Module Module1

	Function Main() As Integer

		Console.WriteLine("AddRemoveCustomProperties starting...")

		Dim ExitCode As Integer = 0  ' 0 means success.  For a more complete example, see FitISOView.
		Dim ErrorMessageList As New List(Of String)

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing

		Dim PropertySets As SolidEdgeFramework.PropertySets = Nothing
        Dim Properties As SolidEdgeFramework.Properties = Nothing
        Dim Prop As SolidEdgeFramework.Property = Nothing
		Dim RemoveProps As New List(Of String)

		Dim OperatingMode As String = ""
		Dim OperatingModeCount As Integer = 0
		Dim OperatingModeIdx As Integer
		Dim idx As Integer = 0

		Dim Proceed As Boolean = True

		RemoveProps = GetRemoveProps()

		For Each Line As String In RemoveProps
			If Line.Contains("OperatingMode") Then
				OperatingModeCount += 1
				OperatingMode = Line.Split(" "c)(1)
				OperatingModeIdx = idx
			End If
			idx += 1
		Next

		If OperatingModeCount <> 1 Then
			Proceed = False
			ExitCode = 1
			If OperatingModeCount = 0 Then
				ErrorMessageList.Add("No OperatingMode specified in property_list.txt")
			Else
				ErrorMessageList.Add("Multiple OperatingModes specified in property_list.txt")
			End If
		End If

		If Proceed Then
			RemoveProps.RemoveAt(OperatingModeIdx)
		End If

		If Proceed Then
			Try
				SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
				SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
			Catch ex As Exception
				Proceed = False
				ExitCode = 1
				ErrorMessageList.Add("Error connecting to Solid Edge")
			End Try

		End If

		If Proceed Then
			PropertySets = CType(SEDoc.Properties, SolidEdgeFramework.PropertySets)

			For Each Properties In PropertySets
				If Properties.Name.ToLower = "custom" Then
					For Each Prop In Properties

						' Some props do not have a name
						Try
							If OperatingMode.ToLower = "remove" Then
								If RemoveProps.Contains(Prop.Name) Then
									Console.WriteLine(Prop.Name)
									Try
										Prop.Delete()
									Catch ex As Exception
										ExitCode = 1
										ErrorMessageList.Add(String.Format("Unable to delete property '{0}'", Prop.Name))
									End Try

								End If
							End If
							If OperatingMode.ToLower = "removeallexcept" Then
								If Not RemoveProps.Contains(Prop.Name) Then
									Console.WriteLine(Prop.Name)
									Try
										Prop.Delete()
									Catch ex As Exception
										ExitCode = 1
										ErrorMessageList.Add(String.Format("Unable to delete property '{0}'", Prop.Name))
									End Try

								End If
							End If

						Catch ex As Exception
						End Try
					Next

					Properties.Save()
				End If
			Next

		End If

		If ExitCode = 0 Then
			If SEDoc.ReadOnly Then
				ExitCode = 1
				ErrorMessageList.Add("Cannot save read-only file")
			Else
				SEDoc.Save()
				SEApp.DoIdle()
			End If
		End If

		If ExitCode <> 0 Then
			SaveErrorMessages(ErrorMessageList)
		End If

		Console.WriteLine("AddRemoveCustomProperties complete")

		' Short pause so the user can see the console feedback.
		System.Threading.Thread.Sleep(500)

		Return ExitCode
	End Function

	Function GetRemoveProps() As List(Of String)
		Dim tmpRemoveProps As New List(Of String)
		Dim RemoveProps As New List(Of String)
		Dim StartupPath As String = AppDomain.CurrentDomain.BaseDirectory

		Dim Line As String
		Dim TrimmedLine As String

		Dim Filename As String = String.Format("{0}\property_list.txt", StartupPath)

		Try
			tmpRemoveProps = IO.File.ReadAllLines(Filename).ToList
			For Each Line In tmpRemoveProps
				TrimmedLine = Line.Trim()
				If (Not TrimmedLine(0) = "'") Then
					If Not TrimmedLine = "" Then
						RemoveProps.Add(Line.Trim())
					End If
				End If
			Next
		Catch ex As Exception
			Console.WriteLine(String.Format("Unable to open {0}", Filename))
		End Try

		Return RemoveProps
	End Function

	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		Dim ErrorFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

	End Sub

End Module

