Option Strict On
Module Module1

	Function Main() As Integer

		Console.WriteLine("FitIsoView starting...")

		Dim ExitCode As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.  Edit as required.

		Dim SEApp As SolidEdgeFramework.Application
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument

		Dim Configuration As New Dictionary(Of String, String)

		Configuration = GetConfiguration()

		' Key-Value pairs for pictorial view selection from the file 'defaults.txt'
		'RadioButtonPictorialViewTrimetric = True
		'RadioButtonPictorialViewDimetric = False
		'RadioButtonPictorialViewIsometric = False

		SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
		SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)

		Try
			If SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument Then
				Try
					If Configuration("RadioButtonPictorialViewIsometric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
					If Configuration("RadioButtonPictorialViewDimetric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewDimetricView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
					If Configuration("RadioButtonPictorialViewTrimetric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewTrimetricView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
				Catch ex As Exception
					ExitCode = 4
				End Try

				SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))

			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igPartDocument Then
				Try
					If Configuration("RadioButtonPictorialViewIsometric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
					If Configuration("RadioButtonPictorialViewDimetric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewDimetricView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
					If Configuration("RadioButtonPictorialViewTrimetric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.SheetMetalViewTrimetricView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If

				Catch ex As Exception
					ExitCode = 4
				End Try

				SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))

			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument Then
				Try
					If Configuration("RadioButtonPictorialViewIsometric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
					If Configuration("RadioButtonPictorialViewDimetric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewDimetricView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If
					If Configuration("RadioButtonPictorialViewTrimetric").ToLower = "true" Then
						SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.SheetMetalViewTrimetricView, SolidEdgeFramework.SolidEdgeCommandConstants))
					End If

				Catch ex As Exception
					ExitCode = 4
				End Try

				SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))

			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igDraftDocument Then
				Dim SheetWindow As SolidEdgeDraft.SheetWindow = CType(SEApp.ActiveWindow, SolidEdgeDraft.SheetWindow)
				SheetWindow.FitEx(SolidEdgeDraft.SheetFitConstants.igFitSheet)

			Else
				ExitCode = 1
			End If

		Catch ex As Exception
			ExitCode = 2
		End Try

		If ExitCode = 0 Then
			If SEDoc.ReadOnly Then
				ExitCode = 3
			Else
				SEDoc.Save()
				SEApp.DoIdle()
			End If
		End If

		Console.WriteLine("FitIsoView complete")

		Return ExitCode
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
