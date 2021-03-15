Option Strict On
Module Module1

	Function Main() As Integer

		Console.WriteLine("FitIsoView starting...")

		Dim ExitCode As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.  Edit as required.

		Dim SEApp As SolidEdgeFramework.Application
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument

		SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
		SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)

		Try
			If SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igAssemblyDocument Then
				SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
				SEApp.StartCommand(CType(SolidEdgeConstants.AssemblyCommandConstants.AssemblyViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))

			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igPartDocument Then
				SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
				SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewFit, SolidEdgeFramework.SolidEdgeCommandConstants))

			ElseIf SEApp.ActiveDocumentType = SolidEdgeFramework.DocumentTypeConstants.igSheetMetalDocument Then
				SEApp.StartCommand(CType(SolidEdgeConstants.PartCommandConstants.PartViewISOView, SolidEdgeFramework.SolidEdgeCommandConstants))
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

End Module
