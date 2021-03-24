Option Strict On
Module Module1

	Function Main() As Integer

		Console.WriteLine("AddRemoveCustomProperties starting...")

		Dim ExitCode As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.  Edit as required.

		Dim SEApp As SolidEdgeFramework.Application
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument

        Dim PropertySets As SolidEdgeFramework.PropertySets = Nothing
        Dim Properties As SolidEdgeFramework.Properties = Nothing
        Dim Prop As SolidEdgeFramework.Property = Nothing
		Dim RemoveProps As New List(Of String)

		RemoveProps = GetRemoveProps()

		SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
		SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)

		PropertySets = CType(SEDoc.Properties, SolidEdgeFramework.PropertySets)

		For Each Properties In PropertySets
			If Properties.Name.ToLower = "custom" Then
				For Each Prop In Properties
					If RemoveProps.Contains(Prop.Name) Then
						Console.WriteLine(Prop.Name)
						Try
							Prop.Delete()
						Catch ex As Exception
							ExitCode = 1
						End Try
					End If
				Next
			End If
			Properties.Save()
		Next


		If ExitCode = 0 Then
			If SEDoc.ReadOnly Then
				ExitCode = 2
			Else
				SEDoc.Save()
				SEApp.DoIdle()
			End If
		End If

		Console.WriteLine("AddRemoveCustomProperties complete")

		' Short pause so the user can see the console feedback.
		System.Threading.Thread.Sleep(500)

		Return ExitCode
	End Function

	Function GetRemoveProps() As List(Of String)
		Dim RemoveProps As New List(Of String)
		Dim StartupPath As String = AppDomain.CurrentDomain.BaseDirectory

		Dim Filename As String = String.Format("{0}\properties_to_remove.txt", StartupPath)
		' Console.WriteLine(Filename)

		Try
			RemoveProps = IO.File.ReadAllLines(Filename).ToList
		Catch ex As Exception
			Console.WriteLine(String.Format("Unable to open {0}", Filename))
		End Try

		Return RemoveProps
	End Function
End Module

