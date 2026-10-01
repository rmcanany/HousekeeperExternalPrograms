Option Strict On

Imports System

Module Program

	Function Main() As Integer

		Console.WriteLine("CheckDisplayConfigurations starting...")

		Dim ExitCode As Integer = 0  ' 0 means success.  For a more complete example, see FitISOView.
		Dim ErrorMessageList As New List(Of String)

		Dim Proceed As Boolean = True

		Dim SEApp As SolidEdgeFramework.Application = Nothing
		Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing

		If Proceed Then
			Try
				SEApp = CType(MarshalHelper.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
				SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
			Catch ex As Exception
				Proceed = False
				ErrorMessageList.Add("Error connecting to Solid Edge")
			End Try
		End If

		Dim Filename As String = ""

		If Proceed Then
			If SEDoc.FullName.Contains("!") Then
				Proceed = False
				ErrorMessageList.Add("Cannot currently process FOA files")
			Else
				Filename = SEDoc.FullName.Split("!"c)(0)  ' Deal with FOA name
				Dim DocType As String = IO.Path.GetExtension(Filename)

				If Not DocType = ".asm" Then Proceed = False
			End If
		End If

		If Proceed Then
			Dim tmpSEDoc As SolidEdgeAssembly.AssemblyDocument = CType(SEDoc, SolidEdgeAssembly.AssemblyDocument)

			CheckConfigurationsAgainstCurrentOccurrences(tmpSEDoc, Filename, ErrorMessageList)
		End If

		If ErrorMessageList.Count > 0 Then
			ExitCode = 1
			SaveErrorMessages(ErrorMessageList)
		End If

		Console.WriteLine("CheckDisplayConfigurations complete")

		Return ExitCode
	End Function

	' 20260930 Three separate automation attempts (Apply() alone, Apply()+DoIdle(),
	' Apply()+Update()+DoIdle()) never once produced the "parts deleted" warning,
	' confirmed via a Spy++-equivalent window dump that the real dialog (class
	' #32770, title "Warning") does exist and that the window-matching logic
	' below was correct -- so the check apparently lives in the interactive
	' "Display Configurations" command's own button-click handler, not in any
	' COM-exposed Configuration method.  A community forum AutoHotKey script for
	' *other* Solid Edge warnings (relationship conflicts, hole fit class, frame
	' cross-sections, etc.) confirms this is a known, recurring category of
	' problem, not specific to this one check:
	' https://community.sw.siemens.com/s/question/0D5Vb00001RdjfCKAR
	'
	' Replaced by CheckConfigurationsAgainstCurrentOccurrences below, which
	' replicates the check directly from the saved configurations in the .cfg
	' file instead of chasing a dialog that automation may not be able to
	' trigger at all.  (A first attempt using Configurations.GetConfigComponentList
	' didn't work either -- it returns the configuration's current, already
	' corrected list, not the stale saved one.)  Left
	' here, commented out, since watching for and dismissing a Solid Edge dialog
	' this way is a real, tested technique that may be exactly what one of
	' those *other* warnings needs someday.
	'Private Function ApplyConfigurationAndCheckIfOutOfDate(SEApp As SolidEdgeFramework.Application, Configuration As SolidEdgeAssembly.Configuration, EdgeProcessIds As List(Of Integer)) As Boolean
	'
	'	Dim StopWatching As Boolean = False
	'	Dim DialogDismissed As Boolean = False
	'
	'	Dim Watcher As New System.Threading.Thread(
	'		Sub()
	'			While Not StopWatching
	'				If TryDismissConfigurationWarning(EdgeProcessIds) Then
	'					DialogDismissed = True
	'					Exit While
	'				End If
	'				System.Threading.Thread.Sleep(50)
	'			End While
	'		End Sub)
	'	Watcher.IsBackground = True
	'	Watcher.Start()
	'
	'	Try
	'		Configuration.Apply()
	'		Configuration.Update()
	'
	'		Dim WaitedMilliseconds As Integer = 0
	'		While Not DialogDismissed AndAlso WaitedMilliseconds < 3000
	'			SEApp.DoIdle()
	'			System.Threading.Thread.Sleep(50)
	'			WaitedMilliseconds += 50
	'		End While
	'	Finally
	'		StopWatching = True
	'		Watcher.Join(1000)
	'	End Try
	'
	'	Return DialogDismissed
	'
	'End Function
	'
	'Private Function TryDismissConfigurationWarning(EdgeProcessIds As List(Of Integer)) As Boolean
	'
	'	For Each hWnd As IntPtr In GetTopLevelWindows()
	'
	'		If Not IsWindowVisible(hWnd) Then Continue For
	'
	'		Dim WindowProcessId As Integer = 0
	'		GetWindowThreadProcessId(hWnd, WindowProcessId)
	'		If Not EdgeProcessIds.Contains(WindowProcessId) Then Continue For
	'
	'		If Not GetClassNameSafe(hWnd) = "#32770" Then Continue For
	'		If Not GetWindowTextSafe(hWnd) = "Warning" Then Continue For
	'
	'		Dim ChildWindows = GetChildWindows(hWnd)
	'
	'		Dim MessageText As String = ""
	'		For Each hChild As IntPtr In ChildWindows
	'			If GetClassNameSafe(hChild) = "Static" Then
	'				MessageText = $"{MessageText}{GetWindowTextSafe(hChild)} "
	'			End If
	'		Next
	'
	'		If Not MessageText.Contains("deleted from the assembly") Then Continue For
	'
	'		For Each hChild As IntPtr In ChildWindows
	'			If GetClassNameSafe(hChild) = "Button" AndAlso GetWindowTextSafe(hChild) = "OK" Then
	'				SendMessage(hChild, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
	'				Return True
	'			End If
	'		Next
	'
	'	Next
	'
	'	Return False
	'
	'End Function
	'
	'Private Const BM_CLICK As Integer = &HF5
	'
	'Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll")>
	'Private Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
	'End Function
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll")>
	'Private Function EnumChildWindows(hWndParent As IntPtr, lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
	'End Function
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll")>
	'Private Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
	'End Function
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
	'Private Function GetClassName(hWnd As IntPtr, lpClassName As System.Text.StringBuilder, nMaxCount As Integer) As Integer
	'End Function
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
	'Private Function GetWindowText(hWnd As IntPtr, lpString As System.Text.StringBuilder, nMaxCount As Integer) As Integer
	'End Function
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll")>
	'Private Function IsWindowVisible(hWnd As IntPtr) As Boolean
	'End Function
	'
	'<System.Runtime.InteropServices.DllImport("user32.dll", CharSet:=System.Runtime.InteropServices.CharSet.Auto)>
	'Private Function SendMessage(hWnd As IntPtr, Msg As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
	'End Function
	'
	'Private Function GetTopLevelWindows() As List(Of IntPtr)
	'	Dim Result As New List(Of IntPtr)
	'	EnumWindows(Function(hWnd, lParam)
	'					Result.Add(hWnd)
	'					Return True
	'				End Function, IntPtr.Zero)
	'	Return Result
	'End Function
	'
	'Private Function GetChildWindows(hWndParent As IntPtr) As List(Of IntPtr)
	'	Dim Result As New List(Of IntPtr)
	'	EnumChildWindows(hWndParent, Function(hWnd, lParam)
	'									 Result.Add(hWnd)
	'									 Return True
	'								 End Function, IntPtr.Zero)
	'	Return Result
	'End Function
	'
	'Private Function GetWindowTextSafe(hWnd As IntPtr) As String
	'	Dim SB As New System.Text.StringBuilder(256)
	'	GetWindowText(hWnd, SB, SB.Capacity)
	'	Return SB.ToString()
	'End Function
	'
	'Private Function GetClassNameSafe(hWnd As IntPtr) As String
	'	Dim SB As New System.Text.StringBuilder(256)
	'	GetClassName(hWnd, SB, SB.Capacity)
	'	Return SB.ToString()
	'End Function

	' Compares each configuration's saved occurrence tree, read from the
	' assembly's .cfg file, against the assembly's current occurrences, matching
	' by OccurrenceID level by level.  A saved ID with no current occurrence is a
	' part deleted since the configuration was saved.  See ConfigurationFile.vb
	' for the file format.
	'
	' Each out-of-date configuration gets one message, worded like Solid Edge's
	' own warning, since which configuration needs re-saving is what the user can
	' act on.  The details -- both sides, and which IDs are missing where -- go to
	' configuration_diagnostic.txt for troubleshooting.
	Private Sub CheckConfigurationsAgainstCurrentOccurrences(
		tmpSEDoc As SolidEdgeAssembly.AssemblyDocument,
		Filename As String,
		ErrorMessageList As List(Of String))

		Dim CfgFilename As String = IO.Path.ChangeExtension(Filename, ".cfg")

		' Solid Edge normally creates the .cfg file along with the assembly, so its
		' absence is unusual enough to report.
		If Not IO.File.Exists(CfgFilename) Then
			ErrorMessageList.Add($"Configuration file not found: '{IO.Path.GetFileName(CfgFilename)}'")
			Exit Sub
		End If

		Dim Configurations As SortedDictionary(Of String, ConfigurationNode)
		Dim ConfigurationErrors As New SortedDictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
		Try
			Configurations = ConfigurationFile.ReadConfigurations(CfgFilename, ConfigurationErrors)
		Catch ex As Exception
			ErrorMessageList.Add($"Could not read '{IO.Path.GetFileName(CfgFilename)}': {ex.Message}")
			Exit Sub
		End Try

		For Each Item In ConfigurationErrors
			ErrorMessageList.Add($"Could not check display configuration '{Item.Key}': {Item.Value}")
		Next

		' Occurrences by ID and name for each assembly document, so a subassembly
		' used in several places, or checked for several configurations, is only
		' walked once.
		Dim OccurrenceCache As New Dictionary(Of String, OccurrenceIndex)(StringComparer.OrdinalIgnoreCase)

		Dim DiagnosticLines As New List(Of String)
		DiagnosticLines.Add($"Assembly: {Filename}")

		For Each Item In ConfigurationErrors
			DiagnosticLines.Add("")
			DiagnosticLines.Add($"Configuration '{Item.Key}'")
			DiagnosticLines.Add($"  not checked: {Item.Value}")
		Next

		' Subassembly files whose contents couldn't be compared, eg. because the
		' file is missing, with the reason.  Not treated as missing occurrences,
		' which would report a configuration out of date when it may not be.
		' Reported once per file -- that's what the user can fix -- rather than once
		' per instance and configuration.  Where each one is used is in the
		' diagnostic file.
		Dim UncheckedSubassemblies As New SortedDictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

		For Each Item In Configurations
			DiagnosticLines.Add("")
			DiagnosticLines.Add($"Configuration '{Item.Key}'")

			Dim MissingEntries As New List(Of String)
			CompareConfigurationNode(Item.Value, tmpSEDoc, "", "    ", OccurrenceCache, MissingEntries, UncheckedSubassemblies, DiagnosticLines)

			If MissingEntries.Count > 0 Then
				ErrorMessageList.Add($"Display configuration '{Item.Key}' out of date")
				DiagnosticLines.Add($"  {MissingEntries.Count} missing: {String.Join(", ", MissingEntries)}")
			End If
		Next

		For Each Item In UncheckedSubassemblies
			ErrorMessageList.Add($"Could not check subassembly '{Item.Key}': {Item.Value}")
		Next

		IO.File.WriteAllLines(String.Format("{0}\configuration_diagnostic.txt", System.AppDomain.CurrentDomain.BaseDirectory), DiagnosticLines)

	End Sub

	' What OccurrenceDocument throws when the occurrence's file is missing.
	Private Const STG_E_FILENOTFOUND As Integer = &H80030002

	Private Sub CompareConfigurationNode(
		Node As ConfigurationNode,
		AssemblyDoc As SolidEdgeAssembly.AssemblyDocument,
		ParentPath As String,
		Indent As String,
		OccurrenceCache As Dictionary(Of String, OccurrenceIndex),
		MissingEntries As List(Of String),
		UncheckedSubassemblies As SortedDictionary(Of String, String),
		DiagnosticLines As List(Of String))

		Dim CurrentOccurrences = GetOccurrenceIndex(AssemblyDoc, OccurrenceCache)

		For Each Child As ConfigurationNode In Node.Children
			Dim Visibility As String = If(Child.IsHidden, "hidden", "shown")
			If Child.IsCollapsed Then Visibility &= ", collapsed"

			' Older (version 4) configurations identify occurrences by name, newer ones by ID.
			Dim Occ As SolidEdgeAssembly.Occurrence = Nothing
			Dim Found As Boolean
			Dim Description As String
			If Child.OccurrenceName IsNot Nothing Then
				Found = CurrentOccurrences.ByName.TryGetValue(Child.OccurrenceName, Occ)
				Description = $"occurrence '{Child.OccurrenceName}'"
			Else
				Found = CurrentOccurrences.ByID.TryGetValue(Child.OccurrenceID, Occ)
				Description = $"occurrence ID {Child.OccurrenceID}"
			End If

			If Not Found Then
				MissingEntries.Add($"{ParentPath}{Description}")
				DiagnosticLines.Add($"{Indent}{Description} ({Visibility}) -- MISSING")
				Continue For
			End If

			DiagnosticLines.Add($"{Indent}{Description} ({Visibility}) = {Occ.Name}")

			If Child.Children.Count > 0 AndAlso Occ.Subassembly Then
				' The Try covers only getting the document, so an error at a deeper
				' level isn't mistaken for this subassembly being unavailable.
				Dim SubDoc As SolidEdgeAssembly.AssemblyDocument = Nothing
				Dim Reason As String = "document not available"
				Try
					SubDoc = TryCast(Occ.OccurrenceDocument, SolidEdgeAssembly.AssemblyDocument)
					If SubDoc IsNot Nothing AndAlso SubDoc.Occurrences Is Nothing Then
						SubDoc = Nothing
						Reason = "occurrences not available"
					End If
				Catch ex As System.Runtime.InteropServices.COMException When ex.HResult = STG_E_FILENOTFOUND
					Reason = "file not found"
				Catch ex As Exception
					Reason = ex.Message.Trim()
				End Try

				If SubDoc Is Nothing Then
					Dim SubFilename As String = Occ.OccurrenceFileName
					If String.IsNullOrEmpty(SubFilename) Then SubFilename = Occ.Name
					UncheckedSubassemblies(IO.Path.GetFileName(SubFilename)) = Reason
					DiagnosticLines.Add($"{Indent}    (not checked: {Reason}: {SubFilename})")
				Else
					CompareConfigurationNode(Child, SubDoc, $"{ParentPath}{Occ.Name} > ", Indent & "    ", OccurrenceCache, MissingEntries, UncheckedSubassemblies, DiagnosticLines)
				End If
			End If
		Next

	End Sub

	Private Class OccurrenceIndex
		Public ByID As New Dictionary(Of Integer, SolidEdgeAssembly.Occurrence)
		Public ByName As New Dictionary(Of String, SolidEdgeAssembly.Occurrence)(StringComparer.OrdinalIgnoreCase)
	End Class

	Private Function GetOccurrenceIndex(
		AssemblyDoc As SolidEdgeAssembly.AssemblyDocument,
		OccurrenceCache As Dictionary(Of String, OccurrenceIndex)
		) As OccurrenceIndex

		Dim Result As OccurrenceIndex = Nothing
		If OccurrenceCache.TryGetValue(AssemblyDoc.FullName, Result) Then Return Result

		Result = New OccurrenceIndex
		For Each Occ As SolidEdgeAssembly.Occurrence In AssemblyDoc.Occurrences
			Result.ByID(Occ.OccurrenceID) = Occ
			Result.ByName(Occ.Name) = Occ
		Next

		OccurrenceCache(AssemblyDoc.FullName) = Result
		Return Result
	End Function


	Private Sub SaveErrorMessages(ErrorMessageList As List(Of String))
		Dim ErrorFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory

		ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

		IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

	End Sub

End Module
