Option Strict On

' 20261001 Reads the <assembly>.cfg file Solid Edge keeps next to an assembly.
' It is a structured storage file with one stream per display configuration
' under the "Configs" storage, plus an "Info" stream naming the active one.
'
' Each configuration stream is a header followed by a tree of nodes, all
' little-endian UInt32 values except for names.  The header is:
'
'   ConfigurationType, Version, (counter), 0, [SubFormat -- version 5 only]
'
' The configuration type is described at ParseConfiguration.  The layout of a
' node depends on the version:
'
'   Version 5, subformat 2 (current):  Flags, Flags2, OccurrenceID
'   Version 5, subformat 1:            Flags, OccurrenceID
'   Version 4:                         Flags, 0, then for named nodes NameLength, Name (ASCII, null-terminated)
'
'   Assembly node:     Node, EntryCount, EntryCount x 12 bytes, ChildCount, ChildCount x child nodes
'   Collapsed node:    Node, ChildCount, ChildCount x child nodes
'
' Flags >> 28 = 9 marks an assembly node (the root, ID 0, is the assembly itself),
' 4 a part.  Version 4 identifies occurrences by name (eg. "bearing1.par:1")
' rather than ID: B is a named assembly node, 6 a named part.  Flags bit 2 means
' hidden in that configuration (version 5).  The 12-byte entries look like the
' document's own objects (reference planes, coordinate systems, etc.) and their
' visibility; they're not needed here and are skipped.
'
' 20261001 Versions 4 and 5/1 were found in files from older Solid Edge versions,
' alongside per-user "default,<username>" configurations.  Version 2, older
' still, has no node tree -- entries look like paths of storage names
' ("Layouts", "References") -- and isn't supported.
'
' 20261001 Flags >> 28 = 8 (seen as &H80000002) is a hidden subassembly stored
' without its contents.  Every one seen so far had a count of 0; treating a
' nonzero count as child nodes is a guess, but the full-length check below
' would catch it if wrong.
'
' Flags2 is almost always 0.  The one exception seen so far was &H00080000 on a
' part inside a since-deleted subassembly, so its meaning is unknown; it isn't
' needed for the check.
'
' Occurrence IDs (and names) are scoped per document -- a subassembly node's
' children are IDs within that subassembly, matching Occurrence.OccurrenceID
' (or Occurrence.Name) there.  The tree
' is the configuration as it was last saved, so an ID that no longer matches a
' current occurrence is a part deleted since, which is what triggers Solid Edge's
' "One or more parts have been deleted..." warning when the configuration is
' activated.  The "default,Solid Edge" configuration is rewritten on every save.

Public Class ConfigurationNode
	Public Property OccurrenceID As Integer

	' Set instead of OccurrenceID for version 4 configurations, otherwise Nothing.
	Public Property OccurrenceName As String

	Public Property Flags As UInteger
	Public Property Flags2 As UInteger
	Public Property Children As New List(Of ConfigurationNode)

	Public ReadOnly Property IsAssembly As Boolean
		Get
			Return (Flags >> 28) = 9UI OrElse (Flags >> 28) = &HBUI
		End Get
	End Property

	Public ReadOnly Property IsNamed As Boolean
		Get
			Return (Flags >> 28) = 6UI OrElse (Flags >> 28) = &HBUI
		End Get
	End Property

	Public ReadOnly Property IsCollapsed As Boolean
		Get
			Return (Flags >> 28) = 8UI
		End Get
	End Property

	Public ReadOnly Property IsHidden As Boolean
		Get
			Return (Flags And 2UI) <> 0UI
		End Get
	End Property
End Class

Public Module ConfigurationFile

	Private Const EntryLength As Integer = 12
	Private Const MinimumNodeLength As Integer = 8

	' How nodes are laid out in a given configuration stream; see the top of the file.
	Private Class NodeFormat
		Public Version As UInteger
		Public HasFlags2 As Boolean
	End Class

	' Returns each configuration's saved tree, keyed by configuration name.  The
	' .cfg file is locked while the assembly is open in Solid Edge, so a temporary
	' copy is read instead.  A configuration that can't be parsed goes to
	' ConfigurationErrors (name -> reason) instead, so one unreadable
	' configuration doesn't keep the rest from being checked.
	Public Function ReadConfigurations(
		CfgFilename As String,
		ConfigurationErrors As SortedDictionary(Of String, String)
		) As SortedDictionary(Of String, ConfigurationNode)

		Dim Result As New SortedDictionary(Of String, ConfigurationNode)(StringComparer.OrdinalIgnoreCase)

		Dim TempFilename As String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"CheckDisplayConfigurations_{System.Guid.NewGuid():N}.cfg")
		System.IO.File.Copy(CfgFilename, TempFilename)

		Try
			For Each Item In ReadConfigStreams(TempFilename)
				Try
					Result(Item.Key) = ParseConfiguration(Item.Value)
				Catch ex As Exception
					ConfigurationErrors(Item.Key) = ex.Message
				End Try
			Next
		Finally
			Try
				System.IO.File.Delete(TempFilename)
			Catch
				' Leftover temp file isn't worth failing the check over.
			End Try
		End Try

		Return Result
	End Function

	' 20261001 The first header value is the configuration type: 0 for an ordinary
	' display configuration, 1 for an exploded view.  An exploded view has the
	' same occurrence tree, followed by the explosion data (events, groups,
	' transforms), which isn't needed here and is skipped.
	Private Const ConfigurationTypeDisplay As UInteger = 0UI
	Private Const ConfigurationTypeExploded As UInteger = 1UI

	Private Function ParseConfiguration(Bytes As Byte()) As ConfigurationNode
		Dim Position As Integer = 0
		Dim ConfigurationType As UInteger = ReadUInt32(Bytes, Position)

		If Not (ConfigurationType = ConfigurationTypeDisplay OrElse ConfigurationType = ConfigurationTypeExploded) Then
			Throw New System.IO.InvalidDataException($"Unrecognized configuration type {ConfigurationType}")
		End If

		Dim Format As New NodeFormat
		Format.Version = ReadUInt32(Bytes, Position)
		ReadUInt32(Bytes, Position)  ' Looks like a count of times this configuration was saved; not needed
		ReadUInt32(Bytes, Position)

		Select Case Format.Version
			Case 4UI
				Format.HasFlags2 = False
			Case 5UI
				Dim SubFormat As UInteger = ReadUInt32(Bytes, Position)
				Select Case SubFormat
					Case 1UI
						Format.HasFlags2 = False
					Case 2UI
						Format.HasFlags2 = True
					Case Else
						Throw New System.IO.InvalidDataException($"Unrecognized format version {Format.Version}, subformat {SubFormat}")
				End Select
			Case 2UI
				' Applying the configuration in a current Solid Edge and saving the
				' assembly rewrites it in the current format (seen with SE2025).
				Throw New System.IO.InvalidDataException("Older format (version 2) not supported; apply the configuration and save the assembly to update it")
			Case Else
				Throw New System.IO.InvalidDataException($"Unrecognized format version {Format.Version}")
		End Select

		Dim Root As ConfigurationNode = ParseNode(Bytes, Position, Format)

		' Every ordinary configuration seen so far parses to exactly its full length.
		' Anything else means the format isn't what we think it is, and the result
		' can't be trusted.
		If ConfigurationType = ConfigurationTypeDisplay AndAlso Position <> Bytes.Length Then
			Throw New System.IO.InvalidDataException($"Unrecognized format (parsed {Position} of {Bytes.Length} bytes)")
		End If

		Return Root
	End Function

	Private Function ParseNode(Bytes As Byte(), ByRef Position As Integer, Format As NodeFormat) As ConfigurationNode
		Dim NodePosition As Integer = Position

		Dim Node As New ConfigurationNode
		Node.Flags = ReadUInt32(Bytes, Position)
		If Format.HasFlags2 Then Node.Flags2 = ReadUInt32(Bytes, Position)
		Dim ID As UInteger = ReadUInt32(Bytes, Position)

		' Fail at the first node that doesn't fit the known format, with enough
		' detail to see what's new, rather than misreading everything after it.
		Dim NodeType As UInteger = Node.Flags >> 28
		Dim KnownType As Boolean
		If Format.Version = 4UI Then
			' Version 4 can identify occurrences by ID too (4, 9), mixed or not with names (6, B).
			KnownType = NodeType = 4UI OrElse NodeType = 6UI OrElse NodeType = 8UI OrElse NodeType = 9UI OrElse NodeType = &HBUI
		Else
			KnownType = NodeType = 4UI OrElse NodeType = 8UI OrElse NodeType = 9UI
		End If
		If Not KnownType OrElse ID > CUInt(Integer.MaxValue) Then
			Throw New System.IO.InvalidDataException(
				$"Unrecognized node at byte offset {NodePosition} (format version {Format.Version}): flags 0x{Node.Flags:X8}, second value 0x{Node.Flags2:X8}, ID 0x{ID:X8}")
		End If

		Node.OccurrenceID = CInt(ID)

		If Node.IsNamed Then
			Dim NameLength As UInteger = ReadUInt32(Bytes, Position)
			If NameLength > CUInt(Bytes.Length - Position) Then
				Throw New System.IO.InvalidDataException($"Name length {NameLength} at byte offset {NodePosition} exceeds stream length")
			End If
			Node.OccurrenceName = System.Text.Encoding.ASCII.GetString(Bytes, Position, CInt(NameLength)).TrimEnd(Chr(0))
			Position += CInt(NameLength)
		End If

		If Node.IsAssembly Then
			Dim EntryCount As UInteger = ReadUInt32(Bytes, Position)
			If EntryCount > CUInt((Bytes.Length - Position) \ EntryLength) Then
				Throw New System.IO.InvalidDataException($"Entry count {EntryCount} exceeds stream length")
			End If
			Position += CInt(EntryCount) * EntryLength

			ParseChildren(Node, Bytes, Position, Format)

		ElseIf Node.IsCollapsed Then
			ParseChildren(Node, Bytes, Position, Format)
		End If

		Return Node
	End Function

	Private Sub ParseChildren(Node As ConfigurationNode, Bytes As Byte(), ByRef Position As Integer, Format As NodeFormat)
		Dim ChildCount As UInteger = ReadUInt32(Bytes, Position)
		If ChildCount > CUInt((Bytes.Length - Position) \ MinimumNodeLength) Then
			Throw New System.IO.InvalidDataException($"Child count {ChildCount} exceeds stream length")
		End If
		For i As Integer = 1 To CInt(ChildCount)
			Node.Children.Add(ParseNode(Bytes, Position, Format))
		Next
	End Sub

	Private Function ReadUInt32(Bytes As Byte(), ByRef Position As Integer) As UInteger
		If Position + 4 > Bytes.Length Then
			Throw New System.IO.InvalidDataException("Unexpected end of stream")
		End If
		Dim Value As UInteger = System.BitConverter.ToUInt32(Bytes, Position)
		Position += 4
		Return Value
	End Function

	' Returns the raw bytes of every stream in the "Configs" storage, keyed by stream name.
	Private Function ReadConfigStreams(Filename As String) As Dictionary(Of String, Byte())

		Dim Result As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)

		Dim Root As IStorage = Nothing
		Dim Configs As IStorage = Nothing

		Dim hr As Integer = StgOpenStorage(Filename, Nothing, STGM_READ_SHARE_EXCLUSIVE, IntPtr.Zero, 0, Root)
		If hr <> 0 Then
			System.Runtime.InteropServices.Marshal.ThrowExceptionForHR(hr)
		End If

		Try
			Try
				Root.OpenStorage("Configs", Nothing, STGM_READ_SHARE_EXCLUSIVE, IntPtr.Zero, 0, Configs)
			Catch ex As System.Runtime.InteropServices.COMException
				' No "Configs" storage means no saved configurations.
				Return Result
			End Try

			Dim Enumerator As IEnumSTATSTG = Nothing
			Configs.EnumElements(0, IntPtr.Zero, 0, Enumerator)

			Dim Stat(0) As System.Runtime.InteropServices.ComTypes.STATSTG
			Dim Fetched As UInteger = 0
			While Enumerator.Next(1, Stat, Fetched) = 0 AndAlso Fetched = 1
				If Stat(0).type <> STGTY_STREAM Then Continue While

				Dim Stream As System.Runtime.InteropServices.ComTypes.IStream = Nothing
				Configs.OpenStream(Stat(0).pwcsName, IntPtr.Zero, STGM_READ_SHARE_EXCLUSIVE, 0, Stream)
				Try
					Dim Bytes(CInt(Stat(0).cbSize) - 1) As Byte
					Stream.Read(Bytes, Bytes.Length, IntPtr.Zero)
					Result(Stat(0).pwcsName) = Bytes
				Finally
					System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Stream)
				End Try
			End While

			System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Enumerator)
		Finally
			' Released explicitly so the temporary copy isn't still open when it's deleted.
			If Configs IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Configs)
			System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Root)
		End Try

		Return Result
	End Function

	Private Const STGM_READ_SHARE_EXCLUSIVE As UInteger = &H10UI
	Private Const STGTY_STREAM As Integer = 2

	<System.Runtime.InteropServices.DllImport("ole32.dll")>
	Private Function StgOpenStorage(
		<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String,
		pstgPriority As IStorage,
		grfMode As UInteger,
		snbExclude As IntPtr,
		reserved As UInteger,
		ByRef ppstgOpen As IStorage) As Integer
	End Function

	' Declared in full because COM calls go by vtable position.
	<System.Runtime.InteropServices.ComImport,
	 System.Runtime.InteropServices.Guid("0000000b-0000-0000-C000-000000000046"),
	 System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
	Private Interface IStorage
		Sub CreateStream(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String, grfMode As UInteger, reserved1 As UInteger, reserved2 As UInteger, ByRef ppstm As System.Runtime.InteropServices.ComTypes.IStream)
		Sub OpenStream(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String, reserved1 As IntPtr, grfMode As UInteger, reserved2 As UInteger, ByRef ppstm As System.Runtime.InteropServices.ComTypes.IStream)
		Sub CreateStorage(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String, grfMode As UInteger, reserved1 As UInteger, reserved2 As UInteger, ByRef ppstg As IStorage)
		Sub OpenStorage(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String, pstgPriority As IStorage, grfMode As UInteger, snbExclude As IntPtr, reserved As UInteger, ByRef ppstg As IStorage)
		Sub CopyTo(ciidExclude As UInteger, rgiidExclude As IntPtr, snbExclude As IntPtr, pstgDest As IStorage)
		Sub MoveElementTo(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String, pstgDest As IStorage, <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsNewName As String, grfFlags As UInteger)
		Sub Commit(grfCommitFlags As UInteger)
		Sub Revert()
		Sub EnumElements(reserved1 As UInteger, reserved2 As IntPtr, reserved3 As UInteger, ByRef ppenum As IEnumSTATSTG)
		Sub DestroyElement(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String)
		Sub RenameElement(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsOldName As String, <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsNewName As String)
		Sub SetElementTimes(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)> pwcsName As String, pctime As IntPtr, patime As IntPtr, pmtime As IntPtr)
		Sub SetClass(ByRef clsid As System.Guid)
		Sub SetStateBits(grfStateBits As UInteger, grfMask As UInteger)
		Sub Stat(ByRef pstatstg As System.Runtime.InteropServices.ComTypes.STATSTG, grfStatFlag As UInteger)
	End Interface

	<System.Runtime.InteropServices.ComImport,
	 System.Runtime.InteropServices.Guid("0000000d-0000-0000-C000-000000000046"),
	 System.Runtime.InteropServices.InterfaceType(System.Runtime.InteropServices.ComInterfaceType.InterfaceIsIUnknown)>
	Private Interface IEnumSTATSTG
		<System.Runtime.InteropServices.PreserveSig>
		Function [Next](celt As UInteger, <System.Runtime.InteropServices.Out, System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray)> rgelt As System.Runtime.InteropServices.ComTypes.STATSTG(), ByRef pceltFetched As UInteger) As Integer
		Sub Skip(celt As UInteger)
		Sub Reset()
		Sub Clone(ByRef ppenum As IEnumSTATSTG)
	End Interface

End Module
