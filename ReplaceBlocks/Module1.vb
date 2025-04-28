Option Strict On

Imports Newtonsoft.Json

Module Module1
    Function Main() As Integer

        Console.WriteLine("ReplaceBlocks starting...")

        Dim ExitStatus As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

        Dim SEApp As SolidEdgeFramework.Application = Nothing
        Dim SEDoc As SolidEdgeDraft.DraftDocument = Nothing

        Dim ErrorMessageList As New List(Of String)

        'Dim Settings As New Dictionary(Of String, String)
        'Settings = GetSettings()

        Try
            SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            SEDoc = CType(SEApp.ActiveDocument, SolidEdgeDraft.DraftDocument)
        Catch ex As Exception
            ExitStatus = 1
            ErrorMessageList.Add("Error connecting to Solid Edge, or no draft file open")
        End Try

        Dim ReplacementsDict As New Dictionary(Of String, String)
        Dim TemplateName As String = ""
        Dim FileBlockName As String
        Dim TemplateBlockName As String
        Dim TemplateDoc As SolidEdgeDraft.DraftDocument = Nothing
        Dim DocBlocksDict As New Dictionary(Of String, SolidEdgeDraft.Block)
        Dim TemplateBlocksDict As New Dictionary(Of String, SolidEdgeDraft.Block)
        Dim ProgramSettings As Dictionary(Of String, List(Of String))

        ' Program settings
        If ExitStatus = 0 Then
            ProgramSettings = GetProgramSettings()
            If ProgramSettings IsNot Nothing Then
                For Each Key As String In ProgramSettings.Keys
                    If Key.ToLower.Contains("templatename") Then
                        TemplateName = ProgramSettings(Key)(0)
                    ElseIf Key.ToLower.Contains("replaceblock") Then
                        FileBlockName = ProgramSettings(Key)(0)
                        TemplateBlockName = ProgramSettings(Key)(1)
                        ReplacementsDict(FileBlockName) = TemplateBlockName
                    End If
                Next
            Else
                ExitStatus = 1
                ErrorMessageList.Add("Unable to parse program settings file")
            End If
        End If

        ' Template
        If ExitStatus = 0 Then
            If System.IO.File.Exists(TemplateName) Then
                Try
                    TemplateDoc = CType(SEApp.Documents.Open(TemplateName), SolidEdgeDraft.DraftDocument)
                Catch ex As Exception
                    ExitStatus = 1
                    ErrorMessageList.Add(String.Format("Unable to open template '{0}'", TemplateName))
                End Try
            Else
                ExitStatus = 1
                ErrorMessageList.Add(String.Format("Template not found '{0}'", TemplateName))
            End If
        End If

        ' Find and replace
        If ExitStatus = 0 Then

            ' Read all blocks in both the file and template
            ' Populate two dicts such that Dict(BlockName) = Block Object
            For Each DocBlock As SolidEdgeDraft.Block In SEDoc.Blocks
                DocBlocksDict(DocBlock.Name) = DocBlock
            Next
            For Each TemplateBlock As SolidEdgeDraft.Block In TemplateDoc.Blocks
                TemplateBlocksDict(TemplateBlock.Name) = TemplateBlock
            Next

            ' Do the replacements
            For Each DocBlockName As String In ReplacementsDict.Keys
                TemplateBlockName = ReplacementsDict(DocBlockName)
                If DocBlocksDict.Keys.Contains(DocBlockName) Then
                    If TemplateBlocksDict.Keys.Contains(TemplateBlockName) Then
                        If Not DocBlockName = TemplateBlockName Then
                            If DocBlocksDict.Keys.Contains(TemplateBlockName) Then
                                ExitStatus = 1
                                Dim s = String.Format("Cannot replace '{0}' with '{1}'.  ", DocBlockName, TemplateBlockName)
                                s = String.Format("{0}A block with that name already exists in the file.", s)
                                If Not ErrorMessageList.Contains(s) Then ErrorMessageList.Add(s)
                                Continue For
                            Else
                                DocBlocksDict(DocBlockName).Name = TemplateBlockName
                            End If
                        End If
                        SEDoc.Blocks.ReplaceBlock(TemplateBlocksDict(TemplateBlockName))
                    Else
                        ExitStatus = 1
                        ErrorMessageList.Add(String.Format("Template does not have a block named '{0}'", TemplateBlockName))
                    End If
                Else
                    ' Not an error
                End If
            Next
        End If

        If TemplateDoc IsNot Nothing Then TemplateDoc.Close(False)
        SEApp.DoIdle()

        If SEDoc.ReadOnly Then
            ExitStatus = 1
            ErrorMessageList.Add("Cannot save read-only document")
        End If

        If ExitStatus = 0 Then
            SEDoc.Save()
            SEApp.DoIdle()
        Else
            SaveErrorMessages(ErrorMessageList)
        End If

        Console.WriteLine("ReplaceBlocks complete")

        Return ExitStatus
    End Function

    Private Function GetProgramSettings() As Dictionary(Of String, List(Of String))
        Dim ProgramSettings As New Dictionary(Of String, List(Of String))
        Dim Settings As List(Of String) = Nothing
        Dim Key As String
        Dim Value As String
        Dim ProgramSettingsFilename As String
        Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory
        Dim tf As Boolean
        Dim RequiredKeys As List(Of String) = {"TemplateName", "ReplaceBlock"}.ToList

        ProgramSettingsFilename = String.Format("{0}program_settings.txt", StartupPath)

        Try
            Settings = IO.File.ReadAllLines(ProgramSettingsFilename).ToList

            Dim Count As Integer = 0

            For Each KVPair As String In Settings

                Dim s As String = KVPair.Trim()

                tf = s = ""
                tf = tf OrElse s(0) = "'"
                tf = tf OrElse Not s.Contains("=")

                If tf Then Continue For

                tf = s.ToLower.Contains("replaceblock")
                tf = tf And Not s.Split(CChar(",")).Count = 2

                If tf Then Throw New Exception(String.Format("Could not parse '{0}'", s))

                ' Expected format example
                ' TemplateName = C:\Program Files\Siemens\Solid Edge 2024\Template\ANSI Inch\A_sheet.dft
                ' ReplaceBlock = Old name A, New name A

                Key = s.Split("="c)(0).Trim                                  ' 'ReplaceBlock = Old name A, New name A' -> 'ReplaceBlock'

                If Key.ToLower = "templatename" Then
                    Value = s.Split("="c)(1).Trim
                    ProgramSettings(Key) = {Value}.ToList

                ElseIf Key.ToLower.Contains("replaceblock") Then
                    Count += 1

                    Key = String.Format("{0}{1}", Key, CStr(Count))          ' 'ReplaceBlock' -> 'ReplaceBlock1'

                    Value = s.Split("="c)(1).Trim                            ' 'ReplaceBlock = Old name A, New name A' -> 'Old name A, New name A'
                    Dim FileBlockName = Value.Split(CChar(","))(0).Trim      ' 'Old name A, New name A' -> 'Old name A'
                    Dim TemplateBlockName = Value.Split(CChar(","))(1).Trim  ' 'Old name A, New name A' -> 'New name A'

                    ProgramSettings(Key) = {FileBlockName, TemplateBlockName}.ToList
                End If

            Next

        Catch ex As Exception
            MsgBox(String.Format("Problem reading {0}: {1}", ProgramSettingsFilename, ex.Message), vbOKOnly)
            Return Nothing
        End Try

        Dim s1 As String = ""
        For Each s As String In RequiredKeys
            Dim GotAMatch As Boolean = False
            For Each Key In ProgramSettings.Keys
                If Key.Contains(s) Then
                    GotAMatch = True
                    Exit For
                End If
            Next
            If Not GotAMatch Then
                s1 = String.Format("    {0}{1}{2}", s1, s, vbCrLf)
            End If
        Next

        If Not s1 = "" Then
            s1 = String.Format("The following variable names not found in program_settings.txt{0}{1}", vbCrLf, s1)
            MsgBox(s1, vbOKOnly)
            Return Nothing
        End If

        Return ProgramSettings

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
