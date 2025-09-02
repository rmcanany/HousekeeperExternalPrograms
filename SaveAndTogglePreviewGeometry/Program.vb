'Contributed by @ih0nza

Option Strict On

Imports Newtonsoft.Json

Module Module1
    Function Main() As Integer

        Console.WriteLine("SaveAndTogglePreviewGeometry starting...")

        Dim ExitStatus As Integer = 0  ' 0 means success.  Error messages are stored in error_messages.txt.

        Dim SEApp As SolidEdgeFramework.Application = Nothing
        Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim DocType As String

        Dim ErrorMessageList As New List(Of String)

        'Dim Settings As New Dictionary(Of String, String)
        'Settings = GetSettings()
        Try
            SEApp = CType(MarshalHelper.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
        Catch ex As Exception
            ExitStatus = 1
            ErrorMessageList.Add("Error connecting to Solid Edge, or no file open")
        End Try

        DocType = GetDocType(SEDoc)

        If DocType = "asm" Then
            Dim PreviewParam As SolidEdgeFramework.ApplicationGlobalConstants
            PreviewParam = SolidEdgeFramework.ApplicationGlobalConstants.seApplicationGlobalStoreGeometryInAssemblyForPreview

            Try
                ' Krok 1 Vypnutí a uložení
                Console.WriteLine("Disabling preview mode")
                SEApp.SetGlobalParameter(PreviewParam, False)
                Console.WriteLine("Saving file")
                SEDoc.Save()
                SEApp.DoIdle()
            Catch ex As Exception
                ExitStatus = 1
                Console.WriteLine(String.Format("An error occurred disabling preview mode: {0}", ex.Message))
            End Try

            Try
                ' Krok 2 Zapnutí a uložení
                Console.WriteLine("Enabling preview mode")
                SEApp.SetGlobalParameter(PreviewParam, True)
                Console.WriteLine("Saving file")
                SEDoc.Save()
                SEApp.DoIdle()
            Catch ex As Exception
                ExitStatus = 1
                Console.WriteLine(String.Format("An error occurred enabling preview mode: {0}", ex.Message))
            End Try

        Else
            Try
                Console.WriteLine("Saving file")
                SEDoc.Save()
                SEApp.DoIdle()
            Catch ex As Exception
                ExitStatus = 1
                Console.WriteLine(String.Format("An error occurred saving the file: {0}", ex.Message))
            End Try
        End If


        Console.WriteLine("SaveAndTogglePreviewGeometry complete")

        Return ExitStatus
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
