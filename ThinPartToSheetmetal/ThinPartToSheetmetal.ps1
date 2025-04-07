$StartupPath = Split-Path $script:MyInvocation.MyCommand.Path

$Source = @"
Imports System
Imports System.Collections.Generic
Public Class Snippet

    Public Shared Function RunSnippet(StartupPath As String) As Integer
        Dim ExitStatus As Integer = 0
        Dim ErrorMessageList As New List(Of String)

        Dim SEApp As Object = Nothing
        Dim SEDoc As Object = Nothing

        Try
            SEApp = Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")
            SEDoc = SEApp.ActiveDocument
            Console.WriteLine(String.Format("Processing {0}", SEDoc.Name))
        Catch ex As Exception
            ExitStatus = 1
            ErrorMessageList.Add("Unable to connect to Solid Edge, or no file is open")
        End Try

        If ExitStatus = 0 Then

            Dim DocType = IO.Path.GetExtension(SEDoc.Fullname)

            Try
                Dim Models = SEDoc.Models

                If Models.Count = 0 Then
                    ExitStatus = 1
                    ErrorMessageList.Add("No models detected.")
                End If

                If Models.Count > 1 Then
                    ExitStatus = 1
                    ErrorMessageList.Add("Cannot process files with multiple models.")
                End If

                If Models.Count = 1 Then
                    Dim ConvToSMs = Models.Item(1).ConvToSMs

                    If ConvToSMs.Count > 0 Then
                        ExitStatus = 1
                        ErrorMessageList.Add("Thin part already converted to sheetmetal.")
                    Else
                        Dim Body = Models.Item(1).Body
                        Dim igQueryPlane = 6
                        Dim Faces = Body.Faces(FaceType:=igQueryPlane)
                        Dim Face As Object = Nothing

                        Dim MaxArea As Double = 0
                        For i As Integer = 1 To Faces.Count
                            If Faces(i).Area > MaxArea Then
                                MaxArea = Faces(i).Area
                                Face = Faces(i)
                            End If
                        Next i

                        If Face IsNot Nothing Then
                            Try
                                Dim ConvToSM = ConvToSMs.AddEx(Face)
                                SEApp.DoIdle()

                                If ConvToSM IsNot Nothing Then
                                    Dim Status = ConvToSM.Status
                                    Dim StatusOK as Integer = 1216476310
                                    If Not Status = StatusOK Then
                                        ExitStatus = 1
                                        ErrorMessageList.Add("Possible error in conversion.  Please verify results.")
                                    Else
                                        SEDoc.Save()
                                        SEApp.DoIdle()
                                    End If
                                Else
                                    ExitStatus = 1
                                    ErrorMessageList.Add("Unable to convert to sheetmetal")
                                End If

                            Catch ex As Exception
                                ExitStatus = 1
                                ErrorMessageList.Add("Unable to convert to sheetmetal")
                            End Try
                
                        End If
                    End If
                
                End If
            Catch ex As Exception
                ExitStatus = 1
                ErrorMessageList.Add(String.Format("{0}", ex.Message))
            End Try
        End If

        If Not ExitStatus = 0 Then
            SaveErrorMessages(StartupPath, ErrorMessageList)
        End If

        Return ExitStatus
    End Function

    Private Shared Sub SaveErrorMessages(StartupPath As String, ErrorMessageList As List(Of String))
        Dim ErrorFilename As String

        ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

        IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

    End Sub

End Class
"@

Add-Type -TypeDefinition $Source -Language VisualBasic

$ExitStatus = [Snippet]::RunSnippet($StartupPath)

Function ExitWithCode($exitcode) {
  $host.SetShouldExit($exitcode)
  Exit $exitcode
}

ExitWithCode($ExitStatus)
