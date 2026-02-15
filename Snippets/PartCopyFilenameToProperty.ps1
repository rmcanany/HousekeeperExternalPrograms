$StartupPath = Split-Path $script:MyInvocation.MyCommand.Path

$DLLs = (
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgeFramework.dll",
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgeFrameworkSupport.dll",
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgeConstants.dll",
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgePart.dll",
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgeAssembly.dll",
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgeDraft.dll",
    "C:\data\CAD\scripts\SolidEdgeHousekeeper\bin\Debug\Interop.SolidEdgeGeometry.dll"
    )

$Source = @"

Imports System
Imports System.Collections.Generic

Public Class Snippet

    Public Shared Function RunSnippet(StartupPath As String) As Integer
        Dim ExitStatus As Integer = 0
        Dim ErrorMessageList As New List(Of String)

        Dim SEApp As SolidEdgeFramework.Application = Nothing
        Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing

        Try
            SEApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)
            SEDoc = CType(SEApp.ActiveDocument, SolidEdgeFramework.SolidEdgeDocument)
            Console.WriteLine(String.Format("Processing {0}", SEDoc.Name))
        Catch ex As Exception
            ExitStatus = 1
            ErrorMessageList.Add("Unable to connect to Solid Edge, or no file is open")
        End Try

        If ExitStatus = 0 Then

            Dim DocType = IO.Path.GetExtension(SEDoc.Fullname)

            Try
                Dim Models As SolidEdgePart.Models = Nothing
                Dim PropertySets As SolidEdgeFramework.PropertySets = Nothing
                
                Select Case DocType
                    Case ".par"
                        Dim tmpSEDoc As SolidEdgePart.PartDocument
                        tmpSEDoc = CType(SEDoc, SolidEdgePart.PartDocument)
                        Models = tmpSEDoc.Models
                        PropertySets = CType(tmpSEDoc.Properties, SolidEdgeFramework.PropertySets)
                    Case ".psm"
                        Dim tmpSEDoc As SolidEdgePart.SheetMetalDocument
                        tmpSEDoc = CType(SEDoc, SolidEdgePart.SheetMetalDocument)
                        Models = tmpSEDoc.Models
                        PropertySets = CType(tmpSEDoc.Properties, SolidEdgeFramework.PropertySets)
                End Select
                
                If Models IsNot Nothing And PropertySets IsNot Nothing Then
                
                    Dim Filename As String = ""
                    Dim FoundPartCopy As Boolean = False
                
                    If Models.Count > 0 Then
                        Dim Model = Models.Item(1)
                        Dim PartCopies = Model.CopiedParts
                        If PartCopies.Count > 0 Then
                            Try
                                Dim PartCopy = PartCopies.Item(1)
                                Filename = PartCopy.FileName
                                FoundPartCopy = True
                            Catch ex As Exception
                                ExitStatus = 1
                                ErrorMessageList.Add("Unable to get part copy file name")
                            End Try
                        End If
                    End If
                
                    If FoundPartCopy Then
                        Dim PropertySet = PropertySets.Item("Custom")
                        Try
                            Dim Prop = PropertySet.Item("PartCopyFilename")
                            Prop.Value = Filename
                            PropertySet.Save()
                            SEApp.DoIdle()
                
                            If Not SEDoc.ReadOnly Then
                                SEDoc.Save()
                                SEApp.DoIdle()
                            End If
                        Catch ex As Exception
                            ExitStatus = 1
                            ErrorMessageList.Add("Unable to update property value")
                        End Try
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

    Public Shared Sub LoadLibrary(ParamArray libs As Object())
        For Each [lib] As String In libs
            'Console.WriteLine(String.Format("Loading library:  {0}", [lib]))
            Dim assm As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom([lib])
            'Console.WriteLine(assm.GetName().ToString())
        Next
    End Sub

    Private Shared Sub SaveErrorMessages(StartupPath As String, ErrorMessageList As List(Of String))
        Dim ErrorFilename As String
        ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)
        IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)
    End Sub

End Class
"@

Add-Type -TypeDefinition $Source -ReferencedAssemblies $DLLs -Language VisualBasic

[Snippet]::LoadLibrary($DLLs)

$ExitStatus = [Snippet]::RunSnippet($StartupPath)

Function ExitWithCode($exitcode) {
  $host.SetShouldExit($exitcode)
  Exit $exitcode
}

ExitWithCode($ExitStatus)
