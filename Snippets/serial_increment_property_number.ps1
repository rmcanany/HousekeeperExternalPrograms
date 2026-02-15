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
                If DocType = ".par" Then
                    Dim tmpSEDoc As SolidEdgePart.PartDocument = CType(SEDoc, SolidEdgePart.PartDocument)
                    Dim DirName As String = IO.Path.GetDirectoryName(tmpSEDoc.FullName)
                    
                    Dim PropertySets As SolidEdgeFramework.PropertySets = CType(tmpSEDoc.Properties, SolidEdgeFramework.PropertySets)
                    Dim PropertySet As SolidEdgeFramework.Properties = PropertySets.Item("Custom")
                    Dim Prop As SolidEdgeFramework.Property = PropertySet.Item("index_number")
                
                    Dim start_index As Integer = CInt(Prop.Value)
                
                    For idx As Integer = start_index + 1 To 10
                        Prop.Value = CStr(idx)
                        PropertySet.Save()
                        SEApp.DoIdle()
                
                        SEApp.StartCommand(CType(11292, SolidEdgeFramework.SolidEdgeCommandConstants)) ' Update active level
                        SEApp.DoIdle()
                
                        tmpSEDoc.SaveCopyAs(String.Format("{0}\Part_{1}.par", DirName, CStr(idx)))
                        SEApp.DoIdle()
                    Next
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
