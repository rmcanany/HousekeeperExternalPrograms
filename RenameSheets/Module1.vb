Option Strict On

Imports SolidEdgeFramework, SolidEdgeFrameworkSupport, SolidEdgeDraft, System

Module Module1


    Function Main() As Integer

        ' Original code by Tushar Suradkar
        ' https://community.sw.siemens.com/s/question/0D54O00006BtAnKSAV/code-rename-sheets-to-the-referenced-model-name

        Console.WriteLine("RenameSheets starting...")

        Dim ExitCode As Integer = 0
        Dim seApp As SolidEdgeFramework.Application = Nothing
        Dim seDoc As SolidEdgeDraft.DraftDocument = Nothing

        Dim Suffix As String

        Try
            seApp = CType(Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application"), Application)
            seDoc = CType(seApp.ActiveDocument, DraftDocument)
        Catch ex As Exception
            ExitCode = 1
        End Try


        If ExitCode = 0 Then
            Dim Sheetnames As New List(Of String)
            Dim Sheetname As String

            Dim seSections As Sections = seDoc.Sections
            Dim seWorkingSection As Section = seSections.WorkingSection
            Dim seSheets As SectionSheets = seWorkingSection.Sheets

            Dim seViews As DrawingViews = Nothing

            Suffix = GetSuffix()

            Try
                ' In case this program has already run this file, first rename sheets to random values
                For Each seSheet As Sheet In seSheets
                    seViews = seSheet.DrawingViews
                    If seViews.Count > 0 Then
                        seSheet.Name = String.Format("RenameSheets-{0}", CInt(Int((1000000 * Rnd()) + 1)))
                    End If
                Next

                ' Rename to modellink name
                For Each seSheet As Sheet In seSheets
                    seViews = seSheet.DrawingViews
                    If seViews.Count > 0 Then
                        Sheetname = Rename(seViews, Sheetnames, Suffix)
                        seSheet.Name = Sheetname
                        Sheetnames.Add(Sheetname)
                        Console.WriteLine(Sheetname)
                    End If
                Next
            Catch ex As Exception
                ExitCode = 2
            End Try

        End If

        If ExitCode = 0 Then
            seDoc.Save()
            seApp.DoIdle()
        End If

        Console.WriteLine("RenameSheets complete")

        Return ExitCode

    End Function


    Private Function Rename(seViews As DrawingViews, Sheetnames As List(Of String), Suffix As String) As String
        ' Sheet names need to be unique.  This function handles the case where two sheets have the same first ModelLink
        Dim BaseName As String
        Dim Name As String
        Dim View As DrawingView = seViews.Item(1)
        Dim count As Integer = 1

        BaseName = IO.Path.GetFileNameWithoutExtension(CType(View.ModelLink, ModelLink).FileName)

        If View.Configuration <> "" Then
            BaseName = String.Format("{0}-{1}", BaseName, View.Configuration)
        End If

        Name = BaseName

        While Sheetnames.Contains(Name)
            Name = String.Format("{0}-{1}({2})", BaseName, Suffix, count)
            count += 1
        End While

        Return Name
    End Function

    Function GetSuffix() As String
        Dim Suffix As String = "Copy"
        Dim StartupPath As String = AppDomain.CurrentDomain.BaseDirectory

        Dim Filename As String = String.Format("{0}suffix.txt", StartupPath)

        Try
            Suffix = IO.File.ReadAllLines(Filename)(0)
        Catch ex As Exception
            Console.WriteLine(String.Format("Problem processing {0}", Filename))
        End Try

        Return Suffix
    End Function

End Module
