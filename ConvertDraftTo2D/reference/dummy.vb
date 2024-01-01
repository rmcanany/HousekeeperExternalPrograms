Public Class dummy

    Public Shared Function ConvertTo2D(StartupPath As String) As Integer
        Dim ExitStatus As Integer = 0
        Dim ErrorMessageList As New List(Of String)

        Dim SEApp As Object = Nothing
        Dim SEDoc As Object = Nothing
        Dim Sheets As Object

        Dim Configuration As New Dictionary(Of String, String)

        Configuration = GetConfiguration(StartupPath)

        Try
            SEApp = Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application")
            SEDoc = SEApp.ActiveDocument
            Console.WriteLine(String.Format("Processing {0}", SEDoc.Name))
        Catch ex As Exception
            ExitStatus = 1
            ErrorMessageList.Add("Unable to connect to Solid Edge, or a Draft file is not open")
        End Try

        If ExitStatus = 0 Then

            Try
                Sheets = SEDoc.Sections.BackgroundSection.Sheets
                ProcessCallouts(Sheets)
                ProcessDrawingViews(Sheets)

                Sheets = SEDoc.Sections.WorkingSection.Sheets
                ProcessCallouts(Sheets)
                ProcessDrawingViews(Sheets)

            Catch ex As Exception
                ExitStatus = 1
                ErrorMessageList.Add("Unable to process all sheets.  No changes made.")
            End Try
        End If

        If (ExitStatus = 0) Or (ExitStatus = 2) Then
            SEDoc.Save()
            SEApp.DoIdle()
        End If

        If Not ExitStatus = 0 Then
            SaveErrorMessages(StartupPath, ErrorMessageList)
        End If

        Return ExitStatus
    End Function

    Private Shared Sub ProcessCallouts(Sheets As Object)
        Dim Sheet As Object
        Dim Balloons As Object
        Dim Balloon As Object

        For Each Sheet In Sheets
            Console.WriteLine(String.Format("{0}Callouts {1}", "    ", Sheet.Name))
            Balloons = Sheet.Balloons
            For Each Balloon In Balloons
                Try
                    Balloon.BalloonText = Balloon.BalloonDisplayedText
                Catch ex2 As Exception
                End Try
            Next
        Next

    End Sub

    Private Shared Sub ProcessDrawingViews(Sheets As Object)
        Dim Sheet As Object
        Dim DrawingViews As Object
        Dim DrawingView As Object

        For Each Sheet In Sheets
            Console.WriteLine(String.Format("{0}Drawing views {1}", "    ", Sheet.Name))
            DrawingViews = Sheet.DrawingViews
            For Each DrawingView In DrawingViews
                ' Some drawing views are already 2D
                Try
                    DrawingView.Drop()
                Catch ex2 As Exception
                End Try
            Next
        Next

    End Sub

    Private Shared Sub SaveErrorMessages(StartupPath As String, ErrorMessageList As List(Of String))
        Dim ErrorFilename As String

        ErrorFilename = String.Format("{0}\error_messages.txt", StartupPath)

        IO.File.WriteAllLines(ErrorFilename, ErrorMessageList)

    End Sub


    Private Shared Function GetConfiguration(StartupPath As String) As Dictionary(Of String, String)
        Dim Configuration As New Dictionary(Of String, String)
        Dim Defaults As String() = Nothing
        Dim Key As String
        Dim Value As String = ""
        Dim DefaultsFilename As String
        Dim KVPairArray As Array
        Dim i As Integer

        DefaultsFilename = String.Format("{0}\defaults.txt", StartupPath)

        Try
            Defaults = IO.File.ReadAllLines(DefaultsFilename)

            For Each KVPair As String In Defaults
                If Not KVPair.Contains("=") Then
                    Continue For
                End If

                KVPairArray = KVPair.Split("=")
                Key = KVPairArray(0)

                For i = 1 To KVPairArray.Length - 1
                    If i = 1 Then
                        Value = KVPairArray(i)
                    Else
                        Value = String.Format("{}={}", Value, KVPairArray(i))
                    End If
                Next

                Configuration(Key) = Value
            Next

        Catch ex As Exception
        End Try


        Return Configuration
    End Function

End Class
