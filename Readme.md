# Readme.md
Example programs for Solid Edge Housekeeper 'Run External Program' task.

The external program is a Console App that works on a single Solid Edge file.  Housekeeper serves up files one at a time for processing.  

The program needs to return an integer exit code, with 0 meaning success.  Anything else indicates an error.  You can optionally supply a text file, error_messages.txt -- in the same directory as the executable, to provide feedback to the user.  The format is given below.

Since a return value is needed, the declaration for Main() should be: Function Main() As Integer.  The usual declaration, Sub Main(), does not return a value.

Housekeeper launches the program as follows:

    Dim ExternalProgram As String = Configuration("TextBoxExternalProgramAssembly")
    Dim P As New Process
    Dim ExitCode As Integer

    P = Process.Start(ExternalProgram)
    P.WaitForExit()
    ExitCode = P.ExitCode

No arguments are passed to the program.  If you need to get a value from Housekeeper, such as a template file location, you can parse the defaults.txt file in Housekeeper's installation directory.  The file is updated just before processing is launched.  It should always reflect the current status of the form.

Housekeeper maintains a reference to the file being processed.  If that reference is broken, an exception will occur.  To avoid that, do not perform Close() or SaveAs() on the document.

No assumptions are made about what the external program does.  If you change a file and want to save it, that needs to be in the program.  If you open another file, you need to close it.


Repo https://github.rmcanany.com/SolidEdgeHousekeeperExternalProgramExamples


The format for error_messages.txt is:

error_number (space character) error message text

For example

1 Some error occurred
2 Some other error occurred

If the ExitCode is not 0, and the file is present, Solid Edge Housekeeper will include the error message in its log file, otherwise it will simply list the exit code.

