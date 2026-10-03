' Command Line Arguments for Report.exe
' https://docs.plm.automation.siemens.com/docs/se/2020/api/webframe.html

' VB Script String Functions
' https://www.w3schools.com/asp/asp_ref_vbscript_functions.asp

Dim ErrorMessageArray(10)
ErrorMessageIdx = 0
ErrorCode = 0

ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
ErrorMessageFile = ScriptDir & "\error_messages.txt"

On Error Resume Next

Set oApp = GetObject(, "SolidEdge.Application")
Set oDocs = oApp.Documents
Set oDoc = oDocs.Item(1)

if Err Then
    Err.Clear
    AddErrorMessage("Solid Edge not running or no document open")
    ErrorCode = 1
    ExitScript()
End If

Dim ProgramName
Dim AssemblyFile
Dim ReportType
Dim ReportFile
Dim Cmd

ProgramName = "C:\Program Files\Siemens\Solid Edge 2022\Program\report.exe"
AssemblyFile = oDoc.FullName
ReportType = "ASM_ATOMIC_PARTS"  ' See options in the Command Line Arguments link.
ReportFile = Replace(AssemblyFile, ".asm", ".txt")

Cmd = Chr(34) & ProgramName & Chr(34) & _
      " " & AssemblyFile & _
      " /t=" & ReportType & _
      " /o=" & ReportFile & _
      " /w=FALSE"

Set WshShell = WScript.CreateObject("WScript.Shell")

Dim StatusCode

StatusCode = WshShell.Run(Cmd, 1, true)

if Err Then
    Err.Clear
    AddErrorMessage("Could not run " & ProgramName)
    ErrorCode = 1
    ExitScript()
End If

set oApp = Nothing
set oDocs = Nothing
set oDoc = Nothing


ExitScript()  ' This saves any error messages and returns the ErrorCode



Private Sub ExitScript()
    SaveErrorMessages() 
    Set SEApp = Nothing
    Set SEDocs = Nothing
    Set SEDoc = Nothing
    WScript.Quit(ErrorCode)
End Sub

Private Sub AddErrorMessage(ErrorMessage)
    ErrorMessageArray(ErrorMessageIdx) = ErrorMessage
    ErrorMessageIdx = ErrorMessageIdx + 1
End Sub

Private Sub SaveErrorMessages()
    Dim i
    Set objFileToWrite = CreateObject("Scripting.FileSystemObject").CreateTextFile(ErrorMessageFile, True, True)
    For i = 0 to ErrorMessageIdx - 1
        objFileToWrite.WriteLine(ErrorMessageArray(i))
    Next
    objFileToWrite.Close
    Set objFileToWrite = Nothing
End Sub
