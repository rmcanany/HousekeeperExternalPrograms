' Command Line Arguments for Report.exe
' https://docs.plm.automation.siemens.com/docs/se/2020/api/webframe.html

' VB Script String Functions
' https://www.w3schools.com/asp/asp_ref_vbscript_functions.asp

Set oApp = GetObject(, "SolidEdge.Application")
Set oDocs = oApp.Documents
Set oDoc = oDocs.Item(1)

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

set oApp = Nothing
set oDocs = Nothing
set oDoc = Nothing
