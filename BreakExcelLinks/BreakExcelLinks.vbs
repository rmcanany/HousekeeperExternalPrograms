' Author @Derek G
' https://community.sw.siemens.com/s/question/0D54O00007VUmLKSA1/can-you-do-a-macro-that-unlinks-variables-attached-to-an-excel-workbook

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

    Set variables = oDoc.variables
    For i = 1 To variables.count Step 1
        If InStr(variables(i).Formula, ".xlsx") Then
            variables(i).Formula = ""
        End If
    Next
    Set dimensions = variables.Query("*", 2, 1)
    For i = 1 To dimensions.count Step 1
        If InStr(dimensions(i).Formula, ".xlsx") Then
            dimensions(i).Formula = ""
        End If
    Next
     
if Err Then
    Err.Clear
    AddErrorMessage("Could not process variables")
    ErrorCode = 1
    ExitScript()
End If


    set oApp = Nothing
    set oDocs = Nothing
    set oDoc = Nothing
    Set variables = Nothing
    Set dimensions = Nothing

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
