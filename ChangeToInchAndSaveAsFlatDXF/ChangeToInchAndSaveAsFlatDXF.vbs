' Robert McAnany 2022
' A more complete example of interacting with Solid Edge Housekeeper.
' See helpful links and other information at the bottom of this script.

Dim ErrorMessageArray(10)
ErrorMessageIdx = 0
ErrorCode = 0

ScriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
DefaultsFileName = ScriptDir & "\defaults.txt"
ErrorMessageFile = ScriptDir & "\error_messages.txt"



' #####################  Your code below  #####################



Dim SEDocFilename
Dim SEDocParentFolderName
Dim DXFFilename
Dim UnitsOfMeasure
Dim UnitOfMeasure
Dim Models

' The following statement lets you, not the VBScript interpreter, handle errors.
' Sometimes, like for syntax checking, you probably want the interpreter to do it.
' In that case, comment out the statement.
On Error Resume Next


' Connect to a running instance of Solid Edge.

Set SEApp = GetObject(, "SolidEdge.Application")
Set SEDocs = SEApp.Documents
Set SEDoc = SEDocs.Item(1)

' Handle errors as shown below.  Use ExitScript() when it doesn't make sense to continue.
' In this case, we are checking if an error occurred connecting to Solid Edge, so exiting 
' the script is the right thing to do.
' You don't have to use AddErrorMessage().  If you do, Housekeeper will report it in the log.
if Err Then
    Err.Clear
    AddErrorMessage("Solid Edge not running or no document open")
    ErrorCode = 1
    ExitScript()
End If


' Change the file's length units to inches.

Set UnitsOfMeasure = SEDoc.UnitsOfMeasure

For Each UnitOfMeasure In UnitsOfMeasure
    If UnitOfMeasure.Type = 1  Then  ' SolidEdgeConstants.UnitTypeConstants.igUnitDistance = 1
        ' UnitOfMeasure.Units = 17  ' SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthMillimeter = 17
        UnitOfMeasure.Units = 0  'SolidEdgeConstants.UnitOfMeasureLengthReadoutConstants.seLengthInch = 0
    End If
Next

if Err Then
    Err.Clear
    AddErrorMessage("Error reading or setting file units")
    ErrorCode = 1
    Set UnitsOfMeasure = Nothing
    Set UnitOfMeasure = Nothing
    ExitScript()
End If


' Create the output filename.

BuildDXFFilename()

'WScript.Echo DXFFilename

if Err Then
    Err.Clear
    AddErrorMessage("Error creating DXF filename '" & DXFFilename & "'")
    ErrorCode = 1
    ExitScript()
End If


' Save the flat pattern.

Set Models = SEDoc.Models
Call Models.SaveAsFlatDXFEx(DXFFilename, Nothing, Nothing, Nothing, True)
SEApp.DoIdle()

if Err Then
    Err.Clear
    AddErrorMessage("Error saving flat pattern.  Filename '" & DXFFilename & "'")
    ErrorCode = 1
    Set Models = Nothing
    ExitScript()
End If




Private Sub BuildDXFFilename()
    SEDocFilename = SEDoc.Name
    SEDocParentFolderName = CreateObject("Scripting.FileSystemObject").GetParentFolderName(SEDoc.FullName)

    Set objRegExp = New RegExp
    objRegExp.Pattern = "\.psm$" 

    DXFFilename = objRegExp.Replace(SEDocFilename, ".dxf")

    If LCase(GetConfigurationValue("CheckBoxSaveAsSheetmetalOutputDirectory")) = "true" Then
        'Save in original directory
        DXFFilename = SEDocParentFolderName & "\" & DXFFilename
    Else
        DXFFilename = GetConfigurationValue("TextBoxSaveAsSheetmetalOutputDirectory") & "\" & DXFFilename
    End If
End Sub




' #####################  Your code above  #####################

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

Private Function GetConfigurationValue(Key)
    ' You can find the settings on the Housekeeper dialog by manually inspecting defaults.txt
    ' They are stored as key-value pairs.
    ' The key is the text on the left of the '='.  The value is on the right.
    
    GetConfigurationValue = ""
    
    Dim Separator
    Dim SepPos
    Dim K
    Dim V
    Dim Line

    Separator = "="

    Set DefaultsFile = CreateObject("Scripting.FileSystemObject").OpenTextFile(DefaultsFileName, 1)

    Do While Not DefaultsFile.AtEndOfStream
         Line = DefaultsFile.ReadLine()
         SepPos = Instr(Line, Separator)
         K = Left(Line, SepPos - 1)
         V = Right(Line, Len(Line) - SepPos)
         If K = Key Then
             GetConfigurationValue = V
        End If
    Loop
    
    DefaultsFile.Close
    Set DefaultsFile = Nothing
    
End Function

' HELPFUL LINKS

' VBScript Language Reference
' https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/d1wf56tt(v=vs.84)

' Error Handling with the Err Object (VBScript)
' https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/sbf5ze0e(v=vs.84)

' FileSystemObject
' https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/filesystemobject-object

' Reading and writing a text file
' https://stackoverflow.com/questions/3117121/reading-and-writing-value-from-a-textfile-by-using-vbscript-code

' Output text in a Message Box.  Not a great idea for batch processing, but handy for debugging.
' WScript.Echo "Text to display to the user"

' Turn on error handling
' On Error Resume Next

' Turn off error handling
' On Error GoTo 0
