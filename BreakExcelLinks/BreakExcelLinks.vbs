' Author @Derek G
' https://community.sw.siemens.com/s/question/0D54O00007VUmLKSA1/can-you-do-a-macro-that-unlinks-variables-attached-to-an-excel-workbook


Set oApp = GetObject(, "SolidEdge.Application")
    Set oDocs = oApp.Documents
    Set oDoc = oDocs.Item(1)
     
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
     
    set oApp = Nothing
    set oDocs = Nothing
    set oDoc = Nothing
    Set variables = Nothing
    Set dimensions = Nothing