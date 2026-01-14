Attribute VB_Name = "A_Working_CT2Xv1_4"
Sub Populate_Table_CT2Xv1_4()
    ' Original Written by A. Simon 5/10/2023
    ' Modified on 8/8/2024 for SharePoint compatibility
    ' Refactored by Michael Boyce with the help of ChatGPT 4, for readability, efficiency, and correct execution order

    Dim ws As Worksheet
    Dim NumYear As Integer, NameMonth As String, intMonth As Integer
    Dim ServerAddress As String, CompleteServerAddress As String
    Dim SaveFileName As String
    Dim wbCopy As Workbook, newWorkbook As Workbook
    Dim FileInFolder As String, fileToOpen As Workbook
    
    ' Ensure the user checks the necessary folder structure
    If Not ConfirmFolderStructure() Then Exit Sub

    ' Get Year and Month from User
    GetUserInput NumYear, NameMonth, intMonth

    ' Define SharePoint Server Path
    ServerAddress = Environ("OneDriveCommercial") & "\Monitoring Wells\Chloride monitoring\"
    CompleteServerAddress = ServerAddress & NumYear & "\" & NameMonth & "\"

    ' Ensure the directory exists
    If Dir(CompleteServerAddress, vbDirectory) = "" Then
        MsgBox "Error: The folder " & CompleteServerAddress & " does not exist.", vbCritical, "Missing Folder"
        Exit Sub
    End If

    ' Save workbook copy
    SaveFileName = CompleteServerAddress & "0" & intMonth & NumYear & ".xlsm"
    Set wbCopy = ThisWorkbook
    wbCopy.SaveCopyAs FileName:=SaveFileName
    Set newWorkbook = Workbooks.Open(SaveFileName)

    ' Process CSV files in the folder
    FileInFolder = Dir(CompleteServerAddress & "*.csv")

    Application.ScreenUpdating = False

    Do While Len(FileInFolder) > 0
        Set fileToOpen = Workbooks.Open(CompleteServerAddress & FileInFolder)
        Set ws = fileToOpen.Sheets(1)
        
        
        ' **Skip processing if sheet name contains "QC"**
        If InStr(1, ws.name, "QC", vbTextCompare) > 0 Then
            'Debug.Print "Skipping file: " & FileInFolder & " (QC detected)"
            fileToOpen.Close False
        Else

            ' Copy to main workbook & clean up unnecessary rows
            ws.Copy After:=newWorkbook.Sheets("Table")
            With ActiveSheet
                .rows("1:31").Delete ' Remove first 31 rows
                CleanColumns .Cells(1, 1).CurrentRegion ' Retain only required columns
                ArrangeColumns ' Move "Pressure (Ft H2O)" to Column A
    '            ReverseData .Cells(1, 1).CurrentRegion ' Reverse the data order
    '            FilterEvery10Feet .Cells(1, 1).CurrentRegion ' Select values every 10 ft
                ProcessData
                ReverseData
                ExtractMWNumberFromName (NameMonth)
                CopyMWDataToTable CompleteServerAddress, FileInFolder
                fileToOpen.Close False ' Close without saving
            End With
        End If
            
        
        FileInFolder = Dir ' Get next file
    Loop

    Application.ScreenUpdating = True

    ' Final Instructions
    'MsgBox "Please copy the conductivity data for each well to the Table. Paste special as values to retain format." & vbCrLf & _
    '       "Also, add the date of sampling and the depth of water in feet for each well.", vbInformation, "Final Steps"
End Sub
'-----------------------------------------------------------------------------------------------------------

Function ConfirmFolderStructure() As Boolean
    Dim message As String
    message = "For this Macro to work, ensure the following:" & vbCrLf & _
              "Main Folder: \Monitoring Wells\Chloride monitoring\" & vbCrLf & _
              "Data Folder: 4-DIGIT YEAR (e.g., 2017)" & vbCrLf & _
              "Subfolder: MONTH (3-letter abbr, e.g., Dec, Aug)" & vbCrLf & _
              "Raw Data File: MW# MONTH (e.g., MW1 Dec)" & vbCrLf & _
              "If you get an ERROR, check names and file locations."
    
    ConfirmFolderStructure = (MsgBox(message, vbOKCancel, "Saline Intrusion Data") = vbOK)
End Function
'-----------------------------------------------------------------------------------------------------------

'Sub GetUserInput(ByRef NumYear As Integer, ByRef NameMonth As String, ByRef intMonth As Integer)
'    NumYear = Val(InputBox("What year?", "File path...", Year(Date)))
'    NameMonth = LCase(InputBox("What month (3-letter abbr, e.g., Dec)?", "File path...", LCase(Format(Date, "mmm"))))
'    intMonth = Month(Date)
'End Sub


Sub GetUserInput(ByRef NumYear As Integer, ByRef NameMonth As String, ByRef intMonth As Integer)
    NumYear = Val(InputBox("What year?", "File path...", Year(Date)))
    NameMonth = LCase(InputBox("What month (3-letter abbr, e.g., Dec)?", "File path...", LCase(Format(Date, "mmm"))))
    'intMonth = Month(Date)
    intMonth = Month(DateValue("01-" & NameMonth & "-" & NumYear))
End Sub
'-----------------------------------------------------------------------------------------------------------


Sub CleanColumns(rng As Range)
    Dim col As Integer
    For col = rng.Columns.Count To 1 Step -1
        Select Case Trim(rng.Cells(1, col).Value)
            Case "Pressure (Ft H2O)", "Conductivity (µS/cm)"
                ' Keep these columns
            Case Else
                rng.Columns(col).Delete
        End Select
    Next col
End Sub
'-----------------------------------------------------------------------------------------------------------

Sub ArrangeColumns()
    Dim col As Integer
    For col = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Trim(ActiveSheet.Cells(1, col).Value) = "Pressure (Ft H2O)" Then
            ActiveSheet.Columns(col).Cut
            ActiveSheet.Columns(1).Insert Shift:=xlToRight
            Application.CutCopyMode = False ' Clear clipboard
            Exit For
        End If
    Next col
End Sub
'-----------------------------------------------------------------------------------------------------------


Sub ReverseData()
    Dim ws As Worksheet
    Dim lastRow As Long, firstRow As Long, i As Long
    Dim tempA As Variant, tempB As Variant
    
    Set ws = ActiveSheet
    firstRow = 2 ' Assuming row 1 contains headers
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    
    ' Swap values from top to bottom
    For i = 0 To (lastRow - firstRow) \ 2
        tempA = ws.Cells(firstRow + i, 1).Value
        ws.Cells(firstRow + i, 1).Value = ws.Cells(lastRow - i, 1).Value
        ws.Cells(lastRow - i, 1).Value = tempA
        
        tempB = ws.Cells(firstRow + i, 2).Value
        ws.Cells(firstRow + i, 2).Value = ws.Cells(lastRow - i, 2).Value
        ws.Cells(lastRow - i, 2).Value = tempB
    Next i
End Sub
'-----------------------------------------------------------------------------------------------------------

Sub FilterEvery10Feet()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim nextDepth As Double

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    nextDepth = 10 ' Start filtering at 10ft increments

    ' Process from top to bottom after reversal
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 1).Value < nextDepth Then
            ws.rows(i).Delete
        ElseIf ws.Cells(i, 1).Value >= nextDepth Then
            nextDepth = nextDepth + 10 ' Move to the next interval
        End If
    Next i
End Sub
'-----------------------------------------------------------------------------------------------------------

Sub ProcessData()
    ' Run the steps in correct order
    ReverseData
    FilterEvery10Feet
End Sub
'-----------------------------------------------------------------------------------------------------------

'Sub ExtractMWNumber(NameMonth As String)
'    Dim wsName As String
'    Dim mwNumber As String
'    Dim mwPos As Integer
'    Dim nmPos As Integer
'    Dim i As Integer
'
''    ' Set the NameMonth variable (Modify this as needed)
''    NameMonth = "Jan" ' Example: Set dynamically based on your needs
'
'    ' Get the active sheet name
'    wsName = ActiveSheet.Name
'
'    ' Find the position of "MW"
'    mwPos = InStr(1, wsName, "MW", vbTextCompare)
'
'    ' Find the position of NameMonth
'    nmPos = InStr(1, wsName, NameMonth, vbTextCompare)
'
'    ' Check if both MW and NameMonth exist in the sheet name
'    If mwPos > 0 And nmPos > mwPos Then
'        ' Start searching for digits after "MW"
'        mwPos = mwPos + 2 ' Move past "MW"
'        mwNumber = ""
'
'        ' Extract digits until NameMonth is reached
'        For i = mwPos To nmPos - 1
'            If Mid(wsName, i, 1) Like "[0-9]" Then
'                mwNumber = mwNumber & Mid(wsName, i, 1)
'            ElseIf mwNumber <> "" Then
'                ' Stop if we've already found digits and a non-digit appears
'                Exit For
'            End If
'        Next i
'    End If
'
'    ' Output result
'    If mwNumber = "" Then
'        MsgBox "MW number not found!", vbExclamation
'    Else
'        'MsgBox "MW Number: " & mwNumber, vbInformation
'    End If
'End Sub

Function ExtractMWNumberFromName(wsName As String) As String
    Dim mwNum As String
    Dim mwPos As Integer, i As Integer

    mwNum = ""
    mwPos = InStr(1, wsName, "MW", vbTextCompare)

    If mwPos > 0 Then
        mwPos = mwPos + 2 ' Move past "MW"
        For i = mwPos To Len(wsName)
            If Mid(wsName, i, 1) Like "[0-9]" Then
                mwNum = mwNum & Mid(wsName, i, 1)
            ElseIf mwNum <> "" Then
                Exit For
            End If
        Next i
    End If

    ExtractMWNumberFromName = mwNum
End Function
'-----------------------------------------------------------------------------------------------------------


'Sub CopyMWDataToTable()
'    Dim wbTarg As Workbook
'    Dim wsTable As Worksheet
'    Dim mwName As String, mwNum As String
'    Dim lastRow As Long, col As Range, targetCol As Integer
'    Dim cell As Range, found As Boolean
'
'    ' Set sheets
'    'MsgBox SaveFileName
'
'
'    Set wbTarg = ActiveWorkbook
'    Set ws = ActiveSheet ' Assumes MW sheet is active
'    Set wsTable = wbTarg.Sheets("Table") ' Change if needed
'
'    ' Extract MW number from the sheet name
'    mwName = ws.Name
'    mwNum = ""
'
'    ' Find "MW" in sheet name and extract first one or two digits
'    Dim mwPos As Integer, i As Integer
'    mwPos = InStr(1, mwName, "MW", vbTextCompare)
'
'    If mwPos > 0 Then
'        mwPos = mwPos + 2 ' Move past "MW"
'        For i = mwPos To Len(mwName)
'            If Mid(mwName, i, 1) Like "[0-9]" Then
'                mwNum = mwNum & Mid(mwName, i, 1)
'            ElseIf mwNum <> "" Then
'                Exit For
'            End If
'        Next i
'    End If
'
'    ' Ensure mwNum is valid
'    If mwNum = "" Then
'        MsgBox "MW number not found in sheet name!", vbExclamation
'        Exit Sub
'    End If
'
'    ' Find matching column in row 3 (B3:K3)
'    found = False
'    For Each cell In wsTable.Range("B3:K3")
'        If Left(cell.Value, Len(mwNum)) = mwNum Then
'            targetCol = cell.Column
'            found = True
'            Exit For
'        End If
'    Next cell
'
'    If Not found Then
'        MsgBox "No matching column found for MW" & mwNum, vbExclamation
'        Exit Sub
'    End If
'
'    ' Find last row in MW sheet column B
'    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
'
'    ' Copy data from MW sheet (B2:B lastRow) to Table sheet (row 32, matching column)
'    wsTable.Cells(6, targetCol).Resize(lastRow - 1, 1).Value = ws.Range("B2:B" & lastRow).Value
'
'    'MsgBox "Data copied to Table sheet!", vbInformation
'End Sub
'Sub CopyMWDataToTable(completeServerAddress As String, fileInFolder As String)
'    Dim wbTarg As Workbook
'    Dim ws As Worksheet, wsTable As Worksheet
'    Dim mwName As String, mwNum As String
'    Dim lastRow As Long, targetCol As Integer
'    Dim cell As Range, found As Boolean
'    Dim trailingNum As Double
'
'    Dim fixedFile As String
'
'    Dim nameCSVinStr As String
'    Dim dateOfCSV As Date
'
'    fixedFile = (completeServerAddress & fileInFolder)
'
'    ' Write date from CSV to row 4
'    dateOfCSV = FileDateTime(fixedFile)
'    nameCSVinStr = Format(dateOfCSV, "dd/mm/yyyy")
'    Debug.Print dateOfCSV
'
'    ' Set workbook and sheets
'    Set wbTarg = ActiveWorkbook
'    Set ws = ActiveSheet  ' Assumes MW sheet is active
'    Set wsTable = wbTarg.Sheets("Table")  ' Table sheet
'
'    ' Extract MW number from the sheet name
'    mwName = ws.Name
'    mwNum = ""
'
'    ' Find "MW" in sheet name and extract the following digits
'    Dim mwPos As Integer, i As Integer
'    mwPos = InStr(1, mwName, "MW", vbTextCompare)
'
'    If mwPos > 0 Then
'        mwPos = mwPos + 2 ' Move past "MW"
'        For i = mwPos To Len(mwName)
'            If Mid(mwName, i, 1) Like "[0-9]" Then
'                mwNum = mwNum & Mid(mwName, i, 1)
'            ElseIf mwNum <> "" Then
'                Exit For
'            End If
'        Next i
'    End If
'
'    ' Ensure mwNum is valid
'    If mwNum = "" Then
'        MsgBox "MW number not found in sheet name!", vbExclamation
'        Exit Sub
'    End If
'
'    ' Find matching column in row 3 (B3:K3)
'    found = False
'    For Each cell In wsTable.Range("B3:K3")
'        If Left(cell.Value, Len(mwNum)) = mwNum Then
'            targetCol = cell.Column
'            found = True
'            Exit For
'        End If
'    Next cell
'
'    If Not found Then
'        MsgBox "No matching column found for MW" & mwNum, vbExclamation
'        Exit Sub
'    End If
'
'    ' Extract and place the trailing number in row 5
'    trailingNum = GetTrailingNumber(ws)
'    wsTable.Cells(5, targetCol).Value = Format(trailingNum, "0.00")
'
'    ' Find last row in MW sheet column B
'    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
'
'    ' Copy data from MW sheet (B2:B lastRow) to Table sheet (row 6, matching column)
'    wsTable.Cells(6, targetCol).Resize(lastRow - 1, 1).Value = ws.Range("B2:B" & lastRow).Value
'
'
'
'    ' MsgBox "Data copied to Table sheet!", vbInformation
'End Sub
Sub CopyMWDataToTable(CompleteServerAddress As String, FileInFolder As String)
    Dim wbTarg As Workbook
    Dim ws As Worksheet, wsTable As Worksheet
    Dim mwName As String, mwNum As String
    Dim lastRow As Long, targetCol As Integer
    Dim cell As Range, found As Boolean
    Dim trailingNum As Double
    Dim dateOfCSV As Date
    Dim fixedFile As String

    ' Check for QC or CCV in filename
    If InStr(1, FileInFolder, "QC", vbTextCompare) > 0 Or InStr(1, FileInFolder, "CCV", vbTextCompare) > 0 Then
        Debug.Print "Skipping file: " & FileInFolder & " (QC/CCV detected)"
        Exit Sub
    End If

    ' Get file creation date
    fixedFile = CompleteServerAddress & FileInFolder
    dateOfCSV = FileDateTime(fixedFile)

    ' Set workbook and sheets
    Set wbTarg = ActiveWorkbook
    Set ws = ActiveSheet  ' Assumes MW sheet is active
    Set wsTable = wbTarg.Sheets("Table")  ' Table sheet

    ' Extract MW number from the sheet name
    mwName = ws.name
    mwNum = ExtractMWNumberFromName(mwName)

    ' Ensure MW number is valid
    If mwNum = "" Then
        MsgBox "MW number not found in sheet name!", vbExclamation
        Exit Sub
    End If

    ' Find matching column in row 3 (B3:K3)
    found = False
    For Each cell In wsTable.Range("B3:K3")
        If Left(cell.Value, Len(mwNum)) = mwNum Then
            targetCol = cell.Column
            found = True
            Exit For
        End If
    Next cell

    If Not found Then
        MsgBox "No matching column found for MW" & mwNum, vbExclamation
        Exit Sub
    End If

    ' Place the CSV date into row 4 of the corresponding column
    wsTable.Cells(4, targetCol).Value = Format(dateOfCSV, "mm/dd/yyyy")

    ' Extract and place the trailing number in row 5
    trailingNum = GetTrailingNumber(ws)
    wsTable.Cells(5, targetCol).Value = Format(trailingNum / 100, "0.00")

    ' Find last row in MW sheet column B
    lastRow = ws.Cells(ws.rows.Count, 2).End(xlUp).Row

    ' Copy data from MW sheet (B2:B lastRow) to Table sheet (row 6, matching column)
    wsTable.Cells(6, targetCol).Resize(lastRow - 1, 1).Value = ws.Range("B2:B" & lastRow).Value

    ' Debug output
    Debug.Print "Data copied to Table sheet for MW" & mwNum
End Sub
'-----------------------------------------------------------------------------------------------------------


'Function GetTrailingNumber(ws As Worksheet) As Double
'    Dim wsName As String
'    Dim matches As Object
'    Dim regex As Object
'
'    wsName = ws.Name
'
'    ' Create regex object
'    Set regex = CreateObject("VBScript.RegExp")
'    regex.Pattern = "(\d{3,4})$"   ' Match trailing 3 or 4 digits
'    regex.Global = False
'
'    If regex.Test(wsName) Then
'        Set matches = regex.Execute(wsName)
'        Dim trailingNum As String
'        trailingNum = matches(0).Value
'
'        ' Format as decimal with two places
'        If Len(trailingNum) = 3 Then
'            GetTrailingNumber = CDbl(Left(trailingNum, 1) & "." & Mid(trailingNum, 2))
'        ElseIf Len(trailingNum) = 4 Then
'            GetTrailingNumber = CDbl(Left(trailingNum, 2) & "." & Mid(trailingNum, 3))
'        End If
'    Else
'        ' Return 0 if no match is found
'        GetTrailingNumber = 0
'    End If
'End Function

Function GetTrailingNumber(ws As Worksheet) As Double
    Dim wsName As String
    Dim matches As Object
    Dim regex As Object

    wsName = ws.name

    ' Create regex object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "(\d{3,4})$"   ' Match trailing 3 or 4 digits
    regex.Global = False

    If regex.Test(wsName) Then
        Set matches = regex.Execute(wsName)
        If matches.Count > 0 Then
            GetTrailingNumber = CDbl(matches(0).Value)
        Else
            GetTrailingNumber = 0
        End If
    Else
        GetTrailingNumber = 0
    End If
End Function
'-----------------------------------------------------------------------------------------------------------









