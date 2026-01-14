Attribute VB_Name = "A_saltReport_Final"
Sub PrepareSaltReport()

    Dim saltbook As Workbook
    Dim Saltfile As String
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim sourceCol As Long
    Dim targetCol As Long
    Dim x As Long

    ' Confirm
    If MsgBox("Continue to copy data into Salt Data Spreadsheet?", _
              vbYesNo + vbQuestion, "Confirm") <> vbYes Then Exit Sub

    ' Source worksheet MUST be "Table" in active workbook
    Set wsSource = ThisWorkbook.Worksheets("Table")

    ' SALTDATA path
    Saltfile = Environ("OneDriveCommercial") & "\Monitoring Wells\SALTDATA-15.xlsx"

    ' Open SALTDATA workbook
    On Error Resume Next
    Set saltbook = Workbooks.Open(Saltfile)
    On Error GoTo 0

    If saltbook Is Nothing Then
        MsgBox "Salt Data file not found!", vbCritical
        Exit Sub
    End If

    ' Loop through SWM columns (B–K)
    For x = 1 To 10

        ' Source column (B = 2)
        sourceCol = x + 1

        ' Target worksheet determined by SWM (already working)
        Set wsTarget = saltbook.Worksheets(x)

        ' Determine target column ONCE (row 9 defines append point)
        targetCol = wsTarget.Cells(9, wsTarget.Columns.Count).End(xlToLeft).Column + 1

        ' Copy ONLY rows 4–32 from Table worksheet
        wsTarget.Cells(9, targetCol).Resize(29, 1).Value = _
            wsSource.Cells(4, sourceCol).Resize(29, 1).Value

        ' Formatting
        wsTarget.Cells(9, targetCol).NumberFormat = "mm/dd/yyyy"
        wsTarget.Range(wsTarget.Cells(10, targetCol), _
                       wsTarget.Cells(37, targetCol)).NumberFormat = "#,##0"

        With wsTarget
            .Cells(40, targetCol).Value = .Cells(10, targetCol).Value
            .Cells(10, targetCol).ClearContents
    
        End With

    
    Next x

    MsgBox "Salt data successfully copied.", vbInformation

End Sub




