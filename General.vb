Sub RefreshAndExpandRange()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim copyRange As Range
    
    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Source of new data
    Set wsTarget = ThisWorkbook.Sheets("Sheet2") ' Destination to refresh
    
    ' Find last row with data in source sheet (assuming data starts in row 1 and in column A)
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Define the source range to copy (you can adjust the columns as needed)
    Set copyRange = wsSource.Range("A1:E" & lastRow) ' Adjust "E" to your last column
    
    ' Clear previous data in target sheet
    wsTarget.Cells.ClearContents
    
    ' Copy to target sheet
    copyRange.Copy Destination:=wsTarget.Range("A1")
    
    MsgBox "Data refreshed and range expanded successfully!", vbInformation
End Sub



Sub ExpandFormulasToMatchData()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
    Dim lastCol As Long
    Dim i As Long

    ' Set worksheets
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    Set wsTarget = ThisWorkbook.Sheets("Sheet2")

    ' Find last row in source (data rows)
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Find last row with formulas in target (assume formulas start at row 2)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row

    ' Find last column with formulas (based on header row 1)
    lastCol = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column

    ' Expand formulas if source has more rows
    If lastRowSource > lastRowTarget Then
        wsTarget.Range(wsTarget.Cells(lastRowTarget, 1), wsTarget.Cells(lastRowTarget, lastCol)).Copy
        wsTarget.Range(wsTarget.Cells(lastRowTarget + 1, 1), wsTarget.Cells(lastRowSource, lastCol)).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False
        MsgBox "Formulas extended from row " & lastRowTarget + 1 & " to " & lastRowSource, vbInformation
    Else
        MsgBox "No new rows to update.", vbInformation
    End If
End Sub



Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    Application.EnableEvents = False

    ' Call your formula expansion macro
    Call ExpandFormulasToMatchData

SafeExit:
    Application.EnableEvents = True
End Sub



Sub CopyRangeAndChartFromAnotherWorkbook()
    Dim sourceWB As Workbook
    Dim targetWB As Workbook
    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet
    Dim chartObj As ChartObject
    Dim filePath As String
    Dim copiedChart As ChartObject

    ' Update this path to the source workbook location
    filePath = "C:\Path\To\SourceWorkbook.xlsx"

    ' Set reference to this workbook
    Set targetWB = ThisWorkbook
    Set targetWS = targetWB.Sheets("Sheet1") ' Change as needed

    ' Open source workbook (hidden)
    Set sourceWB = Workbooks.Open(filePath, ReadOnly:=True)
    Set sourceWS = sourceWB.Sheets("DataSheet") ' Change to actual sheet name

    ' === Copy Data Range ===
    sourceWS.Range("A1:E20").Copy Destination:=targetWS.Range("A1") ' Adjust range and location

    ' === Copy Chart ===
    ' Assuming an embedded chart is on the source worksheet
    If sourceWS.ChartObjects.Count > 0 Then
        sourceWS.ChartObjects(1).Copy
        targetWS.Paste
        Set copiedChart = targetWS.ChartObjects(targetWS.ChartObjects.Count)
        copiedChart.Top = targetWS.Range("G1").Top ' Position chart
        copiedChart.Left = targetWS.Range("G1").Left
    Else
        MsgBox "No chart found in the source worksheet.", vbExclamation
    End If

    ' Close source workbook without saving
    sourceWB.Close False

    MsgBox "Range and chart copied successfully!", vbInformation
End Sub



Sub CopyFilesAndSendEmail()
    Dim fso As Object
    Dim sourceFolderPath As String
    Dim targetFolderPath As String
    Dim sourceFolder As Object
    Dim file As Object
    Dim copiedCount As Integer

    ' === SETUP DYNAMIC SOURCE FOLDER ===
    Dim todayFolder As String
    todayFolder = Format(Date, "yyyy-mm-dd") ' E.g., "2025-05-02"
    
    sourceFolderPath = "C:\Your\Source\Root\" & todayFolder & "\" ' Modify path
    targetFolderPath = "C:\Your\Target\Folder\"                   ' Modify path

    ' === CREATE FILESYSTEMOBJECT ===
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Check if source exists
    If Not fso.FolderExists(sourceFolderPath) Then
        MsgBox "Source folder does not exist: " & sourceFolderPath, vbExclamation
        Exit Sub
    End If

    ' Create target folder if it doesn't exist
    If Not fso.FolderExists(targetFolderPath) Then
        fso.CreateFolder targetFolderPath
    End If

    ' Get folder and copy files
    Set sourceFolder = fso.GetFolder(sourceFolderPath)
    copiedCount = 0

    For Each file In sourceFolder.Files
        fso.CopyFile file.Path, targetFolderPath & fso.GetFileName(file.Path), True
        copiedCount = copiedCount + 1
    Next file

    ' === SEND EMAIL USING OUTLOOK ===
    Dim OutlookApp As Object
    Dim OutlookMail As Object

    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0

    If Not OutlookApp Is Nothing Then
        Set OutlookMail = OutlookApp.CreateItem(0)
        With OutlookMail
            .To = "recipient@example.com"
            .Subject = "Files Copied Notification"
            .Body = copiedCount & " file(s) were copied from '" & sourceFolderPath & "' to '" & targetFolderPath & "'."
            .Send ' or use .Display to review before sending
        End With
    Else
        MsgBox "Outlook is not available.", vbCritical
    End If

    MsgBox copiedCount & " file(s) copied and email sent.", vbInformation
End Sub


