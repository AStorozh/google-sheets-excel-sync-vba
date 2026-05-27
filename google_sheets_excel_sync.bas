Option Explicit

' ============================================================
' Google Sheets to Excel Sync VBA
' ------------------------------------------------------------
' This macro downloads a Google Sheets document as XLSX
' and copies selected sheets into target sheets in Excel.
'
' Features:
' - Cross-platform support: Windows and macOS
' - Copies values, formats, column widths and row heights
' - Enables filters on imported data
' - Fixes barcode columns displayed in scientific notation
'
' IMPORTANT:
' Replace GOOGLE_SHEETS_XLSX_URL and sheet mappings below
' with your own test or production values.
' Do not publish private company links or real business data.
' ============================================================


#If Mac Then
#Else
    #If VBA7 Then
        Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
            ByVal pCaller As LongPtr, _
            ByVal szURL As String, _
            ByVal szFileName As String, _
            ByVal dwReserved As Long, _
            ByVal lpfnCB As LongPtr) As Long
    #Else
        Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" ( _
            ByVal pCaller As Long, _
            ByVal szURL As String, _
            ByVal szFileName As String, _
            ByVal dwReserved As Long, _
            ByVal lpfnCB As Long) As Long
    #End If
#End If


Sub Refresh_Data_From_Google_Sheets()

    Dim googleXlsxUrl As String
    Dim tempPath As String
    Dim sourceWb As Workbook
    Dim oldSheet As Worksheet

    Set oldSheet = ActiveSheet

    Application.ScreenUpdating = False

    ' Replace this URL with your own published/export Google Sheets XLSX link.
    googleXlsxUrl = "https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/export?format=xlsx"

    tempPath = GetTempPath() & "google_sheets_source.xlsx"

    DownloadFile googleXlsxUrl, tempPath

    Set sourceWb = Workbooks.Open(tempPath)

    ' Sheet mapping:
    ' CopySheetData sourceWorkbook, "Google sheet name", targetWorkbook, "Excel sheet name"
    CopySheetData sourceWb, "Source Sheet 1", ThisWorkbook, "Target Sheet 1"
    CopySheetData sourceWb, "Source Sheet 2", ThisWorkbook, "Target Sheet 2"
    CopySheetData sourceWb, "Source Sheet 3", ThisWorkbook, "Target Sheet 3"
    CopySheetData sourceWb, "Source Sheet 4", ThisWorkbook, "Target Sheet 4"
    CopySheetData sourceWb, "Source Sheet 5", ThisWorkbook, "Target Sheet 5"

    sourceWb.Close SaveChanges:=False

    oldSheet.Activate

    Application.ScreenUpdating = True

    MsgBox "Data has been updated successfully.", vbInformation

End Sub


Function GetTempPath() As String

    #If Mac Then
        GetTempPath = Environ("TMPDIR")
    #Else
        GetTempPath = Environ("TEMP") & "\"
    #End If

End Function


Sub DownloadFile(ByVal fileUrl As String, ByVal savePath As String)

    #If Mac Then

        Dim shellCommand As String

        shellCommand = "curl -L -o " & Chr(34) & savePath & Chr(34) & " " & Chr(34) & fileUrl & Chr(34)

        MacScript "do shell script " & Chr(34) & Replace(shellCommand, Chr(34), "\" & Chr(34)) & Chr(34)

    #Else

        Dim result As Long

        result = URLDownloadToFile(0, fileUrl, savePath, 0, 0)

        If result <> 0 Then
            MsgBox "Failed to download the Google Sheets file. Check internet connection and file access permissions.", vbCritical
            End
        End If

    #End If

End Sub


Sub CopySheetData(ByVal sourceWb As Workbook, ByVal sourceSheetName As String, ByVal targetWb As Workbook, ByVal targetSheetName As String)

    Dim sourceWs As Worksheet
    Dim targetWs As Worksheet
    Dim sourceRange As Range

    Set sourceWs = sourceWb.Worksheets(sourceSheetName)
    Set targetWs = targetWb.Worksheets(targetSheetName)
    Set sourceRange = sourceWs.UsedRange

    targetWs.Cells.Clear

    sourceRange.Copy

    With targetWs.Range("A1")
        .PasteSpecial Paste:=xlPasteValues
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteColumnWidths
    End With

    Application.CutCopyMode = False

    CopyRowHeights sourceWs, targetWs, sourceRange
    FormatBarcodeColumns targetWs
    EnableAutoFilter targetWb, targetWs

End Sub


Sub CopyRowHeights(ByVal sourceWs As Worksheet, ByVal targetWs As Worksheet, ByVal sourceRange As Range)

    Dim r As Long
    Dim firstRow As Long
    Dim lastRow As Long

    firstRow = sourceRange.Row
    lastRow = sourceRange.Row + sourceRange.Rows.Count - 1

    For r = firstRow To lastRow
        targetWs.Rows(r).RowHeight = sourceWs.Rows(r).RowHeight
    Next r

End Sub


Sub FormatBarcodeColumns(ByVal ws As Worksheet)

    Dim lastCol As Long
    Dim col As Long
    Dim headerText As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To lastCol

        headerText = LCase(Trim(CStr(ws.Cells(1, col).Value)))

        If headerText = "barcode" Or headerText = "bar code" Or headerText = "баркод" Or InStr(headerText, "barcode") > 0 Or InStr(headerText, "баркод") > 0 Then
            ws.Columns(col).NumberFormat = "0"
            ws.Columns(col).AutoFit
        End If

    Next col

End Sub


Sub EnableAutoFilter(ByVal targetWb As Workbook, ByVal ws As Worksheet)

    Dim lastCell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    Dim filterRange As Range

    Set lastCell = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        Exit Sub
    End If

    lastRow = ws.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    lastCol = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    targetWb.Activate
    ws.Activate

    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If

    Set filterRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    filterRange.Select
    Selection.AutoFilter

End Sub