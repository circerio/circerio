Attribute VB_Name = "modUtils"
Option Explicit

Public Function PickExcelFile() As String
    Dim dlg As Object

    Set dlg = Application.FileDialog(3) ' msoFileDialogFilePicker
    With dlg
        .Title = "Select forecast workbook"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xls;*.csv"
        If .Show = -1 Then PickExcelFile = .SelectedItems(1)
    End With
End Function

Public Function FirstExistingHeader(ByVal ws As Worksheet, ByVal candidates As Variant) As Long
    Dim i As Long
    Dim col As Long

    For i = LBound(candidates) To UBound(candidates)
        col = FindHeaderColumn(ws, CStr(candidates(i)))
        If col > 0 Then
            FirstExistingHeader = col
            Exit Function
        End If
    Next i
End Function

Public Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim currentHeader As String

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        currentHeader = Trim$(CStr(ws.Cells(1, c).Value))
        If StrComp(currentHeader, headerText, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
End Function

Public Function TryNormalizeMonth(ByVal rawValue As Variant, ByRef resultMonth As Date) As Boolean
    On Error GoTo Fail

    If IsDate(rawValue) Then
        resultMonth = DateSerial(Year(CDate(rawValue)), Month(CDate(rawValue)), 1)
        TryNormalizeMonth = True
        Exit Function
    End If

    If Len(Trim$(CStr(rawValue))) = 7 Then
        resultMonth = DateSerial(CInt(Left$(CStr(rawValue), 4)), CInt(Right$(CStr(rawValue), 2)), 1)
        TryNormalizeMonth = True
        Exit Function
    End If

Fail:
End Function

Public Function TryNormalizeMonthHeader(ByVal rawHeader As Variant, ByRef resultMonth As Date) As Boolean
    Dim headerText As String
    Dim monthNo As Long

    headerText = LCase$(Trim$(CStr(rawHeader)))

    Select Case Left$(headerText, 3)
        Case "jan": monthNo = 1
        Case "feb": monthNo = 2
        Case "mar": monthNo = 3
        Case "apr": monthNo = 4
        Case "may": monthNo = 5
        Case "jun": monthNo = 6
        Case "jul": monthNo = 7
        Case "aug": monthNo = 8
        Case "sep": monthNo = 9
        Case "oct": monthNo = 10
        Case "nov": monthNo = 11
        Case "dec": monthNo = 12
    End Select

    If monthNo > 0 Then
        resultMonth = DateSerial(Year(Date), monthNo, 1)
        TryNormalizeMonthHeader = True
    Else
        TryNormalizeMonthHeader = TryNormalizeMonth(rawHeader, resultMonth)
    End If
End Function

Public Function LookupInternalPart(ByVal customer As String, ByVal customerPart As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = EnsureSheet("PartMaster")
    lastRow = LastUsedRow(ws, 1)

    For r = 2 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, 1).Value)), customer, vbTextCompare) = 0 _
           And StrComp(Trim$(CStr(ws.Cells(r, 2).Value)), customerPart, vbTextCompare) = 0 Then
            LookupInternalPart = Trim$(CStr(ws.Cells(r, 3).Value))
            Exit Function
        End If
    Next r
End Function

Public Sub AppendNormalizedForecast(ByVal customer As String, _
                                    ByVal internalPart As String, _
                                    ByVal customerPart As String, _
                                    ByVal forecastMonth As Date, _
                                    ByVal latestForecastQty As Double, _
                                    ByVal sourceFile As String, _
                                    ByVal rowStatus As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = EnsureSheet("StandardizedData")
    EnsureStandardHeaders ws

    nextRow = LastUsedRow(ws, 1) + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = customer
    ws.Cells(nextRow, 2).Value = internalPart
    ws.Cells(nextRow, 3).Value = customerPart
    ws.Cells(nextRow, 4).Value = forecastMonth
    ws.Cells(nextRow, 4).NumberFormat = "yyyy-mm"
    ws.Cells(nextRow, 5).Value = latestForecastQty
    ws.Cells(nextRow, 10).Value = sourceFile
    ws.Cells(nextRow, 11).Value = Now
    ws.Cells(nextRow, 11).NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Cells(nextRow, 12).Value = rowStatus
End Sub

Public Sub EnsureStandardHeaders(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim i As Long

    If Len(ws.Cells(1, 1).Value) > 0 Then Exit Sub

    headers = Array("Customer", "InternalPartNo", "CustomerPartNo", "ForecastMonth", _
                    "LatestForecastQty", "ActualShipmentQty", "YearlyForecastQty", _
                    "VarianceVsYearly", "VariancePct", "Source", "UpdatedAt", "Status")

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
End Sub

Public Sub LogImport(ByVal logTime As Date, ByVal customer As String, ByVal fileName As String, _
                     ByVal importedRows As Long, ByVal warnings As Long, ByVal statusText As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = EnsureSheet("ImportLog")

    If Len(ws.Cells(1, 1).Value) = 0 Then
        ws.Range("A1:F1").Value = Array("Timestamp", "Customer", "File", "Rows Imported", "Warnings", "Status")
    End If

    nextRow = LastUsedRow(ws, 1) + 1
    ws.Cells(nextRow, 1).Value = logTime
    ws.Cells(nextRow, 1).NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Cells(nextRow, 2).Value = customer
    ws.Cells(nextRow, 3).Value = fileName
    ws.Cells(nextRow, 4).Value = importedRows
    ws.Cells(nextRow, 5).Value = warnings
    ws.Cells(nextRow, 6).Value = statusText
End Sub

Public Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Public Function LastUsedRow(ByVal ws As Worksheet, ByVal keyColumn As Long) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, keyColumn).End(xlUp).Row
    If LastUsedRow = 1 And Len(ws.Cells(1, keyColumn).Value) = 0 Then LastUsedRow = 0
End Function
