Attribute VB_Name = "modImportAdapters"
Option Explicit

Public Function ImportForecastFile(ByVal filePath As String, ByVal customer As String) As Long
    Dim srcBook As Workbook
    Dim srcSheet As Worksheet
    Dim partCol As Long
    Dim dateCol As Long
    Dim qtyCol As Long

    On Error GoTo Fail

    Set srcBook = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    Set srcSheet = srcBook.Worksheets(1)

    partCol = FirstExistingHeader(srcSheet, Array("Customer PN", "Part No", "Part Number", "Item", "料號"))
    dateCol = FirstExistingHeader(srcSheet, Array("Forecast Date", "Month", "Forecast Month", "預測月份"))
    qtyCol = FirstExistingHeader(srcSheet, Array("Qty", "Forecast Qty", "Quantity", "預測數量"))

    If partCol = 0 Then Err.Raise vbObjectError + 100, , "Part-number header not found."

    If dateCol > 0 And qtyCol > 0 Then
        ImportForecastFile = ImportLongFormat(srcSheet, customer, partCol, dateCol, qtyCol, filePath)
    Else
        ImportForecastFile = ImportMatrixFormat(srcSheet, customer, partCol, filePath)
    End If

CleanExit:
    If Not srcBook Is Nothing Then srcBook.Close SaveChanges:=False
    Exit Function

Fail:
    If Not srcBook Is Nothing Then srcBook.Close SaveChanges:=False
    Err.Raise Err.Number, "ImportForecastFile", Err.Description
End Function

Private Function ImportLongFormat(ByVal ws As Worksheet, _
                                  ByVal customer As String, _
                                  ByVal partCol As Long, _
                                  ByVal dateCol As Long, _
                                  ByVal qtyCol As Long, _
                                  ByVal sourcePath As String) As Long
    Dim lastRow As Long
    Dim r As Long
    Dim customerPart As String
    Dim internalPart As String
    Dim forecastMonth As Date
    Dim qty As Variant
    Dim status As String

    lastRow = LastUsedRow(ws, partCol)

    For r = 2 To lastRow
        customerPart = Trim$(CStr(ws.Cells(r, partCol).Value))
        qty = ws.Cells(r, qtyCol).Value

        If Len(customerPart) > 0 And IsNumeric(qty) Then
            If TryNormalizeMonth(ws.Cells(r, dateCol).Value, forecastMonth) Then
                internalPart = LookupInternalPart(customer, customerPart)
                status = IIf(Len(internalPart) = 0, "UNMAPPED_PART", "OK")
                AppendNormalizedForecast customer, internalPart, customerPart, forecastMonth, CDbl(qty), Dir$(sourcePath), status
                ImportLongFormat = ImportLongFormat + 1
            End If
        End If
    Next r
End Function

Private Function ImportMatrixFormat(ByVal ws As Worksheet, _
                                    ByVal customer As String, _
                                    ByVal partCol As Long, _
                                    ByVal sourcePath As String) As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Long
    Dim customerPart As String
    Dim internalPart As String
    Dim forecastMonth As Date
    Dim qty As Variant
    Dim status As String

    lastRow = LastUsedRow(ws, partCol)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For r = 2 To lastRow
        customerPart = Trim$(CStr(ws.Cells(r, partCol).Value))
        If Len(customerPart) = 0 Then GoTo NextRow

        internalPart = LookupInternalPart(customer, customerPart)
        status = IIf(Len(internalPart) = 0, "UNMAPPED_PART", "OK")

        For c = 1 To lastCol
            If c <> partCol Then
                If TryNormalizeMonthHeader(ws.Cells(1, c).Value, forecastMonth) Then
                    qty = ws.Cells(r, c).Value
                    If IsNumeric(qty) And Len(qty) > 0 Then
                        AppendNormalizedForecast customer, internalPart, customerPart, forecastMonth, CDbl(qty), Dir$(sourcePath), status
                        ImportMatrixFormat = ImportMatrixFormat + 1
                    End If
                End If
            End If
        Next c
NextRow:
    Next r
End Function
