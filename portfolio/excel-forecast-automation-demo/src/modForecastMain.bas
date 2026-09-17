Attribute VB_Name = "modForecastMain"
Option Explicit

' Portfolio reference implementation.
' Intended to demonstrate structure and error handling for a production .xlsm tool.

Public Sub ImportForecastDemo()
    Dim filePath As String
    Dim customer As String
    Dim importedRows As Long
    Dim startedAt As Date

    On Error GoTo Fail

    filePath = PickExcelFile()
    If Len(filePath) = 0 Then Exit Sub

    customer = Trim$(InputBox("Customer name used by PartMaster:", "Import Forecast"))
    If Len(customer) = 0 Then Exit Sub

    startedAt = Now
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Importing forecast..."

    importedRows = ImportForecastFile(filePath, customer)
    RecalculateVarianceAndStatus
    LogImport startedAt, customer, Dir$(filePath), importedRows, 0, "SUCCESS"

    MsgBox CStr(importedRows) & " normalized forecast rows imported.", vbInformation

CleanExit:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    LogImport startedAt, customer, Dir$(filePath), importedRows, 1, "FAILED: " & Err.Description
    MsgBox "Import failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub

Public Sub RecalculateVarianceAndStatus()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim latestQty As Variant
    Dim baselineQty As Variant
    Dim pct As Double

    Set ws = EnsureSheet("StandardizedData")
    lastRow = LastUsedRow(ws, 1)

    For r = 2 To lastRow
        latestQty = ws.Cells(r, 5).Value
        baselineQty = ws.Cells(r, 7).Value

        If IsNumeric(latestQty) And IsNumeric(baselineQty) And Len(baselineQty) > 0 Then
            ws.Cells(r, 8).Value = CDbl(latestQty) - CDbl(baselineQty)

            If CDbl(baselineQty) <> 0 Then
                pct = ws.Cells(r, 8).Value / CDbl(baselineQty)
                ws.Cells(r, 9).Value = pct
                ws.Cells(r, 9).NumberFormat = "0.0%"
            Else
                ws.Cells(r, 9).ClearContents
                pct = 0
            End If

            If Abs(pct) >= 0.1 Then
                ws.Cells(r, 12).Value = "CHECK"
            ElseIf Len(ws.Cells(r, 12).Value) = 0 Then
                ws.Cells(r, 12).Value = "OK"
            End If
        End If
    Next r
End Sub

Public Sub RefreshForecastReport()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    RecalculateVarianceAndStatus
    ThisWorkbook.RefreshAll
    Application.CalculateFull
    MsgBox "Integration and report refresh completed.", vbInformation

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    MsgBox "Refresh failed: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
