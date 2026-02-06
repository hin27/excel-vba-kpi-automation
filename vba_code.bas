Attribute VB_Name = "Module1"
Sub UpdateChartsAndSendEmail()

    Dim wsData As Worksheet, wsChart As Worksheet
    Dim lastRow As Long
    Dim outlookApp As Object, outlookMail As Object
    Dim tempWorkbook As Workbook
    Dim chartSheets As Collection
    Dim timestamp As String, savePath As String
    Dim i As Long

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsData = ThisWorkbook.Sheets("Output sheet")
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    ' ===============================
    ' DELETE OLD CHART SHEETS
    ' ===============================
    For Each wsChart In ThisWorkbook.Sheets
        If Right(wsChart.Name, 6) = " Chart" _
        Or Right(wsChart.Name, 8) = "Chart BH" Then
            wsChart.Delete
        End If
    Next wsChart

    ' ===============================
    ' CREATE DAILY CHARTS (D ? K)
    ' ===============================
    Call CreateCharts(wsData, lastRow, 4, 11, "Daily", " Chart")

    ' ===============================
    ' CREATE BH CHARTS (L ? S)
    ' ===============================
    Call CreateCharts(wsData, lastRow, 12, 19, "BH", " Chart BH")

    ' ===============================
    ' COLLECT CHART SHEETS
    ' ===============================
    Set chartSheets = New Collection
    For Each wsChart In ThisWorkbook.Sheets
        If Right(wsChart.Name, 6) = " Chart" _
        Or Right(wsChart.Name, 8) = "Chart BH" Then
            chartSheets.Add wsChart.Name
        End If
    Next wsChart

    If chartSheets.Count = 0 Then
        MsgBox "No chart sheets found.", vbExclamation
        Exit Sub
    End If

    ' ===============================
    ' EXPORT TO NEW EXCEL FILE
    ' ===============================
    Sheets(chartSheets(1)).Copy
    Set tempWorkbook = ActiveWorkbook

    For i = 2 To chartSheets.Count
        ThisWorkbook.Sheets(chartSheets(i)).Copy _
            After:=tempWorkbook.Sheets(tempWorkbook.Sheets.Count)
    Next i

    timestamp = Format(Now, "yyyy-mm-dd")
    savePath = ThisWorkbook.Path & "\KPI_ChartSheets_" & timestamp & ".xlsx"

    tempWorkbook.SaveAs savePath, xlOpenXMLWorkbook
    tempWorkbook.Close False

    ' ===============================
    ' SEND EMAIL
    ' ===============================
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then Set outlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If Not outlookApp Is Nothing Then
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .To = "hindiya.ahmed.ext@nokia.com"
            .Subject = "Updated KPI Chart Sheets - " & timestamp
            .Body = "Hi Hindiya," & vbCrLf & vbCrLf & _
                    "Please find attached the updated KPI chart sheets (Daily & BH)." & vbCrLf & vbCrLf & _
                    "Regards," & vbCrLf & "Automated Report"
            .Attachments.Add savePath
            .Send
        End With
        MsgBox "? Daily & BH charts created and email sent.", vbInformation
    Else
        MsgBox "? Outlook not available. File saved to:" & vbCrLf & savePath, vbExclamation
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub


' =====================================================
' GENERIC CHART CREATOR (DAILY / BH)
' =====================================================
Sub CreateCharts(wsData As Worksheet, lastRow As Long, _
                 colStart As Long, colEnd As Long, _
                 valuePrefix As String, sheetSuffix As String)

    Dim wsChart As Worksheet
    Dim dictTechs As Object, techKey As Variant
    Dim chartTop As Long, chartLeft As Long
    Dim valueRange As Range
    Dim chtObj As ChartObject
    Dim chartTitleText As String
    Dim xLabels() As String
    Dim xColCount As Long
    Dim i As Long, j As Long

    xColCount = colEnd - colStart + 1
    ReDim xLabels(1 To xColCount)

    ' X-axis labels
    For j = 1 To xColCount
        xLabels(j) = Replace(wsData.Cells(1, colStart + j - 1).Value, valuePrefix & "_", "")
    Next j

    ' Collect technologies
    Set dictTechs = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        techKey = Trim(wsData.Cells(i, 1).Value)
        If techKey <> "" And Not dictTechs.exists(techKey) Then
            dictTechs.Add techKey, techKey
        End If
    Next i

    ' Create charts per technology
    For Each techKey In dictTechs.Keys

        Set wsChart = Sheets.Add(After:=Sheets(Sheets.Count))
        wsChart.Name = techKey & sheetSuffix

        chartTop = 20: chartLeft = 20

        For i = 2 To lastRow
            If Trim(wsData.Cells(i, 1).Value) = techKey Then

                Set valueRange = wsData.Range(wsData.Cells(i, colStart), wsData.Cells(i, colEnd))
                chartTitleText = Replace(wsData.Cells(i, 2).Value, valuePrefix & "_", "")

                Set chtObj = wsChart.ChartObjects.Add(chartLeft, chartTop, 500, 250)
                With chtObj.Chart
                    .ChartType = xlLineMarkers
                    .SeriesCollection.NewSeries
                    .SeriesCollection(1).Values = valueRange
                    .SeriesCollection(1).XValues = xLabels
                    .HasLegend = False
                    .HasTitle = True
                    .ChartTitle.Text = chartTitleText & " (" & valuePrefix & ")"
                End With

                chartTop = chartTop + 270
                If chartTop > 1000 Then
                    chartTop = 20
                    chartLeft = chartLeft + 520
                End If
            End If
        Next i
    Next techKey

End Sub


