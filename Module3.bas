Attribute VB_Name = "Module3"
Option Explicit
Option Private Module

Public Sub Ageing_chart()

Dim cht As ChartObject
Dim srs As Series
Dim FirstTime As Boolean
Dim Count As Long
Dim ArrSheet As Variant
Dim ws As Variant

'Optimize code
Application.ScreenUpdating = False



'Select sheets that need to be looped to
ArrSheet = Array("Page 13", "Page 14", "Page 15")
'Loop through sheets
For Each ws In Sheets(ArrSheet)

        'Loop through chart objects, set maximum scale, and minimum scale.
        For Each cht In ws.ChartObjects
            cht.Activate
            cht.Chart.Axes(xlValue).MaximumScaleIsAuto = True
            cht.Chart.Axes(xlValue).MinimumScale = 0
        'Count the number of data points in the chart
            Count = ActiveChart.SeriesCollection(1).Points.Count
            
        'Set gap width according to counts
            If Count = 2 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 500
            ElseIf Count = 3 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 400
            ElseIf Count = 4 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 300
            ElseIf Count = 5 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 225
            ElseIf Count = 6 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 185
            ElseIf Count > 6 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 150
            End If
        Next cht
        MsgBox ws.Name + "'s gap width has been reset"

Next ws

Application.ScreenUpdating = True

End Sub



