Attribute VB_Name = "Module5"
Option Explicit
Option Private Module
Dim cht As ChartObject
Dim srs As Series
Dim FirstTime  As Boolean
Dim MaxNumber As Double
Dim MinNumber As Double
Dim MaxChartNumber As Double
Dim MinChartNumber As Double
Dim i As Long
Dim ws As Variant
Dim answer As Boolean
Public Sub Call_Button()
Dim x As Long


'Go to page 7 and activate axis adjust 6 cht
ThisWorkbook.Sheets("Page 7").Activate
ActiveSheet.Range("A1").Activate
Application.Wait (Now + TimeValue("0:00:01"))
    Call axis_adjust_6cht
'Go to page 10 and activate axis adjust 6 cht
ThisWorkbook.Sheets("Page 10").Activate
ActiveSheet.Range("A1").Activate
Application.Wait (Now + TimeValue("0:00:01"))
    Call axis_adjust_4cht
'Go to page 13 and activate axis adjust 6 cht
ThisWorkbook.Sheets("Page 13").Activate
ActiveSheet.Range("A1").Activate
Application.Wait (Now + TimeValue("0:00:01"))
    Call Ageing_chart
'Go to sheet 1
ThisWorkbook.Sheets(1).Activate
ActiveSheet.Range("A1").Activate

'Select relavent sheets and decide if a print is required
answer = MsgBox("Save as PDF?", vbYesNo)
If answer = vbYes Then
    Sheets(1).Select
    For x = 1 To ThisWorkbook.Sheets.Count
        If Sheets(x).Name Like "Page #" Or Sheets(x).Name Like "Page ##" Then Sheets(x).Select Replace:=False
    Next x
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, PrintToFile:=True, Collate _
            :=True, IgnorePrintAreas:=False
Else
    Sheets(1).Select
    For x = 1 To ThisWorkbook.Sheets.Count
        If Sheets(x).Name Like "Page #" Or Sheets(x).Name Like "Page ##" Then Sheets(x).Select Replace:=False
    Next x
End If

End Sub


'Adjust chart axis each chart on the sheet with an additional buffer inputed by the user.
'Charts in one line(row) will keep the same axis

Public Sub axis_adjust_6cht()
Dim sixPadding As Double


On Error GoTo Handler:

'Optimize code
Application.ScreenUpdating = False
sixPadding = 0
'Input Padding on Top of Min/Max Numbers (Percentage)
sixPadding = InputBox("Padding (Between 0-1)") 'Number between 0-1
If sixPadding > 1 Then GoTo Handler:


'Loop Throgh Worksheets with specific names
For Each ws In Sheets(Array("Page 7", "Page 8", "Page 9"))
ws.Activate
i = 0
        'Loop Through Each Chart On ActiveSheet
         For Each cht In ActiveSheet.ChartObjects
           
           i = i + 1
           
        'Only calculate a max number if chart number is one or three
           Select Case i
               Case 2, 3: GoTo OverCalculation
               Case Is > 4: GoTo OverCalculation
           End Select
           
           'First Time Looking at This Chart?
            FirstTime = True
             
           'Determine Chart's Overall Max/Min From Connected Data Source
             For Each srs In cht.Chart.FullSeriesCollection
               'Determine Maximum value in Series
                 MaxNumber = Application.WorksheetFunction.Max(srs.Values)
               
               'Store value if currently the overall Maximum Value
                 If FirstTime = True Then
                   MaxChartNumber = MaxNumber
                 ElseIf MaxNumber > MaxChartNumber Then
                   MaxChartNumber = MaxNumber
                 End If
               
               'First Time Looking at This Chart?
                 FirstTime = False
             Next srs
OverCalculation:
           'Rescale Y-Axis for all charts
             cht.Chart.Axes(xlValue).MinimumScale = 0
             cht.Chart.Axes(xlValue).MaximumScale = MaxChartNumber * (1 + sixPadding)
         
         Next cht
'Output confirmation
MsgBox ws.Name + "'s Axes have been reset with " + CStr(sixPadding) + " Padding"
Next ws

'Optimize Code
  Application.ScreenUpdating = True
Exit Sub
Handler:
MsgBox "Padding should be a number < 1 "

End Sub

'Adjust chart axis each chart on the sheet with an additional buffer (padding) inputed by the user.
'Charts in one line(row) will keep the same axis

Public Sub axis_adjust_4cht()
Dim fourPadding As Double


On Error GoTo Handler:

Application.ScreenUpdating = False
fourPadding = 0
'Input Padding on Top of Min/Max Numbers (Percentage)
fourPadding = InputBox("Padding (Between 0-1)") 'Number between 0-1

If fourPadding > 1 Then GoTo Handler:


For Each ws In Sheets(Array("Page 10", "Page 11", "Page 12"))
ws.Activate
i = 0
'Loop Through Each Chart On ActiveSheet
        For Each cht In ActiveSheet.ChartObjects
           
           i = i + 1

    
            Select Case i
                Case 2: GoTo OverCalculation
                Case 4: GoTo OverCalculation
            End Select
    
         'First Time Looking at This Chart?
             FirstTime = True
             
           'Determine Chart's Overall Max/Min From Connected Data Source
             For Each srs In cht.Chart.FullSeriesCollection
               'Determine Maximum value in Series
                 MaxNumber = Application.WorksheetFunction.Max(srs.Values)
               
               'Store value if currently the overall Maximum Value
                 If FirstTime = True Then
                   MaxChartNumber = MaxNumber
                 ElseIf MaxNumber > MaxChartNumber Then
                   MaxChartNumber = MaxNumber
                 End If
               
        
               
               'First Time Looking at This Chart?
                 FirstTime = False
             Next srs
OverCalculation:

           'Rescale Y-Axis
             cht.Chart.Axes(xlValue).MinimumScale = 0
             cht.Chart.Axes(xlValue).MaximumScale = MaxChartNumber * (1 + fourPadding)

         Next cht
'Output confirmation
MsgBox ws.Name + "'s Axes have been reset with " + CStr(fourPadding) + " Padding"
Next ws


'Optimize Code
  Application.ScreenUpdating = True
Exit Sub
Handler:
MsgBox "Padding should be a number < 1"

End Sub

'Keep the bar chart width the same for all charts in specific sheets

Public Sub Ageing_chart()

Dim Count As Long
Dim ArrSheet As Variant


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
        'Count the number of bars in the chart
            Count = ActiveChart.SeriesCollection(1).Points.Count
            
        'Set gap width (according to stake holder preference) according to count
            If Count = 2 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 500
            ElseIf Count = 3 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 425
            ElseIf Count = 4 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 325
            ElseIf Count = 5 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 250
            ElseIf Count = 6 Then
                ActiveChart.Axes(xlValue).MajorGridlines.Select
                ActiveChart.FullSeriesCollection(1).Select
                ActiveChart.ChartGroups(1).GapWidth = 220
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

Public Sub butterflychart()
Dim cht As ChartObject
Dim ws As Variant
Dim r As Integer
Dim i As Integer

On Error Resume Next


For Each ws In Sheets(Array("Page 5", "Page 6"))
ws.Activate
i = 0
If ws.Name = "Page 6" Then i = -1
    For Each cht In ActiveSheet.ChartObjects
        i = i + 1
        r = i Mod 2
        cht.Activate
        Select Case r
            Case Is = 0: ActiveChart.Axes(xlValue).ReversePlotOrder = True
        End Select
        ActiveChart.Axes(xlCategory).ReversePlotOrder = True
        ActiveChart.Axes(xlValue).MinimumScale = 0
        ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
    Next cht
Next ws
End Sub




