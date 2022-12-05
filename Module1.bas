Attribute VB_Name = "Module1"
Option Explicit
Option Private Module

Public Sub axis_adjust_6cht()
Dim cht As ChartObject
Dim srs As Series
Dim FirstTime  As Boolean
Dim MaxNumber As Double
Dim MinNumber As Double
Dim MaxChartNumber As Double
Dim MinChartNumber As Double
Dim Padding As Double
Dim i As Long
Dim ws As Variant

On Error GoTo Handler:

'Optimize code
Application.ScreenUpdating = False
Padding = 0
'Input Padding on Top of Min/Max Numbers (Percentage)
Padding = InputBox("Padding (Between 0-1)") 'Number between 0-1
If Padding > 1 Then GoTo Handler:




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
             cht.Chart.Axes(xlValue).MaximumScale = MaxChartNumber * (1 + Padding)
         
         Next cht

MsgBox ws.Name + "'s Axes have been reset with " + CStr(Padding) + " Padding"
Next ws

'Optimize Code
  Application.ScreenUpdating = True
Exit Sub
Handler:
MsgBox "Padding should be a number < 1 "

End Sub





