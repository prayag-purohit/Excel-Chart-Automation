Attribute VB_Name = "Module4"
Private Sub Button()
Dim ws As Worksheet
Dim x As Long

ThisWorkbook.Sheets("Page 7").Activate
ActiveSheet.Range("A1").Activate
Application.Wait (Now + TimeValue("0:00:01"))
    Call axis_adjust_6cht
ThisWorkbook.Sheets("Page 10").Activate
ActiveSheet.Range("A1").Activate
Application.Wait (Now + TimeValue("0:00:01"))
    Call axis_adjust_4cht
ThisWorkbook.Sheets("Page 13").Activate
ActiveSheet.Range("A1").Activate
Application.Wait (Now + TimeValue("0:00:01"))
    Call Ageing_chart

ThisWorkbook.Sheets(1).Activate
ActiveSheet.Range("A1").Activate

answer = MsgBox("Save as PDF?", vbYesNo)
If answer = vbYes Then
    Sheets(2).Select
    For x = 2 To ThisWorkbook.Sheets.Count
        If Sheets(x).Name <> "Sheet1" Then Sheets(x).Select Replace:=False
    Next x
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, PrintToFile:=True, Collate _
            :=True, IgnorePrintAreas:=False
Else
    ThisWorkbook.Sheets(1).Activate
    ActiveSheet.Range("A1").Activate
End If

End Sub


