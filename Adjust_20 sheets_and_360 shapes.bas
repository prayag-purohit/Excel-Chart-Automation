Attribute VB_Name = "Module2"
'The VBA code tries to format shape colors depending upon the change in values over a period.
'The values are entered in a vertical fashion within the sheet. the VBA cycles through values in chronological order and stores them in variables
'values are stored as 'x' and 'y', and the value of x-y is stored in 'subtract'.
'Color codes are stored depending upon the value of 'subtract'

Dim color As Long
Dim i As Integer
Dim shp As shape



Sub returnvalues()


i = 0
ActiveSheet.Range("C13:E15").Activate 'select starting cell
    
For i = 1 To 18 'there are 18 values total for which shapes need to be formatted
    
    Trans = 0 ' Transparancy as zero for every iteration
    x = ActiveCell.Value
    If x = "- " Then ' If active cell contains text value, mark it as 0
        x = 0
    End If
    
    Selection.End(xlDown).Activate
    y = ActiveCell.Value ' cell value for cell directly below cell 'x'

    subtract = x - y
    Selection.End(xlUp).Activate ' back to cell x
    'Debug.Print subtract
    'Debug.Print i
    Selection.End(xlToRight).Activate ' go to right and  now that cell becomes x
    
        'store color and transparancy information as per the change
        If x = 0 And y = 0 Then
            Trans = 1
            color = RGB(255, 255, 255) 'white
        ElseIf subtract < -1 Then
            color = RGB(255, 0, 0) 'red
        ElseIf subtract > 1 Then
            color = RGB(38, 175, 67) 'green
        Else:
            color = RGB(127, 127, 127) 'gray
        End If

        'nme = CStr(i)
        'Debug.Print nme
Set shp = ActiveSheet.Shapes(CStr(i)) 'shapes are named from 1 to 18. they are activated as per the iteration cycle

        With shp.Fill 'shapes are filled as per the analytics team requirement
            .Visible = msoTrue
            .ForeColor.RGB = color
            .Transparency = Trans
            .Solid
        End With
        
Next i
    
    Range("A1").Activate
    
End Sub
    


