Sub heatmap()
'variables used
Dim x, y As Double
Dim z As Integer

'set z here
z = 20

'execute script using for loop
For x = 1 To z
For y = 1 To z
    Cells(x, y) = Math.Sin(x) + Math.Cos(y)
        'conditional logic
        If Cells(x, y) > 0 Then
            Cells(x, y).Interior.Color = RGB(0, 255, 0)
        Else
            Cells(x, y).Interior.Color = RGB(255, 0, 0)
        End If
Next y
Next x

End Sub