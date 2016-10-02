Sub testEuler()

Dim x, y, cnt As Integer
Dim cell, rng as Range

Cells.Delete

y = 'add number here

Set rng = ActiveSheet.Range("A1:A(add number here)")

'generate numbers
For x = 1 To y
    Cells(x, 1) = x
Next x

'color numbers divisible by 3 or 5
    For x = 1 To y
        If Cells(x, 1) Mod 3 = 0 Or Cells(x, 1) Mod 5 = 0 Then
            Cells(x, 1).Interior.Color = RGB(255, 255, 0)
        End If
    Next x
	
'count yellow cells
For Each cell In rng
    If cell.Interior.Color = RGB(255, 255, 0) Then
        cnt = cnt + 1
    End If
Next	
	
'delete numbers not colored yellow
Do Until IsEmpty(Cells(cnt + 1, "A"))
    For Each cell In rng
        If cell.Interior.Color <> RGB(255, 255, 0) Then
            cell.Delete shift:=xlUp
        End If
    Next cell
Loop	

End Sub