'fibonacci sequence in Excel
Sub Fibonacci()

Dim Fib as Long
Dim x as Long
Dim y as Long
Dim first as integer
Dim second as integer
Dim sum as integer

Fib=Application.InputBox("Enter number here")

'clear the rows
cells.delete

first = 0
second = 1
sum = 0

if Cells(x,1) = 0 Then



'Fibonacci logic
For x=1 to Fib
    if Fib = 0 then
        Cells(x,1) = 1
    else

    end if        
next x

End Sub


'example
Function Fibonacci(Fib As Long, x As Long) As Long    
    Dim firstNum As Long
    Dim secondNum As Long
    Dim sum As Long
    Dim i As Long
    
    first = 0
    second = 1
    sum = 0
    
    If n = 0 Then
        Fib = first
    ElseIf n = 1 Then
        Fib = second
    Else
        For i = 2 To n
            sum = first + second
            first = second
            second = sum
        Next i
        Fib = sum
    End If
    
End Function


'Fibonacci List Generator
Sub FibonacciList()

dim x, y as integer

'input box
y=Application.InputBox("Enter number here")

cells.delete

'generate numbers
For x = 1 To y
    'Fibonacci formula here
	Cells(x, 1) = Math.Round(((1+5^.5)^x-(1-5^.5)^x)/((2^x)*5^.5),0)
Next x	

End Sub



