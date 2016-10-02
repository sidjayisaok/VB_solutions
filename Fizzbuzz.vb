'FizzBuzz algorithm for multiples of 3 and 5
Sub FizzBuzz()

'variables used
dim x, y as integer

'input box
y=Application.InputBox("Enter number here")

'clear field
cells.delete

'Fizzbuzz logic here
For x=1 to y
	Cells(x,1)=x
		if Cells(x,1) mod 3 = 0 and Cells(x,1) mod 5=0 then
			Cells(x,1)="FizzBuzz"
		elseif Cells(x,1) mod 3 = 0 then
			Cells(x,1)="Fizz"
		elseif Cells(x,1) mod 5 = 0 then
			Cells(x,1)="Buzz"
		end if
next x		

End Sub