Sub ticker()
'Create a script that will loop through each year of stock data and grab the total amount of 
'volume each stock had over the year.
'You will also need to display the ticker symbol to coincide with the total volume.


'For each worksheet, look at the value of column A and if column A is the same, go to column 
'G and add that number to the total. Then print somewhere.

'Define variables:

Dim LastRow As Double
Dim i As Double
Dim total As Double
Dim x As Double

LastRow = Cells(Rows.Count, 1).End(xlUp).Row
total = 0
x = 2

'Create a loop to apply code to each row in column A
For i = 2 to LastRow

'For each row, compare the cell value to the previous cell value.
    If Cells(i, "A").Value = Cells(i+1, "A").Value Then
    'If the values are the same then look at column G for the value in that cell
    'Add the value in Column G to the total (which will start at zero)

        total = Cells(i, "G").value + total
        'print the total and the value of column A to two new columns
        'Cells(i, "K").value = total
    'If the value is NOT the same. Restart the total to zero and start counting again.
    Else 
        
        total = Cells(i, "G").value + total
        Cells(x, "K").value = total
        Cells(x, "J").value = Cells(i, "A").Value
        x = x + 1
        total = 0
        
    End If


'End the loop
Next i
End Sub