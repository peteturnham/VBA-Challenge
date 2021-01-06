Sub test_data()
'initializing ticker variable
Dim ticker As String
'initializing a variable for open_price
Dim open_price As Double
'initializing variable for close_price
Dim close_price As Double
'initializing a counter variable
Dim count As Integer
count = 2
'intitialize var as long for stock volume
Dim vol As LongLong
'initializing variable to store yearly change
Dim yearly_change As Double
'initialize a variable for the whole sheets row total
lastrow = Cells(Rows.count, 1).End(xlUp).Row
'variable for percent change
Dim change_percent As Double
open_price = Cells(2, 3).Value  'hardcoding open_price variable to have initial value

'looping through stock tickers
For i = 2 To lastrow


    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then      'check if in the same symbol, if not...

        Range("I" & count).Value = Cells(i, 1).Value        'printing new symbol

        vol = vol + Cells(i, 7).Value       'store and add volume for each loop
        
        Range("L" & count).Value = vol      'print stock total volume for each stock

        vol = 0     'reset the volume to be 0

        close_price = Cells(i, 6).Value     'get closing price
        
        year_change = close_price - open_price      'find yearly stock change
        
        Range("J" & count).Value = year_change       'print yearly price change
        
        change_percent = ((close_price - open_price) / open_price)      'find yearly perctentage change
        
        Range("K" & count).Value = Format(change_percent, "Percent")       'convert cells to percent
        
        year_change = close_price - open_price      'find yearly stock change
        
        Range("J" & count).Value = year_change       'print yearly price change
        
        open_price = Cells(i + 1, 3).Value     'reset open_price to one after

       count = count + 1       'add 1 to count
        
    

    ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then       'checking if we are in the same value, if so...

        vol = vol + Cells(i, 7).Value       'store and add volume for each loop

        Range("L" & count).Value = vol      'print stock total volume for each stock
    
            
            End If
Next i
Call color_coat
End Sub

Sub color_coat()

lastrow = Cells(Rows.count, 9).End(xlUp).Row        'initialize a variable for the whole sheets row total

For i = 2 To lastrow        'loop through all stocks
    If Cells(i, 11).Value < 0 Then      'if cell value is less than 0
        Cells(i, 11).Interior.ColorIndex = 3        'change cell color red
    ElseIf Cells(i, 10).Value > 0 Then      'if cell value is greater than 0
        Cells(i, 11).Interior.ColorIndex = 4        'change color green
    
    End If
Next i
    
End Sub


