Attribute VB_Name = "Module1"
Sub stock_data()

'Declare ws as Worksheet
Dim ws As Worksheet
        
'For loop to run Macro on all worksheets
For Each ws In ThisWorkbook.Worksheets
    ws.Select


'Declare and assign variables for table summary, lastrow, etc.
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Stock_Summary_Table As Integer
Stock_Summary_Table = 2

Dim Open_Price As Double
Open_Price = Cells(2, 3)
Dim Close_Price As Double

Dim Lastrow As Long
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row


'Output headers for tables
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'For loop to iterate through data and output ticker with variable values to summary table
For i = 2 To Lastrow
    'If block to compare ticker column cells to group by ticker and output variables to summary table
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Close_Price = Cells(i, 6)
        Ticker = Cells(i, 1)
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7)
        Yearly_Change = Close_Price - Open_Price
        Percent_Change = ((Close_Price - Open_Price) / Open_Price)
        Range("I" & Stock_Summary_Table).Value = Ticker
        Range("J" & Stock_Summary_Table).Value = Yearly_Change
        Range("K" & Stock_Summary_Table).Value = Percent_Change
        Range("L" & Stock_Summary_Table).Value = Total_Stock_Volume
        Open_Price = Cells(i + 1, 3)
        Stock_Summary_Table = Stock_Summary_Table + 1
        Total_Stock_Volume = 0
    Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    End If
Next i
    

'Declare and assign Endrow for last row count
Dim Endrow As Long
Endrow = Cells(Rows.Count, "J").End(xlUp).Row

'For loop to iterate through Yearly Change column to set colors to red, green, or white
For i = 2 To Endrow
    Set Yearly_Change = Range("J" & i)
    If Yearly_Change < 0 Then
        Yearly_Change.Interior.Color = vbRed
    ElseIf Yearly_Change > 0 Then
        Yearly_Change.Interior.Color = vbGreen
    Else
        Yearly_Change.Interior.Color = vbWhite
    End If
Next i


'Declare and assign Countrow & Countrow2 for last row count
Dim Countrow As Long
Countrow = Cells(Rows.Count, "K").End(xlUp).Row
Dim Countrow2 As Long
Countrow2 = Cells(Rows.Count, "L").End(xlUp).Row

'Get max & min percentages for percent change and max from total stock volume
Range("Q2").Value = WorksheetFunction.Max(Range("K2:K" & Countrow))
Range("Q3").Value = WorksheetFunction.Min(Range("K2:K" & Countrow))
Range("Q4").Value = WorksheetFunction.Max(Range("L2:L" & Countrow2))

'Assign variables for max percent increase, decrease, and volume
Greatest_Percent_Increase = Range("Q2").Value
Greatest_Percent_Decrease = Range("Q3").Value
Greatest_Total_Volume = Range("Q4").Value

'For loop to iterate through column K to get biggest percent increase and decrease
For i = 2 To Countrow
    If Cells(i, 11) = Greatest_Percent_Increase Then
        Cells(2, 16).Value = Cells(i, 9)
    End If
    
    If Cells(i, 11) = Greatest_Percent_Decrease Then
        Cells(3, 16).Value = Cells(i, 9)
    End If
Next i
  
'For loop to iterate through column J to get biggest stock volume
For i = 2 To Countrow2
    If Cells(i, 12) = Greatest_Total_Volume Then
        Cells(4, 16).Value = Cells(i, 9)
    End If
Next i

'Format greatest percent increase and decrease to percentage with two decimals
Range("Q2").Value = FormatPercent(Range("Q2"))
Range("Q3").Value = FormatPercent(Range("Q3"))

'Format columns 10 & 11 to two decimals and percentage for 11
Columns(10).NumberFormat = "0.00"
Columns(11).NumberFormat = "0.00%"

Next ws

End Sub
