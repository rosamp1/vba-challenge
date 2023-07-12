# vba-challenge

Link to Excel file below as it is too large to upload
https://drive.google.com/file/d/1jUxpHpfU9RmqHj5IFXQlcUQABnUWpxcO/view?usp=drive_link 

Code written for this challenge below:

Sub vba_challenge_yearly_change()

 For Each ws In Worksheets
    
    'Ticker Name
    Dim Category_Name As String
    
    ' Yearly Change
    Dim Yearly_Change_Total As Double
    Yearly_Change_Total = 0
    
    'Percent Change
    Dim Per_Change As Double
    Per_Change = 0
    
    'Total Stock Volume
    Dim Total_Inv_Volume As Double
    Total_Inv_Volume = 0
      
    Dim Summary_Table_Row As Integer
     
    'Formula for it to run to last row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     
    Summary_Table_Row = 2
    
    'Add Column Headers
    ws.Range("I1") = "Ticker"
    ws.Range("J1") = "Yearly Change"
    ws.Range("K1") = "Percent Change"
    ws.Range("L1") = "Total Stock Volume"
    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"
        
   'Loop through all tickers to last row
    
    For i = 2 To LastRow
     
    'Check to see if we are still within the same ticker name, if not....
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'Set Ticker name
    Category_Name = ws.Cells(i, 1).Value
    
    'calculate Yearly change total
    
    Yearly_Change_Total = Yearly_Change_Total + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
    
    Total_Inv_Volume = Total_Inv_Volume + ws.Cells(i, 7).Value
    
    'Calculate Percent change
    
    Per_Change = Per_Change + ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value)
    
    'Print ticker name in the summary table
        ws.Range("i" & Summary_Table_Row).Value = Category_Name
    
    'Print Yearly Change total to the summary table
        ws.Range("j" & Summary_Table_Row).Value = Yearly_Change_Total
        
    'Print Total Stock Volume to the summary table
        ws.Range("L" & Summary_Table_Row).Value = Total_Inv_Volume
       
    'Print Percent Change to the summary table
        ws.Range("K" & Summary_Table_Row).Value = Per_Change
    
      
    'Add one to the summary Table
       Summary_Table_Row = Summary_Table_Row + 1
           
    ' Reset the  total
    
    Yearly_Change_Total = 0
    
    Total_Inv_Volume = 0
    
    Per_Change = 0
    
    Else
    
        Yearly_Change_Total = Yearly_Change_Total + (ws.Cells(i, 6).Value - ws.Cells(i, 3).Value)
    
        Total_Inv_Volume = Total_Inv_Volume + ws.Cells(i, 7).Value
        
        Per_Change = Per_Change + ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value)
   
    End If
    
Next i

'Column Formatting, resource for formatting -  from https://www.reddit.com/r/vba/comments/9ksy0f/need_help_with_looping_through_stock_ticker_data/
    
   ws.Columns("K").NumberFormat = "0.00%"
   
Dim MaxValue As Double
Dim maxPercentageIncrease As Double
Dim minPercentage As Double
Dim TickerMaxValue As String
Dim TickerMaxPercentage As String
Dim TickerMinPercentage As String

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LastRow
    If ws.Range("L" & i).Value >= MaxValue Then
    MaxValue = ws.Range("L" & i).Value
    TickerMaxValue = ws.Range("I" & i).Value
    
End If

   If ws.Range("K" & i).Value > maxPercentage Then
    maxPercentage = ws.Range("K" & i).Value
    TickerMaxPercentage = ws.Range("I" & i).Value
End If

    If ws.Range("K" & i).Value < minPercentage Then
    minPercentage = ws.Range("K" & i).Value
    TickerMinPercentage = ws.Range("I" & i).Value
End If
Next i

ws.Cells(4, 17).Value = MaxValue
ws.Cells(4, 16).Value = TickerMaxValue
ws.Cells(3, 17).Value = minPercentage
ws.Cells(3, 16).Value = TickerMinPercentage
ws.Cells(2, 17).Value = maxPercentage
ws.Cells(2, 16).Value = TickerMaxPercentage

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
   
 'Formatting cells, Conditional Formatting with assitance from chat GPT
 
 Dim rngFormat
Set rngFormat = ws.Range("J2:J" & LastRow)

With rngFormat.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    .Interior.Color = RGB(255, 0, 0)
End With
With rngFormat.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    .Interior.Color = RGB(0, 176, 80)
End With

   Next ws
End Sub





Resources below used as guidance to build on my code:

 'Formatting cells, Conditional Formatting with assitance from chat GPT
![image](https://github.com/rosamp1/vba-challenge/assets/132237292/c3e6b1a3-efce-4219-9cfe-9a801b34cd77)

'Column Formatting, resource for formatting -  from https://www.reddit.com/r/vba/comments/9ksy0f/need_help_with_looping_through_stock_ticker_data/    
![image](https://github.com/rosamp1/vba-challenge/assets/132237292/0a57b82d-8205-4d43-802f-2323101e9dbb)

'Greates % Increase, Decrease and Max Value Table completed with assitance from Chat GPT and Youtube Video: https://www.bing.com/videos/search?q=how+do+I+use+vbs+code+to+pull+max+value&&view=detail&mid=95569E28219E929B2B6A95569E28219E929B2B6A&&FORM=VRDGAR&ru=%2Fvideos%2Fsearch%3Fq%3Dhow%2Bdo%2BI%2Buse%2Bvbs%2Bcode%2Bto%2Bpull%2Bmax%2Bvalue%26FORM%3DHDRSC4
![image](https://github.com/rosamp1/vba-challenge/assets/132237292/cf52e921-3176-4c99-929a-d3ea69f7750a)

