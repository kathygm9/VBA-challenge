Attribute VB_Name = "Stock_Analysis"
Sub StockAnalysis()
 
    Dim CurrentWs As Worksheet
    Dim LastRow As Long
    Dim i As Long
    Dim Summary_Table_Row As Long
    Dim Ticker_Name As String
    Dim Open_Price As Double, Close_Price As Double
    Dim Delta_Price As Double, Delta_Percent As Double
    Dim Total_Ticker_Volume As Double
    
    ' Loop through all worksheets in the workbook
    For Each CurrentWs In Worksheets
        
        ' Initialize the row for writing summary table headers
        CurrentWs.Range("I1").Value = "Ticker"
        CurrentWs.Range("J1").Value = "Yearly Change"
        CurrentWs.Range("K1").Value = "Percent Change"
        CurrentWs.Range("L1").Value = "Total Stock Volume"
        
        Summary_Table_Row = 2
        LastRow = CurrentWs.Cells(CurrentWs.Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables for the first stock
        If LastRow >= 2 Then
            Open_Price = CurrentWs.Cells(2, 3).Value
        End If
        
        Total_Ticker_Volume = 0
        
        ' Loop through all rows in the worksheet
        For i = 2 To LastRow
            
            Total_Ticker_Volume = Total_Ticker_Volume + CurrentWs.Cells(i, 7).Value
            
            ' Check if this is the last row for the current ticker
            If i = LastRow Or CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                Close_Price = CurrentWs.Cells(i, 6).Value
                
                ' Calculate changes
                Delta_Price = Close_Price - Open_Price
                If Open_Price <> 0 Then
                    Delta_Percent = (Delta_Price / Open_Price) * 100
                Else
                    Delta_Percent = 0  ' Avoid division by zero
                End If
                
                ' Output the data to the summary table
                CurrentWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
                CurrentWs.Range("J" & Summary_Table_Row).Value = Delta_Price
                CurrentWs.Range("K" & Summary_Table_Row).Value = Format(Delta_Percent, "0.00") & "%"
                CurrentWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Format the "Yearly Change" cell based on positive or negative change
                If Delta_Price > 0 Then
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.Color = RGB(57, 255, 20) ' Green
                Else
                    CurrentWs.Range("J" & Summary_Table_Row).Interior.Color = RGB(238, 75, 43) ' Red
                End If
                
                ' Prepare for next ticker
                Summary_Table_Row = Summary_Table_Row + 1
                If i < LastRow Then
                    Open_Price = CurrentWs.Cells(i + 1, 3).Value
                End If
                
                Total_Ticker_Volume = 0  ' Reset volume for next ticker
            End If
            
        Next i
        
    Next CurrentWs

End Sub


