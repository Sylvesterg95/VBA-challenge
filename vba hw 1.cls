VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickerloop()

 ' worksheet stock
 
Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
    
        ' find the last row of the table
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).row

        ' add headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quaterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
      
        Dim open_price As String
        Dim close_price As Double
        Dim quaterly_change As Double
        Dim ticker As String
        Dim percent_change As Double
        
        Dim volume As String
        Dim row As Double
        Dim column As Integer
        
        volume = 0
        row = 2
        column = 1
       
       
       
       
        
        ' this is the initial price
        open_price = Cells(2, column + 2).Value
        
          ' check for  any mismatch
        For i = 2 To last_row
        
         
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            
                ' ticker name
                ticker = Cells(i, column).Value
                Cells(row, column + 8).Value = ticker
                
                '  close price
                close_price = Cells(i, column + 5).Value
                
                '  quaterly change
                quaterly_change = close_price - open_price
                Cells(row, column + 8).Value = quaterly_change
               
               '  percent change
                    percent_change = quaterly_change / open_price
                    Cells(row, column + 11).Value = percent_change
                    Cells(row, column + 10).NumberFormat = "0.00%"
               
               
               '  total volume per quater
                volume = volume + Cells(i, column + 6).Value
                Cells(row, column + 11).Value = volume
                
                ' iterate to the next row
                row = row + 1
                
                ' reset open price to next ticker
                open_price = Cells(i + 1, column + 2)
                
                ' reset volume for next ticker
                volume = 0
                
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i
        
        
        ' find the last row of ticker column
        quaterly_change_last_row = ws.Cells(Rows.Count, 9).End(xlUp).row
        
        ' set the Cell Colors
        For j = 2 To quaterly_change_last_row
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
       
        
        ' find the highest value of each ticker
        For k = 2 To quaterly_change_last_row
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & quaterly_change_last_row)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k
        
        ActiveSheet.Range("I:Q").Font.Bold = True
        ActiveSheet.Range("I:Q").EntireColumn.AutoFit
        Worksheets("Q1").Select
        
    Next ws

End Sub

