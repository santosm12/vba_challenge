VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickeranalysis()

        
        Dim tickername As String
    
        Dim tickervolume As Double
        tickervolume = 0

        Dim summary_ticker_row As Integer
        summary_ticker_row = 2
        

        Dim open_price As Double
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 18).Value = "Ticker"
        Cells(1, 19).Value = "Value"
        Cells(2, 17).Value = "Greatest % increase"
        Cells(3, 17).Value = "Greatest % decrease"
        Cells(4, 17).Value = "Greatest total volume"
        

        lastrow = Cells(Rows.Count, 1).End(xlUp).Row


        For i = 2 To lastrow

       
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
              tickername = Cells(i, 1).Value

              tickervolume = tickervolume + Cells(i, 7).Value

              Range("I" & summary_ticker_row).Value = tickername

              Range("L" & summary_ticker_row).Value = tickervolume

              close_price = Cells(i, 6).Value

              yearly_change = (close_price - open_price)

              Range("J" & summary_ticker_row).Value = yearly_change


                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
                End If

              Range("K" & summary_ticker_row).Value = percent_change
              Range("K" & summary_ticker_row).NumberFormat = "0.00%"
   

              summary_ticker_row = summary_ticker_row + 1

              tickervolume = 0


              open_price = Cells(i + 1, 3)
            
            Else
              
              tickervolume = tickervolume + Cells(i, 7).Value

            
            End If
        
        Next i


    lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row
    
    
    For i = 2 To lastrow_summary
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 10
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i
    
    For i = 2 To lastrow_summary
    
        If Cells(i, 11).Value = WorksheetFunction.Max(Range("k2:k" & lastrow_summary)) Then
            Cells(2, 18).Value = Cells(i, 9).Value
            Cells(2, 19).Value = Cells(i, 11).Value
            Cells(3, 19).NumberFormat = "0.00%"
            
        ElseIf Cells(i, 11).Value = WorksheetFunction.Min(Range("k2:k" & lastrow_summary)) Then
            Cells(3, 18).Value = Cells(i, 9).Value
            Cells(3, 19).Value = Cells(i, 11).Value
            Cells(3, 19).NumberFormat = "0.00%"
            
        ElseIf Cells(i, 12).Value = WorksheetFunction.Max(Range("l2:l" & lastrow_summary)) Then
            Cells(4, 18).Value = Cells(i, 9).Value
            Cells(4, 19).Value = Cells(i, 12).Value
            Cells(4, 19).NumberFormat = "0.00"
            
            
        End If
        
    Next i
    
        
        
End Sub
