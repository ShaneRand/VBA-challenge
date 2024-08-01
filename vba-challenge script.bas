Attribute VB_Name = "Module1"
Option Explicit
 
Sub RunStockAnalysis()
    ' This is the main procedure that will run the entire analysis
 
    ' Call the function to set up headers
    SetupHeaders
 
    ' Call the function to perform the analysis
    PerformAnalysis
 
    MsgBox "Stock analysis complete!", vbInformation
End Sub
 
Private Sub SetupHeaders()
    ' Define variables
    Dim qtab As Worksheet
 
    ' Add Summary Table Headers to Each qtab
    For Each qtab In ThisWorkbook.Worksheets
        With qtab
            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Quarterly Price Change"
            .Range("K1").Value = "Percentage Change"
            .Range("L1").Value = "Total Stock Volume"
            .Range("P1").Value = "Ticker"
            .Range("Q1").Value = "Value"
            .Range("O2").Value = "Greatest % Increase"
            .Range("O3").Value = "Greatest % Decrease"
            .Range("O4").Value = "Greatest Total Volume"
               
        End With
    Next qtab
End Sub
 
Private Sub PerformAnalysis()
    ' Define all variables
    Dim qtab As Worksheet
    Dim ticker_name As String
    Dim ticker_total As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim i As Long
    Dim LastRow As Long
    Dim summary_table_row As Integer
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total As Double
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Total_Ticker As String
    
    'set initial variable for keeping track of the location of different date's open price
      Dim Price_Row As Long
      Price_Row = 2
    
 
    ' Loop through each worksheet
    For Each qtab In ThisWorkbook.Worksheets
        ' Find last row
        LastRow = qtab.Cells(qtab.Rows.Count, 1).End(xlUp).Row
 
        ' Initialize variables
        summary_table_row = 2
        ticker_total = 0
        open_price = qtab.Cells(2, 3).Value
 
        ' Go through daily trade data
        For i = 2 To LastRow
            ' If next row is different ticker, process current ticker
            If qtab.Cells(i + 1, 1).Value <> qtab.Cells(i, 1).Value Or i = LastRow Then
                ' Set the ticker name
                ticker_name = qtab.Cells(i, 1).Value
 
                ' Add final volume for this ticker
                ticker_total = ticker_total + qtab.Cells(i, 7).Value
 
                ' Get closing price
                close_price = qtab.Cells(i, 6).Value
 
                ' Calculate changes
                quarterly_change = close_price - open_price
                If open_price <> 0 Then
                    percent_change = (quarterly_change / open_price)
                Else
                    percent_change = 0
                End If
 
                ' Fill out summary table
                With qtab
                    .Cells(summary_table_row, 9).Value = ticker_name
                    .Cells(summary_table_row, 10).Value = quarterly_change
                    .Cells(summary_table_row, 11).Value = percent_change
                    .Cells(summary_table_row, 12).Value = ticker_total
 
                    ' Format percent change as percentage
                    .Cells(summary_table_row, 11).NumberFormat = "0.00%"
 
                    ' Color coding
                    If quarterly_change > 0 Then
                        .Cells(summary_table_row, 10).Interior.ColorIndex = 4 ' Green
                    ElseIf quarterly_change < 0 Then
                        .Cells(summary_table_row, 10).Interior.ColorIndex = 3 ' Red
                    Else
                        .Cells(summary_table_row, 10).Interior.ColorIndex = 2 ' White
                    End If
                End With
 
                ' Move to next summary table row
                summary_table_row = summary_table_row + 1
 
                ' Reset variables for next ticker
                ticker_total = 0
                open_price = qtab.Cells(i + 1, 3).Value
            Else
                ' Add to ticker total
                ticker_total = ticker_total + qtab.Cells(i, 7).Value
            End If
            
            
        Next i
        
        'set the first ticker's percent change and total stock volume as the greatest ones
        Greatest_Increase = qtab.Range("K2").Value
        Greatest_Decrease = qtab.Range("K2").Value
        Greatest_Total = qtab.Range("L2").Value
        
        'Define last row of Ticker column
        LastRow = qtab.Cells(Rows.Count, "I").End(xlUp).Row
        
        'Loop through each row of Ticker column to find the greatest results
         For i = 2 To LastRow:
               If qtab.Range("K" & i + 1).Value > Greatest_Increase Then
                  Greatest_Increase = qtab.Range("K" & i + 1).Value
                  Greatest_Increase_Ticker = qtab.Range("I" & i + 1).Value
               ElseIf qtab.Range("K" & i + 1).Value < Greatest_Decrease Then
                  Greatest_Decrease = qtab.Range("K" & i + 1).Value
                  Greatest_Decrease_Ticker = qtab.Range("I" & i + 1).Value
                ElseIf qtab.Range("L" & i + 1).Value > Greatest_Total Then
                  Greatest_Total = qtab.Range("L" & i + 1).Value
                  Greatest_Total_Ticker = qtab.Range("I" & i + 1).Value
                End If
            Next i
            
            'Print greatest % increase, greatest % decrease, greatest total volume and their ticker names
            qtab.Range("P2").Value = Greatest_Increase_Ticker
            qtab.Range("P3").Value = Greatest_Decrease_Ticker
            qtab.Range("P4").Value = Greatest_Total_Ticker
            qtab.Range("Q2").Value = Greatest_Increase
            qtab.Range("Q3").Value = Greatest_Decrease
            qtab.Range("Q4").Value = Greatest_Total
            qtab.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next qtab
    
  
        
   
    
        
    
    
End Sub


