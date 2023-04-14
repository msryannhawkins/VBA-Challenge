VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockDataAll()
       
Dim j As Integer
Dim ws_num As Integer
Dim ws As Worksheet
ws_num = 3
For j = 1 To ws_num

'find the sheets that needs to take action

If Sheets(j).Name = "2018" Or _
Sheets(j).Name = "2019" Or _
Sheets(j).Name = "2020" Then
ThisWorkbook.Worksheets(j).Activate


                'Define all variables
                    'Ticker Name
                        Dim TN As String
                     'Stock Close
                        Dim SC As Double
                    'Stock Open
                        Dim SO As Double
                        SO = Cells(2, 3).Value
                    'Yearly Change
                        Dim YC As Double
                    'Percent Change
                        Dim PC As Double
                    'Total Stock Volume
                        Dim TSV As Double
                        TSV = 0
                   
                
                    'For Loop Variable
                        Dim i As Double
                    'Row two for first value input
                        Dim SummaryTableRow As Integer
                        SummaryTableRow = 2
                    'End limit for For Loop
                        Dim LastRow As Double
                        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
                
                  
                'Establish Column Names
                    Cells(1, 9).Value = "Ticker"
                    Cells(1, 10).Value = "Yearly Change"
                    Cells(1, 11).Value = "Percent Change"
                    Cells(1, 12).Value = "Total Stock Volume"
                    
                    'Summary Table
                        Cells(2, 14).Value = "Greatest % Increase"
                        Cells(3, 14).Value = "Greatest % Decrease"
                        Cells(4, 14).Value = "Greatest Total Volume"
                        Cells(1, 15).Value = "Ticker"
                        Cells(1, 16).Value = "Value"
                        
                'Adjust Cell width to fit text
                Columns("K:L").ColumnWidth = 15
                Columns("N").ColumnWidth = 20
                
                'Begin For Loop: declare the if statement after setting the limit; assign variables within the if statment
                    For i = 2 To LastRow
                    
                        'If the cells are within  the same ticker name, then...
                        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                                     
                            'Assign all variables
                                'Ticker Name: Column A/1, to LastRow
                                    TN = Cells(i, 1).Value
                                    
                                'Add Total Stock Volume
                                    TSV = TSV + Cells(i, 7).Value
                                    
                            'Execute Calculation and input into Summary Table
                                'Ticker Name input into Cells(1, 9)
                                    Range("I" & SummaryTableRow).Value = TN
                                
                                'Total Stock Volume input into Cells
                                    Range("L" & SummaryTableRow).Value = TSV
                            
                            'Assign remainder of variables
                                'Stock Close: Column F/6, to LastRow
                                    SC = Cells(i, 6).Value
                                    
                            'Execute Calculation and input into Summary Table
                                'Yearly Change = Stock Close - Stock Open
                                    YC = SC - SO
                                    
                                    Range("J" & SummaryTableRow).Value = YC
                                
                                                'Correct Division by Zero error
                                                If (SO = 0) Then
                                
                                                   PC = 0
                                
                                                Else
                                                    
                                                    PC = YC / SO
                                                
                                                End If
                               
                               'Input Percent Change into Summary Table
                                    Range("K" & SummaryTableRow).Value = PC
                                    Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                               
                             'Add one to the Summary Table
                                    SummaryTableRow = SummaryTableRow + 1
                            
                            'Reset the TSV
                                    TSV = 0
                                    
                            'Reset Stock Open
                                    SO = Cells(i + 1, 3).Value
                            
                    'If the cell immediately following a row is the same ticker...
                    Else
                
                        ' Add to the TSV
                                TSV = TSV + Cells(i, 7).Value
                            
                        
                       End If
                               
                    Next i
                  
                '--------------------------------------------------------------------------------
                'Create for loop to color code cells based on positivity or negativity
                
                For i = 2 To LastRow
                    
                    If Range("K" & i).Value >= 0 Then
                        Range("K" & i).Interior.ColorIndex = 4
                    
                    Else
                        Range("K" & i).Interior.ColorIndex = 3
                        
                    End If
                
                Next i
'-------------------------------------------------------------------------------
'
       Dim Greatest As Double
       Dim Least As Double
       Dim GrVol As Double
       Dim TickerSum As String
       
       Greatest = 0
       Least = 0
       GrVol = 0
       
       For i = 2 To 3001
       
       If Range("K" & i).Value > Greatest Then
        Greatest = Range("K" & i).Value
        TickerSum = Range("I" & i).Value
        Cells(2, 15).Value = TickerSum
        Cells(2, 16).Value = Greatest
        End If
        
        If Range("K" & i).Value < Least Then
        Least = Range("K" & i).Value
        TickerSum = Range("I" & i).Value
        Cells(3, 15).Value = TickerSum
        Cells(3, 16).Value = Least
        End If
        
        If Range("L" & i).Value > GrVol Then
        GrVol = Range("L" & i).Value
        TickerSum = Range("I" & i).Value
        Cells(4, 15).Value = TickerSum
        Cells(4, 16).Value = GrVol
        End If
       
       Next i
       
End If

Next j

End Sub

'-------------------------------------------------------------------------------




