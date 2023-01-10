
Sub Stocks()

   For Each ws In Worksheets
    
        Dim Tsymbol As String 'Ticker symbol for all respective ws.Cells
        
        
        Dim TSV As Double ' Total stock volume variable
        TSV = 0
        Dim O_P As Double ' Opening Price variable
        O_P = 0
        Dim C_P As Double ' Closing price variable
        C_P = 0
        Dim Y_C As Double   'Yearly change variable
        Y_C = 0
        Dim Per_C As Double 'Percentage change variable
        Per_C = 0
        'max variabliables for greatest values of data bonus section
               
        Dim maxtv As Double
        maxtv = 0
        Dim maxd As Double
        maxd = 0
        Dim maxi As Double
        maxi = 0
        
        
       ' Variable to remember location of each Ticker symbol starting Below label
        Dim Rmem As Integer
        Rmem = 2
    
        ' Determine the Last Row of column
    
        lastrow = ws.Cells(rows.count, 1).End(xlUp).Row
        
              
              
        ' Label 1st Cell on each column based on data under it
            ws.Cells(1, 9).value = "Ticker"
            ws.Cells(1, 10).value = "Yearly change"
            ws.Cells(1, 11).value = "Percentage Change"
            ws.Cells(1, 12).value = "Total Stock Volume"
            ws.Cells(1, 17).value = "Value"
            
        ' Labels associated with greatest calculated values section
            ws.Cells(2, 15).value = "Greatest % Increase"
            ws.Cells(3, 15).value = "Greatest % Decrease"
            ws.Cells(4, 15).value = "Greatest Total Volume"
            ws.Cells(1, 16).value = "Ticker"
            
        ' keep track of most recent ticker start
            Dim T_S As Double
            
            T_S = 2
            
            For i = 2 To lastrow
            ' If not on same Ticker symbol then
                If ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value Then
                
         
            
            ' Set & Print Ticker Symbol on respective cell within summary
                Tsymbol = ws.Cells(i, 1).value
                
                ws.Range("I" & Rmem).value = Tsymbol
                
                
           ' Set & Print Yearly change value on respective cell within summary
                          
                
                C_P = ws.Cells(i, 6).value
                
                O_P = ws.Cells(T_S, 3).value
               
                Y_C = C_P - O_P
                                               
                ws.Range("J" & Rmem).value = Y_C
                
        ' Set & Print Percentage Change value on respective cell within summary
        
                Per_C = Y_C / O_P
            
                ws.Range("K" & Rmem).value = Per_C
                
                ws.Range("K" & Rmem).numberformat = "0.00%"
                
           ' Set & Print Total Stock Volume value on respective cell within summary
                                              
                TSV = TSV + ws.Cells(i, 7).value
                
                ws.Range("L" & Rmem).value = TSV
                                
            ' Add to the summany table row we're remembering
            
                Rmem = Rmem + 1
                
            ' Add one to keep track of beginning of different ticker after end of same
                
                T_S = i + 1
            ' Reset respective values before jumping to different ticker
                O_P = 0
                
                C_P = 0
                                
                Y_C = 0
                
                Per_C = 0
                
                TSV = 0
                
                
            ' If the cell immediately following a row is the same ticker
                
                Else
                
               'Keep shifting from ticker to ticker on closing price till end of same ticker
                
                C_P = ws.Cells(i, 6).value
                
                ' Add to total stock volume within same ticker
                TSV = TSV + ws.Cells(i, 7).value
                
                               
             End If
             
                
                
             
            Next i
       
           
            For j = 2 To Rmem - 1
            
                    ' Conditional format to change ws.Cells color on Column 10 based on negative or positive values
            
                    If ws.Cells(j, 10) < 0 Then
                    
                    ' Set the Cell Colors to Red
                      ws.Cells(j, 10).Interior.ColorIndex = 3
                    Else
                      ' Set the Cell Colors to Green
                      ws.Cells(j, 10).Interior.ColorIndex = 4
                                
                    End If
                    
                    ' Populate greatest increase with values from summary table
                    
                    If ws.Cells(j, 11).value >= maxi Then
                    
                        maxi = ws.Cells(j, 11).value
                        
                        ws.Cells(2, 17).value = maxi
                        
                        ws.Cells(2, 17).numberformat = "0.00%"
                        
                        Tsymbol = ws.Cells(j, 11).Offset(0, -2)
                        
                        ws.Cells(2, 16).value = Tsymbol
                        
                    End If
                    
                    ' Populate greatest decrease with values from summary table
                    
                    If ws.Cells(j, 11).value <= maxd Then
                    
                        maxd = ws.Cells(j, 11).value
                        
                        ws.Cells(3, 17).value = maxd
                        
                        ws.Cells(3, 17).numberformat = "0.00%"
                        
                        Tsymbol = ws.Cells(j, 9).value
                        
                        ws.Cells(3, 16).value = Tsymbol
                        
                    End If
                    
                    ' Populate Greatest total Volume with values from summary table
                    
                    If ws.Cells(j, 12).value >= maxtv Then
                    
                        maxtv = ws.Cells(j, 12).value
                        
                        ws.Cells(4, 17).value = maxtv
                        
                        Tsymbol = ws.Cells(j, 12).Offset(0, -3)
                        
                        ws.Cells(4, 16).value = Tsymbol
                        
                    End If
                                
            
            Next j
            
            ' Autofit columns
           
     ws.Columns.AutoFit
    
    

   Next ws
End Sub



