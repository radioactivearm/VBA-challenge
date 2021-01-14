Attribute VB_Name = "Module21"
Sub stock_analyzer()
    
    For Each ws In Worksheets  ' what makes it work on all sheets, i had to make all things start with ws. to work
    
        Dim beginning_open As Double
        Dim ending_close As Double
        Dim per_change As Double
        Dim year_end As Double
        
        
        ' adding column titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        
        '-----------------------------------------------------
        ' BONUS STUFF
        'table stuff
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        '-------------------------------------------------------
        
        
        
        beginning_open = ws.Cells(2, 3).Value 'the value of opening days open
        
        Dim off As Integer
        off = 2 'offset counter
        
        Dim stocked As LongLong
        stocked = 0
        
        color_up = 4 ' makes positive years green
        color_down = 3 ' makes negatice years red
        
        
        '------------------------------------------------------------------
        ' BONUS STUFF
        Dim first_tick As String 'holder for ticker for greatest % Increase
        first_tick = ""
        Dim first_per As Double 'holder for % of greatest % increase
        first_per = 0
        
        Dim last_tick As String 'holder for ticker for greatest % decrease
        last_tick = ""
        Dim last_per As Double ' holder for % of greatest % decrease
        last_per = 0
        
        Dim great_tot_tick As String ' holder for greatest total volume ticker
        great_tot_tick = ""
        Dim great_tot As LongLong 'holder for greatest total volume
        great_tot = 0
        '----------------------------------------------------------------------
        
        ' the number of rows in the ticker column
        Dim lastrow As LongLong
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' the number of rows in the ticker column that i made (i know, a little confusing)
        Dim lastrow2 As Long
        lastrow2 = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        
        
        For i = 2 To lastrow ' loop thru all stocks
        
            stocked = stocked + ws.Cells(i, 7).Value ' counter for total stock volume
            
            Dim this_cell As String
            Dim next_cell As String
        
            this_cell = ws.Cells(i, 1).Value 'saves this ticker into variable
            next_cell = ws.Cells(i + 1, 1).Value 'saves next ticker into variable
            
            
            If this_cell <> next_cell Then 'every time there is a new ticker
                
                'calculate end of day values
                ending_close = ws.Cells(i, 6).Value
                
                'calculate end of year values
                year_end = ending_close - beginning_open
                
                If beginning_open = 0 Then 'To fix a problem where there were all zeros. you cannot divide by zero.
                    per_change = 0
                Else
                    'calculate percent change
                    per_change = year_end / beginning_open
                End If
                
                
                
                ' ----------------------------
                ' BONUS STUFF
                If per_change > first_per Then 'if the percent change is greater then the highest so far it will store data
                    first_per = per_change
                    first_tick = this_cell
                    
                ElseIf per_change < last_per Then 'if percent change is lowest so fat it will store data
                    last_per = per_change
                    last_tick = this_cell
                End If
                
                If stocked > great_tot Then ' if the total stock is highest so far it will store data
                    great_tot = stocked
                    great_tot_tick = this_cell
                End If
                '-------------------------------
                
                
        
                ws.Range("I" & off) = this_cell 'adding the ticker marks to table
                ws.Range("J" & off) = year_end 'adding yearly change and offseting one below last tickers
                ws.Range("K" & off) = per_change 'adding per_change to table
                
                ' adding in code to make percent change a percentage format
                ws.Range("K" & off).NumberFormat = "0.00%"
                
                ws.Range("L" & off) = stocked 'adding total stock
                
                ' Adding Fill Formatting to Yearly Change column
                If year_end >= 0 Then
                    ws.Range("J" & off).Interior.ColorIndex = color_up 'if yearly change is equal to or greater than 0 the color is set to green
                Else
                    ws.Range("J" & off).Interior.ColorIndex = color_down ' if yearly change is less than zero the color is set to red
                End If
                
                
                
                off = off + 1 'increase offset by one
                
                
                stocked = 0 'reset stock to zero for next ticket
                
                'do next beginning of year calculations
                beginning_open = ws.Cells(i + 1, 3).Value
        
                
            End If
            
            
        Next i
        
        ' its not working.
        ' going to move into for loop and do on a cellular level step by step.
        'ws.Range("K2:K" & lastrow2).NumberFormat = "0.00%"   'change percent change column to percentages
        
        'so when i run this once it only formats first cell then none of the others. but if i run it again it formats the rest
        ' i am going to try to just format it agian in code to see if i can get it to work with out running whole program twice
        'ws.Range("K2:K" & lastrow2).NumberFormat = "0.00%"
        ' it worked! yay!
        
        
        '----------------------------------------------------------------------
        ' BONUS STUFF
        ws.Range("P2").Value = first_tick 'ticker for highest percentage
        ws.Range("Q2").Value = first_per 'percentage for highest percentage
        ws.Range("Q2").NumberFormat = "0.00%" 'Changing format to percentage
        
        ws.Range("P3").Value = last_tick 'ticker for lowest percentage
        ws.Range("Q3").Value = last_per 'percentage for lowest percentage
        ws.Range("Q3").NumberFormat = "0.00%"
        
        ws.Range("P4").Value = great_tot_tick 'ticker for greatest total stock volume
        ws.Range("Q4").Value = great_tot 'greatest total stock volume
        '------------------------------------------------------------------------------
        
        ws.Columns("I:Q").AutoFit 'adjust size of cells automatically so it is easy to read.
        
    Next ws
    
    MsgBox ("COMPLETE") ' just so that i know it is done
    

End Sub

