Attribute VB_Name = "Module1"
Sub Stock_Analysis():

Dim ws As Worksheet

'Loop through all the worksheets

    For Each ws In Worksheets
    
    
    ' Creating Variables and assigning values

Dim TickerSymbol As String

Dim TickerTotal As Double
TickerTotal = 0

Dim SummaryTable As Integer
SummaryTable = 2


  'Creating Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

        
'  counting rows in a sheet
          Dim RowCount As Long
          RowCount = ws.Range("A1").End(xlDown).Row
          
          
' Looping through all ticker in ws
For i = 2 To RowCount
       
         
         Dim tickervalue1 As String
         tickervalue1 = ws.Cells(i + 1, 1).Value
         
         Dim tickervalue2 As String
         tickervalue2 = ws.Cells(i, 1).Value
         
         Dim FirstOpenPrice As Double
         Dim LastClosePrice As Double
        
         
         
         If i = 2 Then
            FirstOpenPrice = ws.Cells(i, 3).Value
            
         End If
            
If tickervalue1 <> tickervalue2 Then
         
                    TickerSymbol = ws.Cells(i, 1).Value
                    TickerTotal = TickerTotal + ws.Cells(i, 7).Value
                    LastClosePrice = ws.Cells(i, 6).Value
         
                   ws.Range("I" & SummaryTable).Value = TickerSymbol

                    ws.Range("J" & SummaryTable).Value = LastClosePrice - FirstOpenPrice

                    ws.Range("J" & SummaryTable).NumberFormat = "#,##0.00"
                    
                    If ws.Range("J" & SummaryTable).Value > 0 Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4
                    
                    
                    ElseIf ws.Range("J" & SummaryTable).Value < 0 Then
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3
                
                    
                    End If
                    
                    
                    
                    
                    
             If (FirstOpenPrice <> 0) Then
                ws.Range("K" & SummaryTable).Value = ((LastClosePrice - FirstOpenPrice) / FirstOpenPrice)
                
                 ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
         
             End If
         
         'setting variables and assigning values again
         
    
                 ws.Range("L" & SummaryTable).Value = TickerTotal
        
                 SummaryTable = SummaryTable + 1
              
                 TickerTotal = 0
         
                FirstOpenPrice = ws.Cells(i + 1, 3)

            Else
        
              TickerTotal = TickerTotal + ws.Cells(i, 7).Value
        
End If
            

    Next i
    
    ' Finding the Greatest % for increase and decrease and

X = GreatestPercentIncrease = ws.Cells(2, 11).Value
Y = GreatestPercentDecrease = ws.Cells(2, 11).Value
Z = GreatestTotalStockVolume = ws.Cells(2, 12).Value

A = GreatestIncreaseTicker = ws.Cells(2, 9).Value
B = GreatestDecreaseTicker = ws.Cells(2, 9).Value
C = GreatestStockVolumeTicker = ws.Cells(2, 9).Value




Lastrow = ws.Cells(Rows.Count, "I").End(xlUp).Row


For i = 2 To Lastrow

        If ws.Cells(i, 11).Value > X Then
            X = ws.Cells(i, 11).Value
            A = ws.Cells(i, 9).Value

        End If
        

            If ws.Cells(i, 11).Value < Y Then
            Y = ws.Cells(i, 11).Value
            B = ws.Cells(i, 9).Value



            End If

                If ws.Cells(i, 12).Value > Z Then
                Z = ws.Cells(i, 12).Value
                C = ws.Cells(i, 9).Value
                
                
                End If


    ws.Range("P2").Value = Format(A, "Percent")
    ws.Range("P3").Value = Format(B, "Percent")
     ws.Range("P4").Value = C
     
     
    ws.Range("Q2").Value = Format(X, "Percent")
    ws.Range("Q3").Value = Format(Y, "Percent")
   
    ws.Range("Q4").Value = Z
    
                
                
Next i

Next ws


End Sub

