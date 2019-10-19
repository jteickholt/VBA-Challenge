Sub vbachallenge2()


'Variable declaration

Dim ticker As String
Dim startprice As Double
Dim endprice As Double
Dim totvolume As Double
Dim lastrow As Long
Dim yrchange As Double
Dim pctchange As Double
Dim firstmonth As Integer
Dim lastmonth As Integer
Dim tickcount As Long
Dim row As Long
Dim colletter As String
Dim maxtick As String
Dim mintick As String
Dim voltick As String
Dim maxinc As Single
Dim maxdec As Single
Dim incamount As Single
Dim decamount As Single
Dim largestvol As Double


'To alter the code from VBA Challenge 1, I had to encapsulate the code in in a loop across each worksheet
'I then had to update any cell or range refers with the prefix "ws." so it would work on each sheet



For Each ws In Worksheets
  ws.Activate
  


'First clear the columns where the code is putting data, so no junk left during testing
'Define the last row and columne of data and assign initial values to variables
'Sort the data to make sure in correct order.

ws.Range("I:Z").Clear

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row
lastcol = ws.Cells(1, Columns.Count).End(xlToLeft).Column


ws.Range("a1:G" & lastrow).Sort Key1:=ws.Range("A1:A" & lastrow), order1:=xlAscending, Header:=xlYes, _
  Key2:=ws.Range("B1:B" & lastrow), order1:=xlAscending, Header:=xlYes
  

'Assign some variables

startprice = 0
endprice = 0
totvolume = 0
yrchange = 0
pctchange = 0
firstmonth = 1
lastmonth = 0
tickcount = 0
row = 0
maxinc = 0
maxdec = 0
inctick = ""
dectick = ""
incamount = 0
decamount = 0
largestvol = 0
ticker = ""


'Assign column headings to summary output area

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"



'The loop below is used to get the data for the first part of the challenge and write it out
'It Loop through each row to obtain the starting price and ending price for each ticker
'The change and percent change is then calculated for each ticker
'Also added up the total volume as it goes, as well as keeping track of the number of tickers


For i = 2 To lastrow

'If it is the first time a ticker is seen it will set some inital values for the ticker

   If firstmonth = 1 Then
      ticker = ws.Cells(i, 1).Value
      startprice = ws.Cells(i, 3).Value
      totvolume = 0
      tickcount = tickcount + 1
   End If
   
   firstmonth = 0
   totvolume = totvolume + ws.Cells(i, 7).Value
 
   
   If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
     endprice = ws.Cells(i, 6).Value
     yrchange = endprice - startprice
     
     '  There appears to be some bad data that shows all prices as zeroes
     '  So added the check below to avoid division by zero error
     
     
     If startprice = 0 Then
       pctchange = 0
     Else
       pctchange = yrchange / startprice
     End If
     
     row = row + 1
 
 'Once the needed data is calculated for each ticker, it is written to the sheet
     
     ws.Cells(row + 1, 9).Value = ticker
     ws.Cells(row + 1, 10).Value = yrchange
     ws.Cells(row + 1, 11).Value = pctchange
     ws.Cells(row + 1, 12).Value = totvolume
     
    firstmonth = 1
     
   End If
   
Next i

'Clean up some formatting

Range("J:J").NumberFormat = "0.000000000"
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("I:L").Columns.AutoFit

'The next section is used to grab the data needed in the second step of the challenge
'Here it loops through the tick level data created in step 1 to find the biggest increase and decrease,
'It also find the largest volume.
'The values are then written out to the sheeet and formatting is cleaned up


For i = 2 To (tickcount + 1)

  If ws.Cells(i, 10) > 0 Then
     ws.Cells(i, 10).Interior.ColorIndex = 4
  Else
    ws.Cells(i, 10).Interior.ColorIndex = 3
  End If
  If i = 1 Then
    maxinc = ws.Cells(i, 11).Value
    maxdec = ws.Cells(i, 11).Value
    inctick = ws.Cells(i, 9).Value
    dectick = ws.Cells(i, 9).Value
    incamount = ws.Cells(i, 10).Value
    decamount = ws.Cells(i, 10).Value
    largestvol = ws.Cells(i, 12).Value
  End If
  If ws.Cells(i, 11).Value > maxinc Then
    maxinc = ws.Cells(i, 11).Value
    inctick = ws.Cells(i, 9).Value
    incamount = ws.Cells(i, 10).Value
  End If
  If ws.Cells(i, 11).Value < maxdec Then
      maxdec = ws.Cells(i, 11).Value
      dectick = ws.Cells(i, 9).Value
      decamount = ws.Cells(i, 10).Value
  End If
  If ws.Cells(i, 12).Value > largestvol Then
     largestvol = ws.Cells(i, 12).Value
     voltick = ws.Cells(i, 9)
  End If
  
Next i

'Label the rows

  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greater % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  
'Write out the final values
  
  
  ws.Range("P1").Value = "Ticker"
  ws.Range("P2").Value = inctick
  ws.Range("P3").Value = dectick
  ws.Range("P4").Value = voltick
  
  
  ws.Range("Q1").Value = "Value"
  ws.Range("Q2").Value = maxinc
  ws.Range("Q3").Value = maxdec
  ws.Range("Q4").Value = largestvol

  ws.Range("Q2:Q3").NumberFormat = "0.00%"
  ws.Range("O:O").Columns.AutoFit
  ws.Range("Q:Q").Columns.AutoFit

  ' The following formatting was just added to make some room on the sheet so
  'everything could display on the screenshot in the proper format
  
  ws.Range("H:H").ColumnWidth = 5
  ws.Range("M:M").ColumnWidth = 5
  ws.Range("N:N").ColumnWidth = 5
  ws.Range("A:G").ColumnWidth = 9
  

Next ws


End Sub
