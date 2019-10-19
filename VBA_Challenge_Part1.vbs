Sub vbachallenge1()

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



'First clear the columns where the code is putting data, so no junk left during testing
'Define the last row and columne of data and assign initial values to variables
'Sort the data to make sure in correct order.

Range("I:Z").Clear

lastrow = Cells(Rows.Count, 1).End(xlUp).row
lastcol = Cells(1, Columns.Count).End(xlToLeft).Column


Range("a1:G" & lastrow).Sort Key1:=Range("A1:A" & lastrow), order1:=xlAscending, Header:=xlYes, _
  Key2:=Range("B1:B" & lastrow), order1:=xlAscending, Header:=xlYes
  

'Declare some variables


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

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"



'The loop below is used to get the data for the first part of the challenge and write it out
'It Loop through each row to obtain the starting price and ending price for each ticker
'The change and percent change is then calculated for each ticker
'Also added up the total volume as it goes, as well as keeping track of the number of tickers


For i = 2 To lastrow

'If it is the first time a ticker is seen it will set some inital values for the ticker

   If firstmonth = 1 Then
      ticker = Cells(i, 1).Value
      startprice = Cells(i, 3).Value
      totvolume = 0
      tickcount = tickcount + 1
   End If
   
   firstmonth = 0
   totvolume = totvolume + Cells(i, 7).Value
 
   
   If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
     endprice = Cells(i, 6).Value
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
     
     Cells(row + 1, 9).Value = ticker
     Cells(row + 1, 10).Value = yrchange
     Cells(row + 1, 11).Value = pctchange
     Cells(row + 1, 12).Value = totvolume
     
    firstmonth = 1
     
   End If
   
Next i

'Clean up some formatting

Range("J:J").NumberFormat = "0.000000000"
Range("K:K").NumberFormat = "0.00%"
Range("I:L").Columns.AutoFit

'The next section is used to grab the data needed in the second step of the challenge
'Here it loops through the tick level data created in step 1 to find the biggest increase and decrease,
'It also find the largest volume.
'The values are then written out to the sheeet and formatting is cleaned up


For i = 2 To (tickcount + 1)

  If Cells(i, 10) > 0 Then
     Cells(i, 10).Interior.ColorIndex = 4
  Else
  Cells(i, 10).Interior.ColorIndex = 3
  End If
  If i = 1 Then
    maxinc = Cells(i, 11).Value
    maxdec = Cells(i, 11).Value
    inctick = Cells(i, 9).Value
    dectick = Cells(i, 9).Value
    incamount = Cells(i, 10).Value
    decamount = Cells(i, 10).Value
    largestvol = Cells(i, 12).Value
  End If
  If Cells(i, 11).Value > maxinc Then
    maxinc = Cells(i, 11).Value
    inctick = Cells(i, 9).Value
    incamount = Cells(i, 10).Value
  End If
  If Cells(i, 11).Value < maxdec Then
      maxdec = Cells(i, 11).Value
      dectick = Cells(i, 9).Value
      decamount = Cells(i, 10).Value
  End If
  If Cells(i, 12).Value > largestvol Then
     largestvol = Cells(i, 12).Value
     voltick = Cells(i, 9)
  End If
  
Next i

'Label the rows

  Range("O2").Value = "Greatest % Increase"
  Range("O3").Value = "Greater % Decrease"
  Range("O4").Value = "Greatest Total Volume"
  
'Write out the final values
  
  
  Range("P1").Value = "Ticker"
  Range("P2").Value = inctick
  Range("P3").Value = dectick
  Range("P4").Value = voltick
  
  
  Range("Q1").Value = "Value"
  Range("Q2").Value = maxinc
  Range("Q3").Value = maxdec
  Range("Q4").Value = largestvol

  Range("Q2:Q3").NumberFormat = "0.00%"
  Range("O:O").Columns.AutoFit
  Range("Q:Q").Columns.AutoFit

  ' The following formatting was just added to make some room on the sheet so
  'everything could display on the screenshot in the proper format
  
  Range("H:H").ColumnWidth = 5
  Range("M:M").ColumnWidth = 5
  Range("N:N").ColumnWidth = 5
  Range("A:G").ColumnWidth = 9
  
End Sub