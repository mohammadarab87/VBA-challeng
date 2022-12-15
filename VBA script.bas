
Sub Stock_Market():

For Each ws In Worksheets
Dim WorksheetName As String


Dim i As Long
Dim j As Long
Dim TickerSymbol As Long
Dim LastRowA As Long
'-------------------------------------------------------------------

Dim LastRowI As Long
Dim PerChange As Double
Dim GreatIncr As Double
Dim GreatDecr As Double
Dim GreatVol As Double

WorksheetMane = ws.Name
        
'Create column header for Ticker , yearley Change , Percent Change , Total Stock
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------
        
'TickerSymbol first row is #2
TickerSymbol = 2
j = 2

' loop through non-blank cell in column A (#1)
LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
For i = 2 To LastRowA

    
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
'print TickerSymbol in col I (#9)
ws.Cells(TickerSymbol, 9).Value = ws.Cells(i, 1).Value

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------


'Calculate and print Yearly Change in column J (#10)

ws.Cells(TickerSymbol, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value

'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'Calculate and print percent change in column K (#11)
If ws.Cells(j, 3).Value <> 0 Then
PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
'format the percent
ws.Cells(TickerSymbol, 11).Value = Format(PerChange, "Percent")
Else
ws.Cells(TickerSymbol, 11).Value = Format(0, "Percent")
End If

'----------------------------------------------------------------------------
'----------------------------------------------------------------------------

'Calculate and print total volume in column L (#12)
ws.Cells(TickerSymbol, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(j, 7), ws.Cells(i, 7)))

'---------------------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------------------

 'Conditional formating
                    
If ws.Cells(TickerSymbol, 10).Value > 0 Then

'cell color is green
ws.Cells(TickerSymbol, 10).Interior.ColorIndex = 4
Else
'cell color is red
ws.Cells(TickerSymbol, 10).Interior.ColorIndex = 3
End If
'------------------------------------------------------------------------------
If ws.Cells(TickerSymbol, 11).Value > 0 Then

'cell color is green
ws.Cells(TickerSymbol, 11).Interior.ColorIndex = 4
Else
'cell color is red
ws.Cells(TickerSymbol, 11).Interior.ColorIndex = 3
End If
'--------------------------------------------------------------------------------
'--------------------------------------------------------------------------------
TickerSymbol = TickerSymbol + 1
j = i + 1
 End If
Next i
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'great header for the following
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------

'loop through non-blank cell in column I (#9)
LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
'assing value
GreatIncr = ws.Cells(2, 11).Value
GreatDecr = ws.Cells(2, 11).Value
GreatVol = ws.Cells(2, 12).Value
  For i = 2 To LastRowI
'---------------------------------------------------------------
'for Greatest Ingrease
If ws.Cells(i, 11).Value > GreatIncr Then
GreatIncr = ws.Cells(i, 11).Value
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
Else
GreatIncr = GreatIncr
End If
'---------------------------------------------------------------------
'For greatest decrease
If ws.Cells(i, 11).Value < GreatDecr Then
GreatDecr = ws.Cells(i, 11).Value
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
Else
GreatDecr = GreatDecr
End If
'---------------------------------------------------------------------
'for gretest Volume
If ws.Cells(i, 12).Value > GreatVol Then
GreatVol = ws.Cells(i, 12).Value
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
Else
GreatVol = GreatVol
End If
'-------------------------------------------------------------------------
ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")

Next i

Next ws
End Sub

