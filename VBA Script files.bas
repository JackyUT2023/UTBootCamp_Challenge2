Attribute VB_Name = "Module1"
Sub tickerchallenge()

'loop through all sheets
Dim m As Integer
Dim ws_Num As Integer
Dim ws As Worksheet
ws_Num = ThisWorkbook.Worksheets.Count
For m = 1 To ws_Num
ThisWorkbook.Worksheets(m).Activate


'Set Title Row and formatting
TitleRow = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")
Range("I1:L1") = TitleRow
Range("I1:L1").Font.Bold = True
'--------------------------------------------------------------------

'Filter Ticker
  Dim Ticker_Name As String
  Dim Ticker_Count As Integer
  Ticker_Count = 0
  Dim Stock_Volume As Double
    Stock_Volume = 0
  Dim Opening_Price As Double
  Dim Closing_Price As Double
  
  Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
  Dim lRow As Long
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all ticker
  For i = 2 To lRow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      Ticker_Name = Cells(i, 1).Value
      Stock_Volume = Stock_Volume + Cells(i, 7).Value
      Opening_Price = Cells(i - Ticker_Count, 3).Value
      Closing_Price = Cells(i, 6).Value
      
      Range("i" & Summary_Table_Row).Value = Ticker_Name
      Range("j" & Summary_Table_Row).Value = Closing_Price - Opening_Price
      Range("j" & Summary_Table_Row).NumberFormat = "0.00"
      Range("k" & Summary_Table_Row).Value = (Closing_Price - Opening_Price) / Opening_Price
      Range("k" & Summary_Table_Row).NumberFormat = "0.00%"
      Range("l" & Summary_Table_Row).Value = Stock_Volume
      
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock_Volume
      Stock_Volume = 0
      Ticker_Count = 0
      
    ' If the cell immediately following a row is the same Ticker...
    Else
      Stock_Volume = Stock_Volume + Cells(i, 7).Value
      Ticker_Count = Ticker_Count + 1
    
    End If

  Next i
'---------------------------------------------------------------------------------
'formating
Dim j As Integer
Dim Summary_Table_lRow As Long
    Summary_Table_lRow = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To Summary_Table_lRow
    If Cells(j, 10).Value < 0 Then
    Cells(j, 10).Interior.ColorIndex = 3
    Else
    Cells(j, 10).Interior.ColorIndex = 4
    End If

Next j
'----------------------------------------------------------------------------------------
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".

'Set Title Row and formatting
TitleRow = Array("Ticker", "Value")
Range("O2") = "Greatest % increase"
Range("O3") = "Greatest % decrease"
Range("O4") = "Greatest total volume"
Range("P1:Q1") = TitleRow
Range("P1:Q1").Font.Bold = True
Range("O2:O4").Font.Bold = True
'-----------------------------------------------------------------------------
'find out value
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Greatest_Total_Volume As Double

Greatest_Increase = Application.WorksheetFunction.Max(Range("K2:K" & CStr(Summary_Table_lRow)))
Greatest_Decrease = Application.WorksheetFunction.Min(Range("K2:K" & CStr(Summary_Table_lRow)))
Greatest_Total_Volume = Application.WorksheetFunction.Max(Range("l2:l" & CStr(Summary_Table_lRow)))

Range("Q2") = Greatest_Increase
Range("Q2").NumberFormat = "0.00%"
Range("Q3") = Greatest_Decrease
Range("Q3").NumberFormat = "0.00%"
Range("Q4") = Greatest_Total_Volume

'find out ticker
Dim Ticker_GI As String
Dim Ticker_GD As String
Dim Ticker_GV As String

For j = 2 To Summary_Table_lRow
    If Cells(j, 11).Value = Greatest_Increase Then
        Ticker_GI = Cells(j, 9)
    End If
    
    If Cells(j, 11).Value = Greatest_Decrease Then
        Ticker_GD = Cells(j, 9)
    End If
    
    If Cells(j, 12).Value = Greatest_Total_Volume Then
        Ticker_GV = Cells(j, 9)
    End If
Next j

Range("P2") = Ticker_GI
Range("P3") = Ticker_GD
Range("P4") = Ticker_GV
'---------------------------------------------------------------------------------
'next worksheet
Next m

End Sub
