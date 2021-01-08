Attribute VB_Name = "Module2"
Sub Stock_Analyzer()
'
' Stock_Analyzer Macro
'Sub Stock_Analyzer()

  ' Set an initial variable for holding the ticker
  Dim Stock_Ticker As String

  ' Set an initial variable for holding the total volume per ticker
  Dim Stock_Volume_Total As Double
  Stock_Volume_Total = 0
 

  ' Keep track of the location for stock in the summary table
  Dim Summary_Table_Row As Integer

  
  'Nominal Change row
  Yearly_Change_Row = 2
  

  'opening price
  Dim Open_Price_Row As Double
  'Dim Open_Price As Long
  'Dim Close_Price_Row As Long
  'Dim Percent_Change As Double


  
  'worksheet loop
  For Each ws In Worksheets
  WorksheetName = ws.Name
    MsgBox WorksheetName
    Summary_Table_Row = 2
    
    Open_Price_Row = 2

  
  
    'Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stock symbols
  For i = 2 To LastRow

    ' Check if we are still within the same stock ticker, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
      ' Set the ticker name
      Stock_Ticker = ws.Cells(i, 1).Value
      
       'close price variable
      Close_Price = ws.Cells(i, 6).Value
      'MsgBox CStr(Close_Price) & CStr(Stock_Ticker)
        'Open price number
      Open_Price = ws.Cells(Open_Price_Row, 3).Value
        ' calculates yearly change
      Yearly_Change = Close_Price - Open_Price
      'percent change
      
      If Open_Price <> 0 Then
        Percent_Change = ((Close_Price - Open_Price) / Open_Price)
      Else
        Percent_Change = 0
      End If

      ' Add to the Brand Total
      Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value

      ' Print the stock ticker symbol in the Summary Table
      ws.Range("i" & Summary_Table_Row).Value = Stock_Ticker
      ' Print the total volume to the Summary Table
      ws.Range("l" & Summary_Table_Row).Value = Stock_Volume_Total
      ' Print the yearly change to the Summary Table
      ws.Range("j" & Summary_Table_Row).Value = Yearly_Change
      ' Print the percent change to the Summary Table
      ws.Range("k" & Summary_Table_Row).Value = Percent_Change
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

    
      ' Reset the Brand Total
      Stock_Volume_Total = 0
      'Reset open price
      Open_Price_Row = i + 1
  
      

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the ticker Total
      Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value

    End If

  Next i
  
LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

    For j = 2 To LastRow2

   If ws.Cells(j, 10).Value < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    ElseIf ws.Cells(j, 10).Value >= 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
   End If
    
    ws.Cells(j, 11).Style = "Percent"
    
  
    Next j


End Sub

