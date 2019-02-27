Attribute VB_Name = "Module1"

Sub StockTickerSummary()
    
'Prevent Computer Screen from running
  Application.ScreenUpdating = False
    ' Set an initial variable for holding the ticker
  Dim Stock_Symbol As String
    
  ' Set an initial variable for holding the total volume per ticker
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Dim openvalue As Double
  Dim closevalue As Double
  Dim percentchange As Double
  Dim rng As Range
  Dim myaddress As String
  Dim r As String
  Dim mylastrow As Long
  Dim WorksheetName As String
  
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
  For Each ws In ThisWorkbook.Worksheets
 
          ' Determine the Last Row for column A
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        ' Activate the Worksheet
        ws.Activate
        
        'Clear contents and formatting for all the columns after data
        ws.Range("I:Q").ClearContents
        ws.Range("I:Q").ClearFormats

        ' --------------------------------------------
        ' Summary Row
        ' --------------------------------------------

        ' Created a Variable to track of the location for each ticker in the summary table
        Summary_Table_Row = 2


        ' --------------------------------------------
        ' SORT
        ' --------------------------------------------
        
        ' Make Sure Sorted out  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        'Sort the rows based on the ticker and date
        ws.Sort.SortFields.Clear
        ws.Sort.SortFields.Add2 Key:=Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        ws.Sort.SortFields.Add2 Key:=Range("B2:B" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ws.Sort
            .SetRange Range("A1:G" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

        ' --------------------------------------------
        ' Headers
        ' --------------------------------------------
        ' Add the word Ticker Column Header
        ws.Range("I1").Value = "Ticker"
        ' Add the word Yearly Change Column Header
        ws.Range("J1").Value = "Yearly Change"
        ' Add the word Percent Change Column Header
        ws.Range("K1").Value = "Percent Change"
        ' Add the word Total Stock Volume Column Header
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Headers
          ws.Range("P1").Value = "Ticker"
          ws.Range("Q1").Value = "Value"
         'Row
          ws.Range("O2").Value = "Greatest % increase"
          ws.Range("O3").Value = "Greatest % decrease"
          ws.Range("O4").Value = "Greatest Total Volume"
        ' --------------------------------------------
        '  Loop
        ' --------------------------------------------

        ' Loop through all tickers
          For i = 2 To lastRow
                If i = 2 Then
                    openvalue = Cells(i, 3).Value
                End If
                ' Check if we are still within the same ticker, if it is not...
            
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                      ' Set the ticker
                      Stock_Symbol = Cells(i, 1).Value
                      
                      ' Get the close value for the ticker
                      closevalue = Cells(i, 6).Value
                
                      ' Add to the Volume Total
                      Volume_Total = Volume_Total + Cells(i, 7).Value
                      If openvalue > 0 Then  ' skipping stock with openvalue = 0
                              ' Print the ticker in the Summary Table
                              Range("I" & Summary_Table_Row).Value = Stock_Symbol
                              
                               ' Print the yearly change in the Summary Table
                            
                              Range("J" & Summary_Table_Row).Value = closevalue - openvalue
                              If (closevalue - openvalue) > 0 Then
                                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                              Else
                                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                              End If
                              ' Print the percent change in the Summary Table
                                    Range("K" & Summary_Table_Row).Value = Round((closevalue - openvalue) / (openvalue), 4)
                              Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                              ' Print the Brand Amount to the Summary Table
                              Range("L" & Summary_Table_Row).Value = Volume_Total
                        
                              ' Add one to the summary table row
                              Summary_Table_Row = Summary_Table_Row + 1
                        End If
                  
                        ' Reset the Volume Total
                        Volume_Total = 0
                        openvalue = Cells(i + 1, 3).Value
            
                ' If the cell immediately following a row is the same ticker...
                Else
                   
                  ' Add to the Volume Total
                  Volume_Total = Volume_Total + Cells(i, 7).Value
            
                End If
        
          Next i
          
        ' Finding last row in column K
         mylastrow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
         

          
         Set rng = ws.Range("K2:K" & mylastrow)
         myaddress = AddressOfMax(rng)
         Range("P2").Value = Range(myaddress).Offset(0, -2)
         Range("Q2").Value = Range(myaddress).Offset(0, 0)
         Range("Q2").NumberFormat = "0.00%"
         Set rng = ws.Range("K2:K" & mylastrow)
         myaddress = AddressOfMin(rng)
         Range("P3").Value = Range(myaddress).Offset(0, -2)
         Range("Q3").Value = Range(myaddress).Offset(0, 0)
         Range("Q3").NumberFormat = "0.00%"
         Set rng = ws.Range("L2:L" & mylastrow)
         myaddress = AddressOfMax(rng)
         Range("P4").Value = Range(myaddress).Offset(0, -3)
         Range("Q4").Value = Range(myaddress).Offset(0, 0)
         
         ' Autofit results
         ws.Columns("I:Q").AutoFit
         'Unselect the sheet if selected
         ws.Range("A1").Select

    ' --------------------------------------------
    ' PROCESS COMPLETE
    ' --------------------------------------------
Next
'Allow Computer Screen to refresh (not necessary in most cases)
  Application.ScreenUpdating = True
End Sub
Function AddressOfMax(rng As Range) As String
    AddressOfMax = WorksheetFunction.Index(rng, WorksheetFunction.Match(WorksheetFunction.Max(rng), rng, 0)).Address
End Function
Function AddressOfMin(rng As Range) As String
    AddressOfMin = WorksheetFunction.Index(rng, WorksheetFunction.Match(WorksheetFunction.Min(rng), rng, 0)).Address
End Function








