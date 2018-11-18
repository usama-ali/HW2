Attribute VB_Name = "Module1"
Sub Stock_Total()

' --------------------------------------------
' LOOP THROUGH ALL SHEETS
' --------------------------------------------
' Using For Each Loop as ws is an object
' --------------------------------------------
For Each ws In Worksheets
  
    ' Set an initial variable for holding the ticker
    Dim Ticker As String

    ' Set an initial variable for holding the total volume per ticker
    Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0

    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2


    ' Define and determine the Last Row
    Dim LastRow As Long
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    MsgBox ("Last row is: " & LastRow)


    ' Define the yearly change & percent change
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
  
    ' Set a variable to hold the open price at the beginning of Jan
    Dim Open_Jan_01 As Double
    Open_Jan_01 = Cells(2, "C").Value
  
    ' Set a variable to hold the close price at the end of Dec
    Dim Close_Dec_30 As Double
    Close_Dec_30 = Cells(2, "F").Value
  
    ' Loop through all ticker volumes
    For i = 2 To LastRow

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
        
        ' Set the ticker
        Ticker = Cells(i, "A").Value
        
        ' Add to the total stock volume
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, "G").Value
        
        ' Calculate the yearly change
        Close_Dec_30 = Cells(i, "F").Value
        Yearly_Change = Close_Dec_30 - Open_Jan_01
      
        ' Calculate the percent change
        If Open_Jan_01 <> 0 Then
        Percent_Change = 100 * (Yearly_Change / Open_Jan_01)
        Else
        Percent_Change = 0
        End If
        ' Print the ticker in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("I1").Value = "Ticker"
        
        ' Print the Yearly Change in the Summary Table
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        If Yearly_Change >= 0 Then
            Range("J" & Summary_Table_Row).Interior.Color = vbGreen
        Else
            Range("J" & Summary_Table_Row).Interior.Color = vbRed
        End If
        Range("J1").Value = "Yearly_Change"


        ' Print the Percent Change in the Summary Table
        Range("K" & Summary_Table_Row).Value = Percent_Change
        Range("K1").Value = "Percent Change"

        ' Print the total stock volume to the Summary Table
        Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
        Range("L1").Value = "Total Stock Volume"
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' reset the opening price for the next ticker
        Open_Jan_01 = Cells(i + 1, "C").Value

        ' Reset the total stock volume
        Total_Stock_Volume = 0

        ' If the cell immediately following a row is the same ticker...
        Else

        ' Add to the Brand Total
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, "G").Value
        
        End If
    
    Next i
    
Next ws

End Sub
