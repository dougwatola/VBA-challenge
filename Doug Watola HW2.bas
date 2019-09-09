Attribute VB_Name = "Module1"
Sub stock_analysis()

  ' Set an initial variable for holding the brand name
  Dim Ticker As String
  
  Dim Volume As Double
  Volume = 0
  
  'Set Summary Row variable
  Dim Summary_Row As Double
  Summary_Row = 2
  
  'Variable to Capture the Opening Price
  Dim Opening As Double
  
  'Variable to Capture the Closing Price
  Dim Closing As Double
  
  'Year Opening to capture Opening value in loop
  Dim Year_Opening As Double
  
  'This will be used to calculate percentage
  Dim Yearly_Change As Double
  
  Dim Percent_Change As Double
  
  Dim LastRow As Double
  
  'Iterate through values until we have an opening price
  Dim Opening_zero As Integer
  Opening_zero = 1
  
  'Loop through all worksheets
  
  For Each ws In Worksheets
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  Summary_Row = 2
  
  'Populate Column Labels for each Worksheet
  ws.Range("I" & 1).Value = "Ticker"
  ws.Range("J" & 1).Value = "Yearkly Change"
  ws.Range("K" & 1).Value = "Percent Change"
  ws.Range("L" & 1).Value = "Total Stock Volume"

  ' Loop through all columns and collect volume sum
  For i = 2 To LastRow
  
      
    ' Check if we are still within the same credit card brand, if we are not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        
        Ticker = (ws.Cells(i, 1).Value)
        Volume = Volume + ws.Cells(i, 7).Value
              
        ws.Range("I" & Summary_Row).Value = Ticker
        ws.Range("L" & Summary_Row).Value = Volume
              
        Summary_Row = Summary_Row + 1
        Volume = 0
      
    Else
        
        Volume = Volume + ws.Cells(i, 7).Value

    End If

  Next i

'Initialize Summary Row for the next loop through data
Summary_Row = 2

' Loop through all columns and collect Yearly and Percent Change
  For i = 2 To LastRow
  
    
    If (Year_Opening = 0) Then
        Opening = (ws.Cells(i, 3).Value)
            If Opening = 0 Then
            'This is one of those times that I lothe coding
            'A hard set to 1 should work
            Opening = 1
            'Opening = (ws.Cells(i + Opening_zero, 3).Value)
            'Opening_zero = Opening_zero + 1
            Else
            'Opening_zero = 1
            End If
        Year_Opening = 1
    End If
      
    ' Check if we are still within the same credit card brand, if we are not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        
        Closing = (ws.Cells(i, 6).Value)
        
           
        'Range("N" & Summary_Row).Value = Closing
        'Range("M" & Summary_Row).Value = Opening
        
        'Calculate Yearly Change
        Yearly_Change = Closing - Opening
        ws.Range("J" & Summary_Row).Value = Yearly_Change
        
        'Set conditional formatting for yearly change, red if negative, green if positive
        If Yearly_Change > 0 Then
            ws.Range("J" & Summary_Row).Interior.Color = vbGreen
        Else
            ws.Range("J" & Summary_Row).Interior.Color = vbRed
        End If
        
        'Calculate Percent Change
        Percent_Change = Yearly_Change / Opening
        ws.Range("K" & Summary_Row).Value = Format(Percent_Change, "Percent")
              
        Summary_Row = Summary_Row + 1
        
        Year_Opening = 0
      
    End If

  Next i

Next ws

End Sub

