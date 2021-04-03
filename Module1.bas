Attribute VB_Name = "Module1"
Sub Alphabet_Single()

  Cells(1, 9) = "Ticker"
  Cells(1, 10) = "Yearly Change"
  Cells(1, 11) = "Percent Change"
  Cells(1, 12) = "Total Stock Volume"
  
  Dim Letter As String
  
  Dim End_Row As Long, i As Long
  
  Dim Yearly_change As Double
  
  Dim Letter_vol As Double
  
  Letter_vol = 0
  
  Yearly_open = Range("C2").Value
  
  'Percent_change = 0
  
  Dim Summary_Table_Row As Integer
  
  Summary_Table_Row = 2
  
  End_Row = Cells(Rows.Count, "A").End(xlUp).Row

  For i = 2 To End_Row
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Letter = Cells(i, 1).Value
      
      Yearly_close = Range("F" & i).Value

      Letter_vol = Letter_vol + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = Letter

      Range("J" & Summary_Table_Row).Value = Yearly_close - Yearly_open
      
      Range("K" & Summary_Table_Row).Value = Cells(1 - (Yearly_close / Yearly_open)) * 100
      
      Range("L" & Summary_Table_Row).Value = Letter_vol

      Summary_Table_Row = Summary_Table_Row + 1
      
      Letter_vol = 0
      
      Yearly_open = Range("C" & i).Value
      
    Else

      Letter_vol = Letter_vol + Cells(i, 7).Value
    
    'If Yearly_change > 0 Then
        
        'Cells(i, 10).Interior.ColorIndex = 4
    'Else
        'Cells(i, 10).Interior.ColorIndex = 3

    End If
    
  Next i
  
  'Columns("K").NumberFormat = "0.00%"
  Columns("A:L").EntireColumn.AutoFit

End Sub
