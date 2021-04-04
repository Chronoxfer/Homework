
Sub Alphabet_Single()

  'Labled Header
  Cells(1, 9) = "Ticker"
  Cells(1, 10) = "Yearly Change"
  Cells(1, 11) = "Percent Change"
  Cells(1, 12) = "Total Stock Volume"
  
  'Dim Settings
  Dim Letter As String
  Dim End_Row As Long, i As Long
  Dim Yearly_change As Double
  Dim Letter_vol As Double
  Dim first_open As Double
  Dim Summary_Table_Row As Integer
  Dim Yearly_close As Double
  
  first_open = Cells(2, 3).Value
  
  Letter_vol = 0
  
  Summary_Table_Row = 2
  
  End_Row = Cells(Rows.Count, "A").End(xlUp).Row

  'Main For Loop
  For i = 2 To End_Row
    
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
      Yearly_close = Range("F" & i).Value
      
      Yearly_change = Yearly_close - first_open

      Percent_change = ((Yearly_close / first_open) - 1)
      
      Letter = Cells(i, 1).Value

      Letter_vol = Letter_vol + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = Letter
      
      Range("J" & Summary_Table_Row).Value = Yearly_close - first_open
      
      Range("K" & Summary_Table_Row).Value = Percent_change
      
      Range("L" & Summary_Table_Row).Value = Letter_vol

      Summary_Table_Row = Summary_Table_Row + 1
      
      Letter_vol = 0
      
      first_open = Cells(i + 1, 3).Value
      
    Else

      Letter_vol = Letter_vol + Cells(i, 7).Value
    
    End If
    
Next i

    'format columns colors
    Dim rg As Range
    Dim g As Long
    Dim c As Long
    Dim color_cell As Range
    
    Set rg = Range("J2", Range("J2").End(xlDown))
    c = rg.Cells.Count
    
    For g = 1 To c
    Set color_cell = rg(g)
    Select Case color_cell
        Case Is >= 0
            With color_cell
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With color_cell
                .Interior.Color = vbRed
            End With
       End Select
    Next g

  'Formating and Auto fit
  Columns("K").NumberFormat = "0.00%"
  Columns("A:L").EntireColumn.AutoFit

End Sub


