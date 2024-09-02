Sub GetRowHeights()
  Dim row As Integer
  Dim height As Double
  Dim table As Integer
  Dim index As Integer
  Dim endIndex As Integer
  

  ' Set the range of rows to check
  For table = 101 To 200
    If table = 0 Then
        index = 1
    Else
        index = 1 + (39 * table)
    End If
    'endIndex = 39 + (40 * table)
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 27
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 36
    index = index + 1
    Rows(index).RowHeight = 18.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 12.25
    index = index + 1
    Rows(index).RowHeight = 11.25
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 12.75
    index = index + 1
    Rows(index).RowHeight = 17
    index = index + 1
    Rows(index).RowHeight = 19.75
    index = index + 1
    Rows(index).RowHeight = 19.75
    index = index + 1
    Rows(index).RowHeight = 19.75
    index = index + 1
    Rows(index).RowHeight = 19.75
    index = index + 1
    Rows(index).RowHeight = 19.75
    index = index + 1
    Rows(index).RowHeight = 19.75
    Next table
  '11.25 x
  '11,25x
  '11,25x
  '22,5x
  '11,25x
  '34x
  '18,25x
  '11,25x
  '11,25x
  '11,25x
  '11,25x
  '11,25x
  '11,25x
  '11,25x
  '11,25x
  '12,25x
  '11,25x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '12,75x
  '15x
  '19,75
  '19,75
  '19,75
  '19,75
  '19,75
  '19,75
  
  
  
  
  
End Sub


