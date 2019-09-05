Sub RealStock()

'define my variables

Dim N As Long
Dim sum As Double
Dim Count As Long
Dim ws As Worksheet
Dim lastrow As Long
Dim lr As Long

Dim MyMax As Double
Dim MyMin As Double
Dim VMax As Double
    
Dim tkrs1 As String
Dim tkrs2 As String
Dim tkrs3 As String

'Looping over each workcheet of the workbook
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

   'Setting up my counters 
    Count = 0
    N = 2
    sum = 0

    'retreiving the Last row count 
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
   'setting up my headers 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly change"
    ws.Cells(1, 11).Value = "percent change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest total volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Looping through all rows 
    For i = 2 To lastrow
        'summing all values in column 7 if they have the same Ticker
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
            sum = sum + ws.Cells(i, 7).Value
            Count = Count + 1
        'Otherwise store the sums and re-set all my variable to 0     
        Else
            ws.Cells(N, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(N, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i - Count, 3).Value
            
            'Coloring my cells based on their value, positive green, negative red
            If ws.Cells(N, 10).Value < 0 Then
                ws.Cells(N, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(N, 10).Interior.ColorIndex = 4
            End If

            'setting if conditon to avoind deivision by 0
            If ws.Cells(i - Count, 3).Value = 0 Then
            ws.Cells(N, 11).Value = "0"
            Else
            ws.Cells(N, 11).Value = (ws.Cells(N, 10).Value / ws.Cells(i - Count, 3).Value)
            ws.Cells(N, 11).NumberFormat = "0.00%"
            End If

            'Calculating the sum of the value in column 7
            ws.Cells(N, 12).Value = sum + ws.Cells(i, 7).Value

            'resetting my counters 
            Count = 0
            sum = 0
            N = N + 1
        End If
    Next i
    
    'retreiving my last row in my new table
    lr = ws.Cells(Rows.Count, 9).End(xlUp).Row
    'looping through the table
    For i = 2 To lr
        'calculating the maximum value in column 11, and retreiving its corresponding value in column 9 
        If ws.Cells(i, 11).Value > MyMax Then
            MyMax = ws.Cells(i, 11).Value
            tkrs1 = ws.Cells(i, 9).Value
        End If
        'calculating the minimum value in column 11, and retreiving its corresponding value in column 9 
        If ws.Cells(i, 11) < MyMin Then
            MyMin = ws.Cells(i, 11).Value
            tkrs2 = ws.Cells(i, 9).Value
        End If
        'calculating the maximum value in column 12, and retreiving its corresponding value in column 9 
        If ws.Cells(i, 12) > VMax Then
            VMax = ws.Cells(i, 12).Value
            tkrs3 = ws.Cells(i, 9).Value
        End If
    Next i
    
    'Assigning the calculated values to cells
    ws.Cells(2, 17).Value = MyMax
    ws.Cells(3, 17).Value = MyMin
    'converting values to percentage format
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 17).Value = VMax
    ws.Cells(2, 16) = tkrs1
    ws.Cells(3, 16) = tkrs2
    ws.Cells(4, 16) = tkrs3
  'Resetting my values for next sheet  
  MyMax = 0
  MyMin = 0
  VMax = 0

Next

End Sub



