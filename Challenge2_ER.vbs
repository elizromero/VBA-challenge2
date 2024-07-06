Sub challenge2()

For Each ws In ThisWorkbook.Worksheets
ws.Activate
    'Headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly change"
    Range("K1").Value = "Percent change"
    Range("L1").Value = "Total stock volume"
    
    'Main Variables
    Dim tikcer As String
    Dim quarch As Double
    Dim perch As Double
    Dim totv As Double
    
    'Support variables
    Dim opprice As Double
    Dim clprice As Double
    
    'Last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row - 1
    
    'Results rows
    Dim result_row As Integer
    resultrow = 2
    
    'Inicializing opening price
    opprice = Cells(2, 3).Value
    
    'Loop
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        clprice = Cells(i, 6).Value
        quarch = clprice - opprice
        perch = (quarch / opprice)
        
        'Capture values
        Cells(resultrow, 9).Value = ticker
        Cells(resultrow, 10).Value = quarch
        Cells(resultrow, 11).Value = perch
        Cells(resultrow, 12).Value = totv
        
        'Format for percentage
        Cells(resultrow, 11).NumberFormat = "0.00%"
        
        'Conditional formating for quarch
        If quarch > 0 Then
            Cells(resultrow, 10).Interior.ColorIndex = 4
        ElseIf quarch < 0 Then
            Cells(resultrow, 10).Interior.ColorIndex = 3
        End If
        
        'Prepating next loop
        resultrow = resultrow + 1
        totv = 0
        opprice = Cells(i + 1, 3)
        
        Else
        totv = totv + Cells(i, 7).Value
        
        End If
        
    Next i
    
'Second part
    
    'Headers
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % decrease"
    Range("O4").Value = "Greatest total volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Variables
    Dim maxin As Double
    Dim maxdec As Double
    Dim maxvol As Double
    
    'New lastrow
    lastrow = Cells(Rows.Count, 11).End(xlUp).Row - 1
    
    'Initial values
    maxin = Cells(2, 11).Value
    maxdec = Cells(2, 11).Value
    maxvol = Cells(2, 12)
    
    'Loop for greatest % increase
    For j = 2 To lastrow
        If Cells(j + 1, 11).Value > maxin Then
        maxin = Cells(j + 1, 11).Value
        Range("P2").Value = Cells(j + 1, 9).Value
        Range("Q2").Value = maxin
        Range("Q2").NumberFormat = "0.00%"
        End If
    Next j
    
    'Loop for greatest % decrease
    For j = 2 To lastrow
        If Cells(j + 1, 11).Value < maxdec Then
        maxdec = Cells(j + 1, 11).Value
        Range("P3").Value = Cells(j + 1, 9).Value
        Range("Q3").Value = maxdec
        Range("Q3").NumberFormat = "0.00%"
        End If
    Next j
    
    'Loop for volume
    For j = 2 To lastrow
        If Cells(j + 1, 12).Value > maxvol Then
        maxvol = Cells(j + 1, 12).Value
        Range("P4").Value = Cells(j + 1, 9).Value
        Range("Q4").Value = maxvol
        End If
    Next j
    
'Autofit for all the result columns
ws.Columns("A:Q").AutoFit

Next ws

End Sub

