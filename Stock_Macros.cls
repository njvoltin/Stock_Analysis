VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabetical_testing()
    Dim i As Long
    Dim ws As Worksheet
    Dim column As Integer
    Dim lastrow As Long
    Dim nextEmptyRowH As Long
    
    ' Loop through all worksheets
    For Each ws In Worksheets
        ' Set the header for column H
        ws.Cells(1, 8).Value = "Ticker"
        
        column = 1 ' Assuming column A for comparison
        lastrow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row ' Get the last row in column A
        
        ' Loop through each row from row 2 to last row
        For i = 2 To lastrow
            ' If the value in the current row is different from the next row in column A
            If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                ' Find the next empty row in column H
                nextEmptyRowH = ws.Cells(ws.Rows.Count, 8).End(xlUp).Row + 1
                ' Copy the value from column A to the next empty cell in column H
                ws.Cells(nextEmptyRowH, 8).Value = ws.Cells(i, column).Value
            End If
        Next i
    Next ws
End Sub

Sub Quarterly_Change()

    Dim i As Long
    Dim ws As Worksheet
    Dim column As Integer
    Dim lastrow As Long
    Dim nextEmptyRowI As Long
        
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Quarterly Change"
        column = 1
        lastrow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
        firstRow = 2
        For i = 3 To lastrow
           If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                nextEmptyRowI = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row + 1
                diff = ws.Cells(i, 6).Value - ws.Cells(firstRow, 3).Value
                ws.Cells(nextEmptyRowI, 9).Value = diff
                If diff > 0 Then
                    ws.Cells(nextEmptyRowI, 9).Interior.ColorIndex = 4
                ElseIf diff < 0 Then
                    ws.Cells(nextEmptyRowI, 9).Interior.ColorIndex = 3
                Else
                    ws.Cells(nextEmptyRowI, 9).Interior.ColorIndex = 0
                End If
                firstRow = i + 1
           End If
        Next i
      Next ws
End Sub

Sub Percent_Change()

    Dim i As Long
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim firstRow As Long
    Dim nextEmptyRowj As Long
    Dim begquarter As Double
    Dim endquarter As Double
    Dim percentchange As Double
    
    For Each ws In Worksheets
        ws.Cells(1, 10).Value = "Percent Change"
        column = 1
        lastrow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
        firstRow = 2
        For i = 3 To lastrow
           If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                nextEmptyRowj = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row + 1
                endquarter = ws.Cells(i, 6).Value
                begquarter = ws.Cells(firstRow, 3).Value
                percentchange = ((endquarter - begquarter) / begquarter) * 100
                ws.Cells(nextEmptyRowj, 10).Value = percentchange & "%"
                If percentchange > 0 Then
                    ws.Cells(nextEmptyRowj, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(nextEmptyRowj, 10).Interior.ColorIndex = 3
                    End If
                firstRow = i + 1
            End If
        Next i
    Next ws
End Sub


Sub Total_Volume()

    Dim i As Long
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim firstRow As Long
    Dim nextEmptyRowk As Long
    Dim totalsum As Double
    
        For Each ws In Worksheets
        ws.Cells(1, 11).Value = "Total Volume"
        column = 1
        lastrow = ws.Cells(ws.Rows.Count, column).End(xlUp).Row
        firstRow = 2
        For i = 3 To lastrow
           If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
                nextEmptyRowk = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row + 1
                totalsum = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(firstRow, 7), ws.Cells(i, 7)))
                ws.Cells(nextEmptyRowk, 11).Value = totalsum
                firstRow = i + 1
            End If
        Next i
    Next ws
            
End Sub

Sub Increase_decrease()

Dim i As Long
Dim ws As Worksheet
Dim lastrow As Long
Dim percentchange As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestincreaserow As Long
Dim greatestdecreaserow As Long

greatestincrease = -999
greatestdecrease = 999

    For Each ws In Worksheets
        ws.Cells(2, 14).Value = "Ticker"
        ws.Cells(2, 15).Value = "Value"
        ws.Cells(3, 13).Value = "Greatest % Increase"
        ws.Cells(4, 13).Value = "Greatest % Decrease"
        lastrow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        For i = 2 To lastrow
            percentchange = ws.Cells(i, 10).Value
            If percentchange > greatestincrease Then
                greatestincrease = percentchange
                greatestincreaserow = i
            End If
            If percentchange < greatestdecrease Then
                greatestdecrease = percentchange
                greatestdecreaserow = i
            End If
        Next i
        ws.Cells(3, 15).Value = greatestincrease * 100 & "%"
        ws.Cells(4, 15).Value = greatestdecrease * 100 & "%"
        ticker = ws.Cells(greatestincreaserow, 8).Value
        ws.Cells(3, 14).Value = ticker
        ticker1 = ws.Cells(greatestdecreaserow, 8).Value
        ws.Cells(4, 14).Value = ticker1
        
    Next ws
            
End Sub

Sub Total_Volume_Table()
    Dim i As Long
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim currenttotal As Double
    Dim largesttotal As Double
    Dim largesttotalrow As Double
    
        For Each ws In Worksheets
            ws.Cells(5, 13).Value = "Greatest Total Volume"
            largesttotal = 0
            lastrow = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
            For i = 2 To lastrow
                currenttotal = ws.Cells(i, 11).Value
                If currenttotal > largesttotal Then
                    largesttotal = currenttotal
                    largesttotalrow = i
                End If
            Next i
            ws.Cells(5, 15).Value = largesttotal
            ticker2 = ws.Cells(largesttotalrow, 8).Value
            ws.Cells(5, 14).Value = ticker2
        Next ws
            
End Sub
