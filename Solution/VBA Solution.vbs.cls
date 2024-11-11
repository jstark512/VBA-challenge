VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis():

'set dimensions

    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

'set title row

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("P2").Value = "Greatest % Increase"
    Range("P3").Value = "Greatest % Decrease"
    Range("P4").Value = "Greatest Total Volume"

'set initial values

    j = 0
    total = 0
    change = 0
    start = 2

'row number with data

    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount


        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then


'store results

            total = total + Cells(i, 7).Value

'handle zero

            If total = 0 Then
            
'print results

                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
            
'first non zero starting value

                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

'calculate change

                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)

'start of ticker

                start = i + 1

'print results

                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("J" & 2 + j).NumberFormat = "0.00"
                Range("K" & 2 + j).Value = percentChange
                Range("K" & 2 + j).NumberFormat = "0.00%"
                Range("L" & 2 + j).Value = total

'colors positives and negatives

                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

'reset variables for new ticker

            total = 0
            change = 0
            j = j + 1
            days = 0

'add same results

        Else
            total = total + Cells(i, 7).Value

        End If
        
'max and min move to another part of worksheet

    Next i
    Range("O2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("O3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("O4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

'return less

    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

End Sub


