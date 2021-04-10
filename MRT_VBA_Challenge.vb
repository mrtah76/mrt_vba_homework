Option Explicit
Sub Stockticker()
'   column 1 is ticker
'   column 2 is date
'   column 3 is open
'   column 4 is high
'   column 5 is low
'   column 6 is close
'   column 7 is vol
'   column 8 is vol 000's

'I will use this to Loop Through All Worksheets
        Dim WS_Count As Integer
        Dim I As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
             WS_Count = ActiveWorkbook.Worksheets.Count
         
         ' Begin the loop.
            For I = 1 To WS_Count
            Sheets(I).Select
                     
'Before I start, I want to sort first by the ticker,then date
Columns("A:G").Select
    ActiveWorkbook.Worksheets(I).Sort.SortFields.Clear
    ActiveWorkbook.Worksheets(I).Sort.SortFields.Add2 Key:=Range("A:A"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets(I).Sort.SortFields.Add2 Key:=Range("B:B"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(I).Sort
        .SetRange Range("A:G")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'Here I define Headers
    Cells(1, 10).Value = "<Ticker>"
    Cells(1, 11).Value = "<Gain/Loss>"
    Cells(1, 12).Value = "<Percent Change>"
    Cells(1, 13).Value = "<Total Vol.000's>"
    Columns("J:M").AutoFit
    'I want to header<Gain/Loss> show Gain in Green/Loss in Red
        Dim P, L As Integer
        P = InStr(Cells(1, 11).Value, "/")
        L = Len(Cells(1, 11).Value)
        Cells(1, 11).Characters(Start:=P - 4, Length:=L - 7).Font.ColorIndex = 4
        Cells(1, 11).Characters(Start:=P + 1, Length:=L).Font.ColorIndex = 3
        Range("J1:Q1").Font.Bold = True
    

'summarize tickers into a single ticker
    ' I will set column 10 for ticker
    ' start from Row 2
    ' I need If to compair tickers and stop when there is change
    Dim F_data_row, L_Data_row As Long
    Dim Data_row As Long
    Dim outputrow As Long
    Dim vol As Long
    vol = 0
    outputrow = 2
    F_data_row = 2
        
        For Data_row = 2 To Range("A2").End(xlDown).Row
    
    ' this where ticker changes & I establish the Last Data in the ticker
    
    If Cells(Data_row, 1).Value <> Cells(Data_row + 1, 1) Then
       L_Data_row = Data_row
       
        ' Calculate yearly change&color Format
        'I will set column 11 for yearly change
        'get the opening price for the first day of the year
        'get the closing price for the last day of the year
        'format cells for +green&-Red
    Cells(outputrow, 11).Value = Cells(L_Data_row, 6).Value - Cells(F_data_row, 3)
        If Cells(outputrow, 11).Value > 0 Then
            Cells(outputrow, 11).Interior.ColorIndex = 4
        Else
            Cells(outputrow, 11).Interior.ColorIndex = 3
        End If
    'calculate percentage change
        On Error Resume Next
    Cells(outputrow, 12).Value = Cells(outputrow, 11).Value / Cells(F_data_row, 3).Value
    Cells(outputrow, 12).NumberFormat = "0.00%"
    Cells(outputrow, 13).Value = vol
    vol = 0
        'Going to the next Ticker in Output
        outputrow = outputrow + 1
        F_data_row = L_Data_row + 1
        
        
'I need to establish a position on column 10 for current ticker
    Else: Cells(outputrow, 10).Value = Cells(Data_row, 1).Value
    'calculate Total vol while we are on the same Ticker
    vol = vol + Cells(Data_row, 7) / 1000
    
End If
Next Data_row


  
    
    
' This is where I will score My Bonus!!! LOL
'setting the header
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Columns("O:Q").AutoFit
'calculating Max,Min,Total
Dim Gr_inc, Gr_dec As Double
Dim Gr_vol As Long
Dim Mx_rw, Min_rw, Vol_rw As Double
'First Max Value and Ticker
    Gr_inc = WorksheetFunction.Max(Range("L:L"))
    Range("Q2") = Gr_inc
    Range("Q2").NumberFormat = "0.00%"
    Mx_rw = WorksheetFunction.Match(Gr_inc, Range("L:L"), 0)
    Cells(2, 16).Value = Cells(Mx_rw, 10).Value
    Cells(2, 16).Font.ColorIndex = 5
'Second Min Value and Ticker
    Gr_dec = WorksheetFunction.Min(Range("L:L"))
    Range("Q3") = Gr_dec
    Range("Q3").NumberFormat = "0.00%"
    Min_rw = WorksheetFunction.Match(Gr_dec, Range("L:L"), 0)
    Cells(3, 16).Value = Cells(Min_rw, 10).Value
    Cells(3, 16).Font.ColorIndex = 5
'Third Max Total Value & Ticker
    Gr_vol = WorksheetFunction.Max(Range("M:M"))
    Range("Q4") = Gr_vol
    Range("Q4").NumberFormat = "0"
    Vol_rw = WorksheetFunction.Match(Gr_vol, Range("M:M"), 0)
    Cells(4, 16).Value = Cells(Vol_rw, 10).Value
    Cells(4, 16).Font.ColorIndex = 5
'MsgBox (Min_rw)







'Here we go to the next sheet
Next I
MsgBox ("ALL WORK IS DONE")
Sheets(1).Select

End Sub
