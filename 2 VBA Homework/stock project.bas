Attribute VB_Name = "stock"

Sub stock()
' Setting variable types
Dim Stock_Name As String
Dim Total_Stock_Volumn As Double
Dim Sum_Table_Row As Long
Dim i As Long
Dim LastRow As Long
Dim Year_Opening_Price As Double
Dim Year_Closing_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVolumn As Double
Dim percentlastrow As Long
Dim ws As Worksheet
Dim mincell As Range

For Each ws In ThisWorkbook.Worksheets
ws.Activate





'headers
Range("i1") = "Ticker"
Range("j1") = "PriceChange"
Range("k1") = "Percent Change"
Range("l1") = "Total Volume"
Range("n2") = "Greatest % Increase"
Range("n3") = "Greatest % of Decrease"
Range("n4") = "Greatest total volume"
Range("o1") = "value"
Range("p1") = "Ticker"


' Setting initial variable values
Sum_Table_Row = 2
YearOpenFlag = False
Year_Opening_Price = Cells(2, 3).Value

' Determining number of last row and percent last row
LastRow = Range("A" & Rows.Count).End(xlUp).Row




' Proving rows were counted
' MsgBox (LastRow)
' Loop for finding names and volumn
For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   Stock_Name = Cells(i, 1).Value

   ' ***** Setting the stock's year opening price
   If Not YearOpenFlag Then
       Year_Opening_Price = Cells(i, 3).Value
       YearOpenFlag = True
   End If
    ' ***** End of first time loop code
Total_Stock_Volumn = Total_Stock_Volumn + Cells(i, 7).Value
Range("I" & Sum_Table_Row).Value = Stock_Name
Range("L" & Sum_Table_Row).Value = Total_Stock_Volumn
Sum_Table_Row = Sum_Table_Row + 1
Total_Stock_Volumn = 0
Else
   Year_Closing_Price = Cells(i + 1, 6).Value
   Total_Stock_Volumn = Total_Stock_Volumn + Cells(i, 7).Value
   Yearly_Change = Year_Closing_Price - Year_Opening_Price
   Range("J" & Sum_Table_Row).Value = Yearly_Change
   'Add coloring formatter
   If Yearly_Change < 0 Then
   Range("J" & Sum_Table_Row).Interior.ColorIndex = 3
   ElseIf Yearly_Change > 0 Then
   Range("J" & Sum_Table_Row).Interior.ColorIndex = 4
   End If
   
   ' Calculate percent change
On Error Resume Next
   Percent_Change = (Year_Closing_Price / Year_Opening_Price) - 1
       Range("K2:lastrow").Select
    Selection.NumberFormat = "0.00%"
   Range("K" & Sum_Table_Row).Value = Percent_Change
   
   


   ' Resetting Flags
       YearOpenFlag = False

End If
Next i

With ActiveSheet

percentlastrow = .Cells(.Rows.Count, "e").End(xlUp).Row
MsgBox (percentlastrow)


MinPercent = Application.WorksheetFunction.Min(Range(Cells(2, 11), Cells(percentlastrow, 11)))
Set mincell = .Cells(.Rows.Count, "e").End(xlUp).Row.Find(what:=MinPercent, LookIn:=xlValues)

'MsgBox (mincell)

MaxPercent = Application.WorksheetFunction.Max(Range(Cells(2, 11), Cells(percentlastrow, 11)))
MsgBox (MinPercent)
Range("o3") = MinPercent
Range("o2") = MaxPercent
   Range("o2:o3").Select
    Selection.NumberFormat = "0.00%"
    
maxvolume = Application.WorksheetFunction.Max(Range(Cells(2, 12), Cells(percentlastrow, 12)))
Range("o4") = maxvolume
Range("o4").Select
Selection.NumberFormat = general

 

End With


Next ws


End Sub

