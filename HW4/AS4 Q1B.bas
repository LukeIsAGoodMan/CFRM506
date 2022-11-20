Attribute VB_Name = "InterfaceFunctions"
Option Explicit





Public Function CalcPV(FV As Double, CR As Double, CF As Double, SD As Date, ED As Date, MR As Double) As Double
    Dim dbond As Bond
    
    dbond.CouponFreq = CF
    dbond.FaceValue = FV
    dbond.CouponRate = CR
    dbond.SettlementDate = SD
    dbond.ExpiryDate = ED
    dbond.MktRate = MR
    
    Call CalcPresentValue(dbond)
    
    CalcPV = dbond.PresentValue
    
    
End Function

Public Sub FormatColumns()
  BondData.Activate
  Dim rng As Range
  Set rng = Range(Range("A1").Offset(0, 0), Range("A1").End(xlToRight))
  Worksheets("BondData").Names.Add Name:="ColNames", RefersTo:=rng
  
  With Range("ColNames")
    .Font.Italic = True
    .HorizontalAlignment = xlRight
    .Interior.Color = RGB(217, 217, 217)
    .Borders.LineStyle = xlContinuous
  End With
  
End Sub

Public Sub FormatData()
  BondData.Activate
  Dim rng As Range
  Set rng = Range(Range("A1").Offset(1, 0), Range("A1").End(xlDown).End(xlToRight))
  Worksheets("BondData").Names.Add Name:="InputData", RefersTo:=rng
  
  With Range("InputData")
    .HorizontalAlignment = xlRight
    .Interior.Color = RGB(226, 239, 218)
    .Borders.LineStyle = xlContinuous
  End With
  
  Dim rng1 As Range
  Dim rng2 As Range
  Dim i As Integer
  Dim j As Integer
  i = Worksheets("BondData").Range("InputData").Columns.Count
  j = Worksheets("BondData").Range("InputData").Rows.Count
  Set rng1 = Range(Range("A1").Offset(1, i), Range("A1").Offset(j, i))
  Worksheets("BondData").Names.Add Name:="pv", RefersTo:=rng1
  Set rng2 = Range(Range("A1").Offset(1, i + 1), Range("A1").Offset(j, i + 1))
  Worksheets("BondData").Names.Add Name:="ytm", RefersTo:=rng2
  
  With Range("pv")
    .HorizontalAlignment = xlRight
    .Interior.Color = RGB(221, 235, 247)
    .Borders.LineStyle = xlContinuous
  End With

  With Range("ytm")
    .HorizontalAlignment = xlRight
    .Interior.Color = RGB(221, 235, 247)
    .Borders.LineStyle = xlContinuous
  End With
  
  
End Sub
Public Sub CalcYieldToMaturity()
    'This subroutine will loop through each row in BondData worksheet,
    'use Goal Seek to calculate yield to maturity (You can create extra column if needed).
    'Place a button on the BondData sheet in order to run this subroutine.
    BondData.Activate
    Dim i As Integer
    Dim j As Integer
    
    j = Range("InputData").Columns.Count
    
    Dim n As Integer
    n = Range("InputData").Rows.Count
    
    For i = 1 To n
        Range("A1").Offset(i, j).GoalSeek Goal:=Range("A1").Offset(i, j - 1).Value, ChangingCell:=Range("A1").Offset(i, j + 1)
    Next
   
End Sub



