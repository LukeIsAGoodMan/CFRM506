Attribute VB_Name = "BondAnalyticsLib"
Option Explicit

Enum Ccy        'Public by default
    USD = 100
    JPY = 101
    GBP = 102
End Enum

Type Bond       'Public by default
    ID As Integer               'Serial integer, unique (PK)
    CouponRate As Double
    CouponFreq As Double             'Restricted to 3, 6, or 12 months
    SettlementDate As Date
    ExpiryDate As Date
    Currency As Ccy             'Enumeration
    FaceValue As Double
    MktRate As Double           'Const mkt int rate
    MktValue As Double          'Price of bond in the market
    PresentValue As Double      'Calculated field
    YieldToMaturity As Double   'Calculated field

End Type

Public Sub CalcPresentValue(myBond As Bond)
'Calculate present value for the bond passed as argument in this sub and assign it to member variable PresentValue of myBond
    
    Dim par As Double
    
    par = Excel.Application.WorksheetFunction.Price(myBond.SettlementDate, myBond.ExpiryDate, myBond.CouponRate, myBond.MktRate, 100, 12 / myBond.CouponFreq)
    myBond.PresentValue = par * myBond.FaceValue / 100
    
End Sub


