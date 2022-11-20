Attribute VB_Name = "Module1"
Option Explicit

Public Sub Tracker(EV, SW, Max, Min, Total)
    SolverReset
    SolverOK setCell:=EV, MaxMinVal:=2, ByChange:=SW
    SolverAdd CellRef:=SW, Relation:=1, FormulaText:=Min
    SolverAdd CellRef:=SW, Relation:=3, FormulaText:=Max
    SolverAdd CellRef:=Total, Relation:=2, FormulaText:=1
    SolverSolve userFinish:=False
    SolverSave SaveArea:=SW
End Sub


Sub Track1()
    
    Call Tracker("ErrVariance1", "StyleWeights1", "MinWeights1", "MaxWeights1", "Total1")
End Sub


Sub Track2()
    
    Call Tracker("ErrVariance2", "StyleWeights2", "MinWeights2", "MaxWeights2", "Total2")
End Sub
Sub Track()
    Call Tracker("ErrVariance", "StyleWeights", "MinWeights", "MaxWeights", "Total")
End Sub
    
