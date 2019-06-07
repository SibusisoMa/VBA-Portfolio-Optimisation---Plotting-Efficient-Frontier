
Sub Plot_EfficientFrontier()


Dim Risk_Interval As Double
Dim iPt_Variance As Double
Dim minVar_Risk As Double
Dim maxVar_Risk As Double
Dim nPts As Integer
Dim iPt As Integer
Dim St_Dev As Double
Dim Expected_Return As Double
Dim nAssets As Integer
Dim iAsset As Integer
Dim sharpe_ratio As Double

Sheets("Data").Activate

'calculate the minimum variance portfolio
    solverreset
    SolverOk SetCell:="$B$5", MaxMinVal:=2, ValueOf:=0, ByChange:="$C$2:$K$2", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverAdd CellRef:="$C$2:$K$2", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$G$2", Relation:=2, FormulaText:="0"
    SolverAdd CellRef:="$M$2", Relation:=2, FormulaText:="1"
    SolverSolve True

'save variance of minVar portfolio
    minVar_Risk = Sheets("Data").Range("B5").Value

'calculate the maximum return portfolio
    solverreset
    SolverOk SetCell:="$B$6", MaxMinVal:=1, ValueOf:=0, ByChange:="$C$2:$K$2", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverAdd CellRef:="$C$2:$K$2", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$G$2", Relation:=2, FormulaText:="0"
    SolverAdd CellRef:="$M$2", Relation:=2, FormulaText:="1"
    SolverSolve True

'save variance of maxReturn portfolio
    maxVar_Risk = Sheets("Data").Range("B5").Value

'calculate the variance intervals
nPts = Sheets("Data").Range("B8")
Risk_Interval = (maxVar_Risk - minVar_Risk) / (nPts - 1)

nAssets = Sheets("Data").Range("B10")

' calculate the portfolios along the efficient frontiesheer
For iPt = 1 To nPts
    iPt_Variance = minVar_Risk + (iPt - 1) * Risk_Interval
    Sheets("Data").Range("B9") = iPt_Variance
    
    
    solverreset
    SolverOk SetCell:="$B$6", MaxMinVal:=1, ValueOf:=0, ByChange:="$C$2:$K$2", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverAdd CellRef:="$C$2:$K$2", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$M$2", Relation:=2, FormulaText:="1"
    SolverAdd CellRef:="$B$5", Relation:=2, FormulaText:="$B$9"

    SolverSolve True
    
    St_Dev = Range("B4")
    Expected_Return = Range("B6")
    sharpe_ratio = Range("E8")
    Set opt_weights = Range("C2:K2")
    
    Sheets("dEff_Frontier").Range("A6").Offset(iPt - 1, 0) = iPt
    Sheets("dEff_Frontier").Range("B6").Offset(iPt - 1, 0) = iPt_Variance
    Sheets("dEff_Frontier").Range("C6").Offset(iPt - 1, 0) = St_Dev
    Sheets("dEff_Frontier").Range("D6").Offset(iPt - 1, 0) = Expected_Return
    Sheets("dEff_Frontier").Range("N6").Offset(iPt - 1, 0) = sharpe_ratio
    
    For iAsset = 1 To nAssets
        Sheets("dEff_Frontier").Range("E6").Offset(iPt - 1, iAsset - 1) = Sheets("Data").Range("C2").Offset(0, iAsset - 1)
    
    Next iAsset
    
   
Next iPt

'calculate optimal portfolio
    solverreset
    SolverOk SetCell:="$E$8", MaxMinVal:=1, ValueOf:=0, ByChange:="$C$2:$K$2", _
        Engine:=1, EngineDesc:="GRG Nonlinear"
    SolverAdd CellRef:="$C$2:$K$2", Relation:=3, FormulaText:="0"
    SolverAdd CellRef:="$M$2", Relation:=2, FormulaText:="1"
    SolverAdd CellRef:="$G$2", Relation:=2, FormulaText:="0"
    SolverSolve True
    
 'save max sharpe ratio portfolio
  max_Sharpe = Sheets("Data").Range("E8").Value

    St_Dev = Range("B4")
    Expected_Return = Range("B6")
    sharpe_ratio = Range("E8")
    Variance = Range("B9")
    opt_weights = Range("C2:K2")
    
    Sheets("dEff_Frontier").Range("B2") = Variance
    Sheets("dEff_Frontier").Range("C2") = St_Dev
    Sheets("dEff_Frontier").Range("D2") = Expected_Return
    Sheets("dEff_Frontier").Range("N2") = sharpe_ratio
    Sheets("dEff_Frontier").Range("E2:M2") = opt_weights

    
End Sub

