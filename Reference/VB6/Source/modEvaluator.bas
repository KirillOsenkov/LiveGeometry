Attribute VB_Name = "modEvaluator"
Option Explicit

Public WasThereAnErrorEvaluatingLastExpression As EvalErrorType
Public DiffVariable As String

Private Const PI = 3.14159265358979
Private Const E = 2.71828182845905
Private Const Ln10 = 2.30258509299405
Private Const Ln2 = 0.693147180559945
Private Const GoldenSection = 0.618033988749895
Private Const DegreeSign = "°"
Private Const PiSign = "¶"

Private Const ToDeg = 180 / PI
Private Const ToRad = PI / 180

Public Const ErraticBranch As Long = -1

Public Const unMinus = "-"
Public Const unLogicalNegation = "NOT "

Public Const biAddition = "+"
Public Const biSubtraction = "-"
Public Const biMultiplication = "*"
Public Const biDivision = "/"
Public Const biIntegerDivision = "\"
Public Const biModulus = "%"
Public Const biPower = "^"
Public Const biLogicalAnd = " AND "
Public Const biLogicalOr = " OR "
Public Const biLogicalXor = " XOR "
Public Const biLogicalImp = " IMP "
Public Const biLogicalEqv = " EQV "
Public Const biLogicalScheffer = " SHF "
Public Const biLogicalPierce = " PRC "
Public Const biComparisonEqv = "="
Public Const biComparisonNotEqv = "<>"
Public Const biComparisonMore = ">"
Public Const biComparisonLess = "<"
Public Const biComparisonMoreEqv = ">="
Public Const biComparisonLessEqv = "<="

Public Const fnSin = "SIN"
Public Const fnCos = "COS"
Public Const fnTg = "TG"
Public Const fnTan = "TAN"
Public Const fnCtg = "CTG"
Public Const fnCot = "COT"

Public Const fnArcsin = "ARCSIN"
Public Const fnArccos = "ARCCOS"
Public Const fnAsin = "ASIN"
Public Const fnAcos = "ACOS"
Public Const fnAtn = "ATN"
Public Const fnArctg = "ARCTG"
Public Const fnArctan = "ARCTAN"
Public Const fnAtan = "ATAN"
Public Const fnArcctg = "ARCCTG"
Public Const fnArccot = "ARCCOT"
Public Const fnAcot = "ACOT"

Public Const fnSinH = "SINH"
Public Const fnCosH = "COSH"
Public Const fnTgH = "TANH"
Public Const fnCtgH = "COTH"
Public Const fnAreasin = "AREASIN"
Public Const fnAreacos = "AREACOS"
Public Const fnAreatg = "AREATAN"
Public Const fnAreactg = "AREACOT"
Public Const fnArcsinH = "ARCSINH"
Public Const fnArccosH = "ARCCOSH"
Public Const fnArctanH = "ARCTANH"
Public Const fnArccotH = "ARCCOTH"

Public Const fnLn = "LN"
Public Const fnLg = "LG"
Public Const fnLog = "LOG"
Public Const fnExp = "EXP"
Public Const fnSqr = "SQR"
Public Const fnSqrt = "SQRT"
Public Const fnInt = "INT"
Public Const fnRound = "ROUND"
Public Const fnRandom = "RANDOM"
Public Const fnRnd = "RND"
Public Const fnAbs = "ABS"
Public Const fnSgn = "SGN"
Public Const fnSign = "SIGN"
Public Const fnFactorial = "FACTORIAL"
Public Const fnFact = "FACT"
Public Const fnComb = "COMB"

Public Const fnMax = "MAX"
Public Const fnMin = "MIN"
Public Const fnIf = "IF"

Public Const fnToDeg = "TODEG"
Public Const fnToRad = "TORAD"
Public Const fnDeg = "DEG"
Public Const fnRad = "RAD"
Public Const fnDistance = "DISTANCE"
Public Const fnDist = "DIST"
Public Const fnAngle = "ANGLE"
Public Const fnAng = "ANG"
Public Const fnOAngle = "OANGLE"
Public Const fnOAng = "OANG"
Public Const fnArea = "AREA"
Public Const fnArg = "ARG"
Public Const fnNorm = "NORM"
Public Const fnXAng = "XANG"
Public Const fnXAngle = "XANGLE"
Public Const fnGetX = ".X"
Public Const fnGetY = ".Y"
Public Const fnGetR = ".R"
Public Const fnGetA = ".A"

Public Enum UnaryOperationType
    uMinus
    uLogicalNegation
End Enum

Public Enum BinaryOperationType
    bAddition
    bSubtraction
    bMultiplication
    bDivision
    bIntegerDivision
    bModulus
    bPower
    bLogicalAnd
    bLogicalOr
    bLogicalXor
    bLogicalImp
    bLogicalEqv
    bLogicalSch
    bLogicalPrc
    bComparisonEqv
    bComparisonNotEqv
    bComparisonLess
    bComparisonLessEqv
    bComparisonMore
    bComparisonMoreEqv
End Enum

Public Enum FunctionType
    fSin
    fCos
    fTg
    fTan
    fCtg
    fCot
    fArcSin
    fArccos
    fArctan
    fAsin
    fAcos
    fAtn
    fAtan
    fArctg
    fArcctg
    fArccot
    fAcot
    fSinH
    fCosH
    fTgH
    fCtgH
    fAreaSin
    fAreaCos
    fAreaTg
    fAreaCtg
    fLn
    fLg
    fLog
    fExp
    fSqr
    fInt
    fRound
    fRnd
    fAbs
    fSgn
    fRandom
    fToDeg
    fToRad
    fDeg
    fRad
    fFactorial
    fCombin
    fMax
    fMin
    fIf
    
    fDistance
    fAngle
    fOAngle
    fArea
    fArg
    fNorm
    fXAng
    
    fMinus
    fLogicalNegation
    
    fGetX
    fGetY
    
    fUser
End Enum

Public Enum BType
    bConstant
    bVariable
    bOperatorUnary
    bOperatorBinary
    bFunction
    bWatch
End Enum

Public Type Branch
    BranchType As BType
    CurrentValue As Double
    VariableValue As Double
    VariableName As String
    UnaryOpType As UnaryOperationType
    BinaryOpType As BinaryOperationType
    FuncType As FunctionType
    BranchReference1 As Long
    BranchReference2 As Long
    BranchReference3 As Long
    Points() As Long
    NumberOfPoints As Long
End Type

Public Enum EvalErrorType
    eetEverythingOK
    eetZeroLengthExpression
    eetDivisionByZero
    eetNegativeRoot
    eetNegativeRationalPower
    eetArcFunctionOutOfBounds
    eetNonPositiveLogarithm
    eetPointNotFound
    eetWrongFactorialOperand
    eetInvalidParentheses
    eetExpressionNotFinished
    eetParameterNotOptional
    eetUnrecognizedExpression
End Enum

Public Type CustomVariable
    Name As String
    CurrentValue As Double
End Type

Public Type Tree
    Branches() As Branch
    BranchCount As Long
    Erroneous As Boolean
    Error As EvalErrorType
    Points() As Long
    WEs() As Long
    NumberOfPoints As Long
    NumberOfWEs As Long
End Type

Public CVCount As Long
Public CustomVariables() As CustomVariable
Public Trees() As Tree
Public SubstitutionOn As Boolean
Public SubstitutionFigure As Long

Public Function Evaluate(ByVal Expression As String) As Double
Dim MainTree As Tree

If Expression = "" Then Evaluate = 0: Exit Function
If IsNumeric(Expression) Then Evaluate = CDbl(Expression): Exit Function
If Val(Expression) <> 0 Then Evaluate = Val(Expression): Exit Function

MainTree = BuildTree(Expression)
If MainTree.Erroneous Then Evaluate = EmptyVar: Exit Function

Evaluate = RecalculateTree(MainTree, 1)
End Function

Public Function BuildTree(ByVal S As String, Optional ByVal ResolveFromWE As Boolean = False, Optional ByVal EvaluateToBoolean As Boolean = True, Optional ByVal ShouldEval As Boolean = True) As Tree
On Error GoTo EH
Dim MainTree As Tree
ReDim MainTree.Branches(1 To 1)
WasThereAnErrorEvaluatingLastExpression = eetEverythingOK

S = Replace(S, "[", "(")
S = Replace(S, "]", ")")
BuildBranch UCase(S), MainTree, ResolveFromWE, EvaluateToBoolean
If WasThereAnErrorEvaluatingLastExpression <> eetEverythingOK Then
    MainTree.Erroneous = True
    MainTree.Error = WasThereAnErrorEvaluatingLastExpression
Else
    MainTree.Erroneous = False
    MainTree.Error = eetEverythingOK
    If ShouldEval Then RecalculateTree MainTree
End If

BuildTree = MainTree
Exit Function

EH:
End Function

Public Function BuildBranch(ByVal S As String, MainTree As Tree, Optional ByVal ResolveFromWE As Boolean = False, Optional ByVal EvaluateToBoolean As Boolean = True) As Long
Dim ParenthesesBalance As Long, Z As Long, Q As Long, Priority As Long, CB As Long, LenS As Long
Dim Op1 As String, Op2 As String
Dim MidStr As String, MidStr2 As String, MidStr3 As String
Dim Ref1 As Long, Ref2 As Long, Ref3 As Long
Dim A As String, fType As FunctionType
Dim TN As Long, tS As String, LowBoundPriority As Long

S = Trim(S)
If S = "" Then ErrorEvaluating eetZeroLengthExpression: BuildBranch = ErraticBranch: Exit Function

BinaryOpSearch:
ParenthesesBalance = 0
LenS = Len(S)
LowBoundPriority = IIf(EvaluateToBoolean, -3, 1)

For Priority = LowBoundPriority To 3
    Op1 = ""
    Op2 = ""
    For Z = LenS To 1 Step -1
        MidStr = Mid(S, Z, 1)
        If MidStr = ")" Then ParenthesesBalance = ParenthesesBalance + 1
        If MidStr = "(" Then ParenthesesBalance = ParenthesesBalance - 1
        If ParenthesesBalance < 0 Then
            ErrorEvaluating eetInvalidParentheses
            BuildBranch = ErraticBranch
            Exit Function
        End If
        
        If ParenthesesBalance = 0 Then
            If Z > 4 And Priority = -3 Then
                If UCase(Mid(S, Z - 4, 5)) = biLogicalPierce Then
                    Op1 = Trim(Left(S, Z - 5))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalPrc, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If UCase(Mid(S, Z - 4, 5)) = biLogicalScheffer Then
                    Op1 = Trim(Left(S, Z - 5))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalSch, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If UCase(Mid(S, Z - 4, 5)) = biLogicalImp Then
                    Op1 = Trim(Left(S, Z - 5))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalImp, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If UCase(Mid(S, Z - 4, 5)) = biLogicalEqv Then
                    Op1 = Trim(Left(S, Z - 5))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalEqv, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If UCase(Mid(S, Z - 4, 5)) = biLogicalXor Then
                    Op1 = Trim(Left(S, Z - 5))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalXor, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
            End If
            
            If Z > 3 And Priority = -2 Then
                If UCase(Mid(S, Z - 3, 4)) = biLogicalOr Then
                    Op1 = Trim(Left(S, Z - 4))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalOr, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
            End If
            
            If Z > 4 And Priority = -1 Then
                If UCase(Mid(S, Z - 4, 5)) = biLogicalAnd Then
                    Op1 = Trim(Left(S, Z - 5))
                    CB = AddBranch(MainTree, bOperatorBinary, bLogicalAnd, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
            End If
                
            If Z > 1 And Priority = 0 Then
                If Mid(S, Z - 1, 2) = biComparisonLessEqv Then
                    Op1 = Trim(Left(S, Z - 2))
                    CB = AddBranch(MainTree, bOperatorBinary, bComparisonLessEqv, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If Mid(S, Z - 1, 2) = biComparisonMoreEqv Then
                    Op1 = Trim(Left(S, Z - 2))
                    CB = AddBranch(MainTree, bOperatorBinary, bComparisonMoreEqv, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If Mid(S, Z - 1, 2) = biComparisonNotEqv Then
                    Op1 = Trim(Left(S, Z - 2))
                    CB = AddBranch(MainTree, bOperatorBinary, bComparisonNotEqv, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If MidStr = biComparisonEqv Then
                    Op1 = Trim(Left(S, Z - 1))
                    CB = AddBranch(MainTree, bOperatorBinary, bComparisonEqv, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If MidStr = biComparisonMore Then
                    Op1 = Trim(Left(S, Z - 1))
                    CB = AddBranch(MainTree, bOperatorBinary, bComparisonMore, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
                If MidStr = biComparisonLess Then
                    Op1 = Trim(Left(S, Z - 1))
                    CB = AddBranch(MainTree, bOperatorBinary, bComparisonLess, , Op1, Op2, ResolveFromWE)
                    Ref1 = MainTree.Branches(CB).BranchReference1
                    Ref2 = MainTree.Branches(CB).BranchReference2
                    If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                    BuildBranch = CB
                    Exit Function
                End If
            End If
            
            If Priority > 0 Then
                Op1 = RTrim(Left(S, Z - 1))
                If Op1 <> "" Then
                    Select Case MidStr
                        Case biAddition
                            If Priority = 1 Then
                                CB = AddBranch(MainTree, bOperatorBinary, bAddition, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                        Case biSubtraction
                            If Priority = 1 And Mid(S, Z - 1, 1) <> unMinus Then
                                CB = AddBranch(MainTree, bOperatorBinary, bSubtraction, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                        Case biMultiplication
                            If Priority = 2 Then
                                CB = AddBranch(MainTree, bOperatorBinary, bMultiplication, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                        Case biDivision
                            If Priority = 2 Then
                                CB = AddBranch(MainTree, bOperatorBinary, bDivision, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                        Case biIntegerDivision
                            If Priority = 2 Then
                                CB = AddBranch(MainTree, bOperatorBinary, bIntegerDivision, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                        Case biModulus
                            If Priority = 2 Then
                                CB = AddBranch(MainTree, bOperatorBinary, bModulus, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                        Case biPower
                            If Priority = 3 Then
                                CB = AddBranch(MainTree, bOperatorBinary, bPower, , Op1, Op2, ResolveFromWE)
                                Ref1 = MainTree.Branches(CB).BranchReference1
                                Ref2 = MainTree.Branches(CB).BranchReference2
                                If Ref1 = ErraticBranch Or Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
                                BuildBranch = CB
                                Exit Function
                            End If
                    End Select
                End If
            End If
        End If 'parentheses balance == 0
        
        Op2 = MidStr & Op2
    Next ' for each symbol in s
    
    If ParenthesesBalance <> 0 Then
        ErrorEvaluating eetInvalidParentheses
        BuildBranch = ErraticBranch
        Exit Function
    End If
    
    If Priority = -1 Then
        If UCase(Left(S, 4)) = unLogicalNegation Then
            Op1 = Trim(Right(S, LenS - 4))
            CB = AddBranch(MainTree, bFunction, , fLogicalNegation, Op1)
            Ref1 = BuildBranch(Op1, MainTree, ResolveFromWE, True)
            If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
            MainTree.Branches(CB).BranchReference1 = Ref1
            BuildBranch = CB
            Exit Function
        End If
    End If
Next

'###############################################################################
'###############################################################################
'###############################################################################
'###############################################################################
'###############################################################################
'###############################################################################

If Left(S, 1) = unMinus Then
    S = LTrim(Right(S, Len(S) - 1))
    CB = AddBranch(MainTree, bFunction, , fMinus, S)
    Ref1 = BuildBranch(S, MainTree, ResolveFromWE, EvaluateToBoolean)
    If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
    MainTree.Branches(CB).BranchReference1 = Ref1
    BuildBranch = CB
    Exit Function
End If

If Left(S, 1) = "(" And Right(S, 1) = ")" Then
    S = Mid(S, 2, Len(S) - 2)
    GoTo BinaryOpSearch
End If

If Right(S, 1) = DegreeSign Then
    CB = AddBranch(MainTree, bFunction, , fDeg)
    
    Ref1 = BuildBranch(Left(S, Len(S) - 1), MainTree, ResolveFromWE, EvaluateToBoolean)
    If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
    
    MainTree.Branches(CB).BranchReference1 = Ref1
    
    BuildBranch = CB
    Exit Function
End If

Z = InStr(S, "(")
If Z = 0 Then 'it's not a function or at least a function with no parameters
    If IsNumeric(S) Or Val(S) <> 0 Then 'let's see if it is a number
        CB = AddBranch(MainTree, bConstant)
        If IsNumeric(S) Then
            MainTree.Branches(CB).CurrentValue = CDbl(S)
        Else
            MainTree.Branches(CB).CurrentValue = Val(S)
        End If
        BuildBranch = CB
        Exit Function
    End If
    
'    Q = GetWatchExpressionByName(S)
'    If Q <> 0 Then
'        If WatchExpressions(Q).Expression <> S Then
'            If ResolveFromWE Then
'                S = WatchExpressions(Q).Expression
'                GoTo BinaryOpSearch
'            Else
'                CB = AddBranch(MainTree, bWatch)
'                MainTree.Branches(CB).VariableName = S
'                MainTree.Branches(CB).BranchReference1 = Q
'                AddWEToTree MainTree, Q
'                BuildBranch = CB
'                Exit Function
'            End If
'        End If
'    End If
    
    If S = "PI" Or S = PiSign Then
        CB = AddBranch(MainTree, bConstant)
        MainTree.Branches(CB).CurrentValue = PI
        BuildBranch = CB
        Exit Function
    End If
    
    If S = "E" Then
        CB = AddBranch(MainTree, bConstant)
        MainTree.Branches(CB).CurrentValue = E
        BuildBranch = CB
        Exit Function
    End If
    
    If S = "TRUE" Then
        CB = AddBranch(MainTree, bConstant)
        MainTree.Branches(CB).CurrentValue = True
        BuildBranch = CB
        Exit Function
    End If
    
    If S = "FALSE" Then
        CB = AddBranch(MainTree, bConstant)
        MainTree.Branches(CB).CurrentValue = False
        BuildBranch = CB
        Exit Function
    End If
    
    If S = fnToDeg Or S = fnDeg Then
        CB = AddBranch(MainTree, bConstant)
        MainTree.Branches(CB).CurrentValue = ToDeg
        BuildBranch = CB
        Exit Function
    End If
    
    If S = fnToRad Or S = fnRad Then
        CB = AddBranch(MainTree, bConstant)
        MainTree.Branches(CB).CurrentValue = ToRad
        BuildBranch = CB
        Exit Function
    End If

    If S = fnRnd Then
        CB = AddBranch(MainTree, bFunction)
        MainTree.Branches(CB).CurrentValue = Rnd
        MainTree.Branches(CB).FuncType = fRnd
        BuildBranch = CB
        Exit Function
    End If
    
    If Right(S, 2) = fnGetX Then
        Ref1 = GetPointByName(Left(S, Len(S) - 2))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        CB = AddBranch(MainTree, bFunction, , fGetX)
        MainTree.Branches(CB).BranchReference1 = Ref1
        AddPointToTree MainTree, Ref1
        BuildBranch = CB
        Exit Function
    End If
    
    If Right(S, 2) = fnGetY Then
        Ref1 = GetPointByName(Left(S, Len(S) - 2))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        CB = AddBranch(MainTree, bFunction, , fGetY)
        MainTree.Branches(CB).BranchReference1 = Ref1
        AddPointToTree MainTree, Ref1
        BuildBranch = CB
        Exit Function
    End If

    If Right(S, 2) = fnGetR Then
        Ref1 = GetPointByName(Left(S, Len(S) - 2))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        CB = AddBranch(MainTree, bFunction, , fNorm)
        MainTree.Branches(CB).BranchReference1 = Ref1
        AddPointToTree MainTree, Ref1
        BuildBranch = CB
        Exit Function
    End If
    
    If Right(S, 2) = fnGetA Then
        Ref1 = GetPointByName(Left(S, Len(S) - 2))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        CB = AddBranch(MainTree, bFunction, , fArg)
        MainTree.Branches(CB).BranchReference1 = Ref1
        AddPointToTree MainTree, Ref1
        BuildBranch = CB
        Exit Function
    End If

Else              ' there is a "(" in S
    
    A = Left(S, Z - 1)
    Ref1 = 0
    Ref2 = 0
    Ref3 = 0
    
    'which function is S?
    Select Case A
    Case fnAng, fnAngle
        CB = AddBranch(MainTree, bFunction, , fAngle)
        Ref1 = GetPointByName(GetParameter(S))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        Ref2 = GetPointByName(GetParameter(S, 2))
        If Ref2 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference2 = Ref2
        Ref3 = GetPointByName(GetParameter(S, 3))
        If Ref3 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference3 = Ref3
        AddPointToTree MainTree, Ref1
        AddPointToTree MainTree, Ref2
        AddPointToTree MainTree, Ref3
        BuildBranch = CB
        Exit Function
    Case fnDist, fnDistance, ""
        CB = AddBranch(MainTree, bFunction, , fDistance)
        Ref1 = GetPointByName(GetParameter(S))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        Ref2 = GetPointByName(GetParameter(S, 2))
        If Ref2 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference2 = Ref2
        AddPointToTree MainTree, Ref1
        AddPointToTree MainTree, Ref2
        BuildBranch = CB
        Exit Function
    Case fnNorm
        CB = AddBranch(MainTree, bFunction, , fNorm)
        Ref1 = GetPointByName(GetParameter(S))
        If Not IsPoint(Ref1) Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        AddPointToTree MainTree, Ref1
        BuildBranch = CB
        Exit Function
    Case fnArg
        CB = AddBranch(MainTree, bFunction, , fArg)
        Ref1 = GetPointByName(GetParameter(S))
        If Not IsPoint(Ref1) Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        AddPointToTree MainTree, Ref1
        BuildBranch = CB
        Exit Function
    Case fnOAng, fnOAngle
        CB = AddBranch(MainTree, bFunction, , fOAngle)
        Ref1 = GetPointByName(GetParameter(S))
        If Ref1 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        Ref2 = GetPointByName(GetParameter(S, 2))
        If Ref2 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference2 = Ref2
        Ref3 = GetPointByName(GetParameter(S, 3))
        If Ref3 = 0 Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference3 = Ref3
        AddPointToTree MainTree, Ref1
        AddPointToTree MainTree, Ref2
        AddPointToTree MainTree, Ref3
        BuildBranch = CB
        Exit Function
    Case fnXAng, fnXAngle
        CB = AddBranch(MainTree, bFunction, , fXAng)
        Ref1 = GetPointByName(GetParameter(S))
        If Not IsPoint(Ref1) Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        Ref2 = GetPointByName(GetParameter(S, 2))
        If Not IsPoint(Ref2) Then ErrorEvaluating eetPointNotFound: BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference2 = Ref2
        AddPointToTree MainTree, Ref1
        AddPointToTree MainTree, Ref2
        BuildBranch = CB
        Exit Function
    Case fnArea
        CB = AddBranch(MainTree, bFunction, , fArea)
        TN = 1
        Do
            tS = GetParameter(S, TN, True)
            If tS = "" Then Exit Do
            ReDim Preserve MainTree.Branches(CB).Points(1 To TN)
            MainTree.Branches(CB).Points(TN) = GetPointByName(tS)
            If MainTree.Branches(CB).Points(TN) = 0 Then
                ErrorEvaluating eetPointNotFound
                BuildBranch = ErraticBranch: Exit Function
                Exit Function
            End If
            AddPointToTree MainTree, MainTree.Branches(CB).Points(TN)
            TN = TN + 1
        Loop Until tS = ""
        BuildBranch = CB
        Exit Function
        
    Case fnIf
        CB = AddBranch(MainTree, bFunction, , fIf)
        Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, True)
        If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        Ref2 = BuildBranch(GetParameter(S, 2), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference2 = Ref2
        Ref3 = BuildBranch(GetParameter(S, 3), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref3 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference3 = Ref3
        BuildBranch = CB
        Exit Function
    Case fnMin
        CB = AddBranch(MainTree, bFunction, , fMin)
        TN = 1
        Do
            tS = GetParameter(S, TN, True)
            If tS = "" Then Exit Do
            ReDim Preserve MainTree.Branches(CB).Points(1 To TN)
            MainTree.Branches(CB).Points(TN) = BuildBranch(tS, MainTree, ResolveFromWE, EvaluateToBoolean)
            If MainTree.Branches(CB).Points(TN) <= 0 Then
                BuildBranch = ErraticBranch
                Exit Function
            End If
            TN = TN + 1
        Loop Until tS = ""
        If TN < 3 Then
            BuildBranch = ErraticBranch
            Exit Function
        End If
        BuildBranch = CB
        Exit Function
    Case fnMax
        CB = AddBranch(MainTree, bFunction, , fMax)
        TN = 1
        Do
            tS = GetParameter(S, TN, True)
            If tS = "" Then Exit Do
            ReDim Preserve MainTree.Branches(CB).Points(1 To TN)
            MainTree.Branches(CB).Points(TN) = BuildBranch(tS, MainTree, ResolveFromWE, EvaluateToBoolean)
            If MainTree.Branches(CB).Points(TN) <= 0 Then
                BuildBranch = ErraticBranch
                Exit Function
            End If
            TN = TN + 1
        Loop Until tS = ""
        If TN < 3 Then
            BuildBranch = ErraticBranch
            Exit Function
        End If
        BuildBranch = CB
        Exit Function
        
        
    Case fnAbs
        fType = fAbs
    Case fnInt
        fType = fInt
    Case fnExp
        fType = fExp
    Case fnRandom
        fType = fRandom
        CB = AddBranch(MainTree, bFunction, , fType)
        Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        Ref2 = BuildBranch(GetParameter(S, 2), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        MainTree.Branches(CB).BranchReference2 = Ref2
        BuildBranch = CB
        Exit Function
        
    Case fnRound
        fType = fRound
        CB = AddBranch(MainTree, bFunction, , fType)
        Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        Ref2 = Val(GetParameter(S, 2, True))
        MainTree.Branches(CB).BranchReference1 = Ref1
        MainTree.Branches(CB).BranchReference2 = Ref2
        BuildBranch = CB
        Exit Function
        
    Case fnSgn, fnSign
        fType = fSgn
    Case fnSqr, fnSqrt
        fType = fSqr
    Case fnDeg, fnToDeg
        fType = fDeg
    Case fnRad, fnToRad
        fType = fRad
    Case fnLg
        fType = fLg
    Case fnLn
        fType = fLn
    Case fnLog
        If GetParameter(S, 2, True) <> "" Then
            CB = AddBranch(MainTree, bFunction, , fLog)
            Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, EvaluateToBoolean)
            If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
            Ref2 = BuildBranch(GetParameter(S, 2, True), MainTree, ResolveFromWE, EvaluateToBoolean)
            If Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        Else
            CB = AddBranch(MainTree, bFunction, , fLn)
            Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, EvaluateToBoolean)
            If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
            Ref2 = -1
        End If
        MainTree.Branches(CB).BranchReference1 = Ref1
        MainTree.Branches(CB).BranchReference2 = Ref2
        BuildBranch = CB
        Exit Function
        
    Case fnSin
        fType = fSin
    Case fnCos
        fType = fCos
    Case fnTan, fnTg
        fType = fTan
    Case fnCtg, fnCot
        fType = fCot
    
    Case fnAsin, fnArcsin
        fType = fAsin
    Case fnAcos, fnArccos
        fType = fAcos
    Case fnAtn, fnAtan, fnArctg, fnArctan
        fType = fAtn
    Case fnAcot, fnArccot, fnArcctg
        fType = fAcot
    
    Case fnSinH
        fType = fSinH
    Case fnCosH
        fType = fCosH
    Case fnTgH
        fType = fTgH
    Case fnCtgH
        fType = fCtgH
    
    Case fnAreasin, fnArcsinH
        fType = fAreaSin
    Case fnAreacos, fnArccosH
        fType = fAreaCos
    Case fnAreatg, fnArctanH
        fType = fAreaTg
    Case fnAreactg, fnArccotH
        fType = fAreaCtg
        
    Case fnComb
        fType = fCombin
        CB = AddBranch(MainTree, bFunction, , fType)
        Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        Ref2 = BuildBranch(GetParameter(S, 2), MainTree, ResolveFromWE, EvaluateToBoolean)
        If Ref2 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
        MainTree.Branches(CB).BranchReference1 = Ref1
        MainTree.Branches(CB).BranchReference2 = Ref2
        BuildBranch = CB
        Exit Function
        
    Case fnFact
        fType = fFactorial
    Case Else
        ErrorEvaluating eetUnrecognizedExpression
        BuildBranch = ErraticBranch
        Exit Function
    End Select              'function selection complete
    
    CB = AddBranch(MainTree, bFunction, , fType)
    
    Ref1 = BuildBranch(GetParameter(S), MainTree, ResolveFromWE, EvaluateToBoolean)
    If Ref1 = ErraticBranch Then BuildBranch = ErraticBranch: Exit Function
    
    MainTree.Branches(CB).BranchReference1 = Ref1
    MainTree.Branches(CB).BranchReference2 = Ref2
    MainTree.Branches(CB).BranchReference3 = Ref3
    
    BuildBranch = CB
    Exit Function
End If      'function and constant processing complete

Z = GetCustomVariable(S)
If Z <> 0 Then
    CB = AddBranch(MainTree, bVariable)
    MainTree.Branches(CB).VariableName = S
    MainTree.Branches(CB).BranchReference1 = Z
    BuildBranch = CB
    Exit Function
End If

Dim tempC As Long
For Z = 1 To PointCount
    If InStr(S, BasePoint(Z).Name) Then
        If tempC = 0 Then
            tempC = Z
        Else
            CB = AddBranch(MainTree, bFunction, , fDistance)
            MainTree.Branches(CB).BranchReference1 = tempC
            MainTree.Branches(CB).BranchReference2 = Z
            If Not BasePoint(tempC).Visible Or Not BasePoint(Z).Visible Then
                ErrorEvaluating eetPointNotFound
            End If
            BuildBranch = CB
            AddPointToTree MainTree, tempC
            AddPointToTree MainTree, Z
            Exit Function
        End If
    End If
Next

ErrorEvaluating eetUnrecognizedExpression
BuildBranch = ErraticBranch
End Function

Public Function AddBranch(MainTree As Tree, brType As BType, Optional brBinaryOpType As BinaryOperationType, Optional brFuncType As FunctionType, Optional ByVal Op1 As String, Optional ByVal Op2 As String, Optional ByVal ResolveFromWE As Boolean = False) As Long
Dim CB As Long, Ref1 As Long, Ref2 As Long

CB = MainTree.BranchCount + 1
MainTree.BranchCount = CB
ReDim Preserve MainTree.Branches(1 To CB)
MainTree.Branches(CB).BranchType = brType

If brType = bOperatorBinary Then
    MainTree.Branches(CB).BinaryOpType = brBinaryOpType
    If Op1 = "" Or Op2 = "" Then
        ErrorEvaluating eetExpressionNotFinished
        MainTree.Branches(CB).BranchReference1 = ErraticBranch
        MainTree.Branches(CB).BranchReference2 = ErraticBranch
    Else
        MainTree.Branches(CB).BranchReference1 = BuildBranch(Op1, MainTree, ResolveFromWE, IIf(brBinaryOpType >= bLogicalAnd And brBinaryOpType <= bLogicalPrc, True, False))
        MainTree.Branches(CB).BranchReference2 = BuildBranch(Op2, MainTree, ResolveFromWE, IIf(brBinaryOpType >= bLogicalAnd And brBinaryOpType <= bLogicalPrc, True, False))
    End If
ElseIf brType = bFunction Then
    MainTree.Branches(CB).FuncType = brFuncType
End If

AddBranch = CB
End Function

Public Function GetParameter(ByVal Expression As String, Optional ByVal ParamNum As Long = 1, Optional ByVal IsOptional As Boolean = False) As String
On Local Error Resume Next
Dim ParBalance As Long, Z As Long, S As String, Curr As Long, Responce As String
If Right(Expression, 1) = ")" Then Expression = Left(Expression, Len(Expression) - 1)
Expression = Right(Expression, Len(Expression) - InStr(Expression, "("))

Curr = 1
For Z = 1 To Len(Expression)
    S = Mid(Expression, Z, 1)
    Select Case S
        Case "("
            ParBalance = ParBalance + 1
        Case ")"
            ParBalance = ParBalance - 1
        Case ","
            If ParBalance = 0 Then Curr = Curr + 1
    End Select
    If Curr = ParamNum And (ParBalance > 0 Or S <> ",") Then Responce = Responce & S
    If Curr > ParamNum Then Exit For
Next
If Responce = "" And Not IsOptional Then ErrorEvaluating eetParameterNotOptional
GetParameter = Trim(Responce)
End Function

Public Function GetVariable(ByVal S As String) As Double
GetVariable = 0
End Function

Public Function AddPointToTree(MainTree As Tree, ByVal Point1 As Long)
Dim Z As Long
With MainTree
    For Z = 1 To .NumberOfPoints
        If .Points(Z) = Point1 Then Exit Function
    Next
    .NumberOfPoints = .NumberOfPoints + 1
    ReDim Preserve .Points(1 To .NumberOfPoints)
    .Points(.NumberOfPoints) = Point1
End With
End Function

Public Function AddWEToTree(MainTree As Tree, ByVal WE1 As Long)
Dim Z As Long
With MainTree
    For Z = 1 To .NumberOfWEs
        If .WEs(Z) = WE1 Then Exit Function
    Next
    .NumberOfWEs = .NumberOfWEs + 1
    ReDim Preserve .WEs(1 To .NumberOfWEs)
    .WEs(.NumberOfWEs) = WE1
End With
End Function

Public Function RecalculateTree(MainTree As Tree, Optional ByVal Node As Long = 1) As Double
On Local Error Resume Next

WasThereAnErrorEvaluatingLastExpression = eetEverythingOK
Dim T As Double, T2 As Double, T3 As Double, Z As Long

Select Case MainTree.Branches(Node).BranchType
    Case bFunction
        Select Case MainTree.Branches(Node).FuncType
            Case fDistance
                T = MainTree.Branches(Node).BranchReference1
                T2 = MainTree.Branches(Node).BranchReference2
                If Not IsPoint(T) Or Not IsPoint(T2) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                If Not BasePoint(T).Visible Or Not BasePoint(T2).Visible Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = Distance(BasePoint(T).X, BasePoint(T).Y, BasePoint(T2).X, BasePoint(T2).Y)
            Case fAngle
                T = MainTree.Branches(Node).BranchReference1
                T2 = MainTree.Branches(Node).BranchReference2
                T3 = MainTree.Branches(Node).BranchReference3
                If Not IsPoint(T) Or Not IsPoint(T2) Or Not IsPoint(T3) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                If Not BasePoint(T).Visible Or Not BasePoint(T2).Visible Or Not BasePoint(T3).Visible Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = Angle(BasePoint(T).X, BasePoint(T).Y, BasePoint(T2).X, BasePoint(T2).Y, BasePoint(T3).X, BasePoint(T3).Y)
            Case fGetX
                T = MainTree.Branches(Node).BranchReference1
                If Not IsPoint(T) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = BasePoint(T).X
            Case fGetY
                T = MainTree.Branches(Node).BranchReference1
                If Not IsPoint(T) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = BasePoint(T).Y
            Case fMax
                For Z = 1 To UBound(MainTree.Branches(Node).Points)
                    T = RecalculateTree(MainTree, MainTree.Branches(Node).Points(Z))
                Next
                T = MaxValue(MainTree, MainTree.Branches(Node).Points)
            Case fMin
                For Z = 1 To UBound(MainTree.Branches(Node).Points)
                    T = RecalculateTree(MainTree, MainTree.Branches(Node).Points(Z))
                Next
                T = MinValue(MainTree, MainTree.Branches(Node).Points)
            Case fIf
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T Then
                    T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                Else
                    T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference3)
                End If
            Case fSqr
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T < 0 Then ErrorEvaluating eetNegativeRoot: RecalculateTree = 0: Exit Function
                T = Sqr(T)
            Case fToDeg, fDeg
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = T * ToDeg
            Case fToRad, fRad
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = T * ToRad
            Case fOAngle
                T = MainTree.Branches(Node).BranchReference1
                T2 = MainTree.Branches(Node).BranchReference2
                T3 = MainTree.Branches(Node).BranchReference3
                If Not IsPoint(T) Or Not IsPoint(T2) Or Not IsPoint(T3) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                If Not BasePoint(T).Visible Or Not BasePoint(T2).Visible Or Not BasePoint(T3).Visible Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = OAngle(BasePoint(T).X, BasePoint(T).Y, BasePoint(T2).X, BasePoint(T2).Y, BasePoint(T3).X, BasePoint(T3).Y)
            Case fArg
                T = MainTree.Branches(Node).BranchReference1
                If Not IsPoint(T) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = 2 * PI - GetAngle(0, 0, BasePoint(T).X, BasePoint(T).Y)
            Case fXAng
                T = MainTree.Branches(Node).BranchReference1
                T2 = MainTree.Branches(Node).BranchReference2
                If Not IsPoint(T) Or Not IsPoint(T2) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = 2 * PI - GetAngle(BasePoint(T).X, BasePoint(T).Y, BasePoint(T2).X, BasePoint(T2).Y)
            Case fNorm
                T = MainTree.Branches(Node).BranchReference1
                If Not IsPoint(T) Then ErrorEvaluating eetPointNotFound: RecalculateTree = 0: Exit Function
                T = Distance(0, 0, BasePoint(T).X, BasePoint(T).Y)
            Case fArea
                T = GetPolygonArea(MainTree.Branches(Node).Points)
            
            Case fRnd
                T = Rnd
            Case fAbs
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Abs(T)
            Case fSin
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Sin(T)
            Case fCos
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Cos(T)
            Case fTg, fTan
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T2 = Cos(T)
                If T2 = 0 Then ErrorEvaluating eetDivisionByZero: RecalculateTree = 0: Exit Function
                T = Sin(T) / T2
            Case fCot, fCtg
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T2 = Sin(T)
                If T2 = 0 Then ErrorEvaluating eetDivisionByZero: RecalculateTree = 0: Exit Function
                T = Cos(T) / T2
            Case fArcSin, fAsin
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T < -1 Or T > 1 Then ErrorEvaluating eetArcFunctionOutOfBounds: RecalculateTree = 0: Exit Function
                T = Arcsin(T)
            Case fArccos, fAcos
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T < -1 Or T > 1 Then ErrorEvaluating eetArcFunctionOutOfBounds: RecalculateTree = 0: Exit Function
                T = Arccos(T)
            Case fArctg, fArctan, fAtan, fAtn
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Atn(T)
            Case fArcctg, fAcot, fArccot
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Arcctg(T)
            
            Case fSinH
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = SinH(T)
            Case fCosH
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = CosH(T)
            Case fTgH
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = TanH(T)
            Case fCtgH
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = CotH(T)
            Case fAreaSin
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = ArcsinH(T)
            Case fAreaCos
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = ArccosH(T)
            Case fAreaTg
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = ArctanH(T)
            Case fAreaCtg
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = ArccotH(T)
                
            Case fMinus
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = -T
            Case fExp
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T < 100 Then T = Exp(T)
            Case fInt
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Int(T)
            Case fLg
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T <= 0 Then ErrorEvaluating eetNonPositiveLogarithm: RecalculateTree = 0: Exit Function
                T = Log(T) / Ln10
            Case fLn
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If T <= 0 Then ErrorEvaluating eetNonPositiveLogarithm: RecalculateTree = 0: Exit Function
                T = Log(T)
            Case fRandom
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Random(T, RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2))
                'T = Random(T, MainTree.Branches(MainTree.Branches(Node).BranchReference2).CurrentValue)
            Case fRound
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Round(T, MainTree.Branches(Node).BranchReference2)
            Case fSgn
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Sgn(T)
            Case fFactorial
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Factorial(T)
            Case fCombin
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T2 = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                T = Combin(T, T2)
            Case fLog
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                If MainTree.Branches(Node).BranchReference2 = -1 Then
                    If T <= 0 Then ErrorEvaluating eetNonPositiveLogarithm: RecalculateTree = 0: Exit Function
                    T = Log(T)
                Else
                    T2 = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                    If T2 = 1 Or T2 <= 0 Or T <= 0 Then ErrorEvaluating eetNonPositiveLogarithm: RecalculateTree = 0: Exit Function
                    T = Log(T) / Log(T2)
                End If
            Case fLogicalNegation
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T = Not T
            Case Else
                T = 0
        End Select
        MainTree.Branches(Node).CurrentValue = T
        
    Case bOperatorBinary
        Select Case MainTree.Branches(Node).BinaryOpType
            Case bAddition
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) + RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bSubtraction
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) - RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bMultiplication
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) * RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bDivision
                T2 = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                If T2 = 0 Then ErrorEvaluating eetDivisionByZero: RecalculateTree = 0: Exit Function
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) / T2
            Case bIntegerDivision
                T2 = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                If Int(T2) = 0 Or Fix(T2) = 0 Then ErrorEvaluating eetDivisionByZero: RecalculateTree = 0: Exit Function
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) \ T2
            Case bModulus
                T2 = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                If Int(T2) = 0 Or Fix(T2) = 0 Then ErrorEvaluating eetDivisionByZero: RecalculateTree = 0: Exit Function
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) Mod T2
            Case bPower
                T = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1)
                T2 = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
                If (T < 0 And Int(T2) <> T2) Or Abs(T2) > 32 Then ErrorEvaluating eetNegativeRationalPower: RecalculateTree = 0: Exit Function
                MainTree.Branches(Node).CurrentValue = T ^ T2
            
            Case bLogicalAnd
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) And RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bLogicalOr
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) Or RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            
            Case bComparisonEqv
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bComparisonMore
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) > RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bComparisonLess
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) < RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bComparisonNotEqv
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) <> RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bComparisonLessEqv
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) <= RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bComparisonMoreEqv
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) >= RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            
            Case bLogicalXor
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) Xor RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bLogicalImp
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) Imp RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bLogicalEqv
                MainTree.Branches(Node).CurrentValue = RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) Eqv RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2)
            Case bLogicalSch
                MainTree.Branches(Node).CurrentValue = Not (RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) And RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2))
            Case bLogicalPrc
                MainTree.Branches(Node).CurrentValue = Not (RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference1) Or RecalculateTree(MainTree, MainTree.Branches(Node).BranchReference2))
        End Select
    Case bVariable
        T = MainTree.Branches(Node).BranchReference1
        If IsCustomVariable(T) Then
            MainTree.Branches(Node).CurrentValue = CustomVariables(T).CurrentValue
        Else
            MainTree.Branches(Node).CurrentValue = GetVariable(MainTree.Branches(Node).VariableName)
        End If
    Case bWatch
        MainTree.Branches(Node).CurrentValue = WatchExpressions(MainTree.Branches(Node).BranchReference1).Value
End Select

RecalculateTree = MainTree.Branches(Node).CurrentValue
End Function

Public Function AddCustomVariable(ByVal VarName As String, ByVal CurrentValue As Double) As Long
Dim Z As Long
For Z = 1 To CVCount
    If VarName = CustomVariables(Z).Name Then
        CustomVariables(Z).CurrentValue = CurrentValue
        AddCustomVariable = Z
        Exit Function
    End If
Next
CVCount = CVCount + 1
ReDim Preserve CustomVariables(1 To CVCount)
CustomVariables(CVCount).Name = VarName
CustomVariables(CVCount).CurrentValue = CurrentValue
AddCustomVariable = CVCount
End Function

Public Function ModifyCustomVariable(ByVal VarName As String, ByVal CurrentValue As Double) As Double
Dim Z As Long
For Z = 1 To CVCount
    If VarName = CustomVariables(Z).Name Then
        ModifyCustomVariable = CustomVariables(Z).CurrentValue
        CustomVariables(Z).CurrentValue = CurrentValue
        Exit Function
    End If
Next
End Function

Public Function GetCustomVariable(ByVal VarName As String) As Long
Dim Z As Long
For Z = 1 To CVCount
    If VarName = CustomVariables(Z).Name Then
        GetCustomVariable = Z
        Exit Function
    End If
Next
End Function

Public Function DeleteCustomVariable(ByVal VarName As String)
Dim Index As Long, Z As Long
If CVCount = 0 Then Exit Function

For Z = 1 To CVCount
    If VarName = CustomVariables(Z).Name Then
        Index = Z
        Exit For
    End If
Next Z

If Index = 0 Then Exit Function
If Index < CVCount Then
    For Z = Index To CVCount - 1
        CustomVariables(Z) = CustomVariables(Z + 1)
    Next
End If
CVCount = CVCount - 1
If CVCount > 0 Then ReDim Preserve CustomVariables(1 To CVCount)
End Function

Public Sub ErrorEvaluating(ByVal ErrorType As EvalErrorType)
WasThereAnErrorEvaluatingLastExpression = ErrorType
End Sub

Public Function IsSevere(ByVal ErrorType As EvalErrorType) As Boolean
If ErrorType = eetExpressionNotFinished Or ErrorType = eetInvalidParentheses Or ErrorType = eetZeroLengthExpression Or ErrorType = eetPointNotFound Or ErrorType = eetParameterNotOptional Or ErrorType = eetUnrecognizedExpression Then IsSevere = True Else IsSevere = False
End Function

Public Function RestoreExpressionFromTree(tTree As Tree, Optional ByVal SubstituteNumbers As Boolean = False) As String
RestoreExpressionFromTree = RestoreExpressionFromBranch(tTree, 1, SubstituteNumbers)
End Function

Public Function RestoreExpressionFromBranch(tTree As Tree, tBranch As Long, Optional ByVal SubstituteNumbers As Boolean = False) As String
Dim FName As String, SubParam As String, Z As Long, S As String

With tTree.Branches(tBranch)
    Select Case .BranchType
       Case bConstant
            RestoreExpressionFromBranch = .CurrentValue
       Case bVariable
            If IsCustomVariable(.BranchReference1) Then
                RestoreExpressionFromBranch = CustomVariables(.BranchReference1).Name
            Else
                RestoreExpressionFromBranch = .VariableName
            End If
       
       Case bFunction
            Select Case .FuncType
                Case fAbs
                    FName = fnAbs
                Case fSin
                    FName = fnSin
                Case fCos
                    FName = fnCos
                Case fTg, fTan
                    FName = fnTg
                Case fCot, fCtg
                    FName = fnCtg
                Case fArcSin, fAsin
                    FName = fnArcsin
                Case fArccos, fAcos
                    FName = fnArccos
                Case fArctg, fArctan, fAtan, fAtn
                    FName = fnArctg
                Case fArcctg, fAcot, fArccot
                    FName = fnArcctg
                Case fExp
                    FName = fnExp
                Case fInt
                    FName = fnInt
                Case fLg
                    FName = fnLg
                Case fLn
                    FName = fnLn
                Case fRnd
                    RestoreExpressionFromBranch = Proper(fnRnd)
                    Exit Function
                Case fRandom
                    FName = fnRandom
                    SubParam = "," & RestoreExpressionFromBranch(tTree, .BranchReference2, SubstituteNumbers)
                Case fLog
                    FName = fnLog
                    If .BranchReference2 > -1 Then
                        SubParam = "," & RestoreExpressionFromBranch(tTree, .BranchReference2, SubstituteNumbers)
                    Else
                        FName = fnLn
                    End If
                Case fRound
                    FName = fnRound
                    If .BranchReference2 > 0 Then SubParam = "," & .BranchReference2
                Case fSgn
                    FName = fnSgn
                Case fAbs
                    FName = fnAbs
                Case fSqr
                    FName = fnSqr
                Case fFactorial
                    FName = fnFact
                Case fMinus
                    FName = "-"
                Case fLogicalNegation
                    FName = "Not "
                    
                Case fSinH
                    FName = fnSinH
                Case fCosH
                    FName = fnCosH
                Case fTgH
                    FName = fnTgH
                Case fCtgH
                    FName = fnCtgH
                Case fAreaSin
                    FName = fnArcsinH
                Case fAreaCos
                    FName = fnArccosH
                Case fAreaTg
                    FName = fnArctanH
                Case fAreaCtg
                    FName = fnArccotH
                    
                Case fCombin
                    FName = fnComb
                    SubParam = "," & RestoreExpressionFromBranch(tTree, .BranchReference2, SubstituteNumbers)
                Case fToDeg, fDeg
                    FName = fnDeg
                Case fToRad, fRad
                    FName = fnRad
                Case fMax
                    FName = Proper(fnMax) & "("
                    For Z = 1 To UBound(.Points)
                        FName = FName & RestoreExpressionFromBranch(tTree, .Points(Z), SubstituteNumbers) & IIf(Z = UBound(.Points), ")", ",")
                    Next
                    RestoreExpressionFromBranch = FName
                    Exit Function
                Case fMin
                    FName = Proper(fnMin) & "("
                    For Z = 1 To UBound(.Points)
                        FName = FName & RestoreExpressionFromBranch(tTree, .Points(Z), SubstituteNumbers) & IIf(Z = UBound(.Points), ")", ",")
                    Next
                    RestoreExpressionFromBranch = FName
                    Exit Function
                Case fIf
                    FName = fnIf
                    SubParam = "," & RestoreExpressionFromBranch(tTree, .BranchReference2, SubstituteNumbers) & "," & RestoreExpressionFromBranch(tTree, .BranchReference3, SubstituteNumbers)
                
                Case fDistance
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = Proper(fnDist) & "($(" & .BranchReference1 & "),$(" & .BranchReference2 & "))"
                    Else
                        RestoreExpressionFromBranch = Proper(fnDist) & "(" & BasePoint(.BranchReference1).Name & "," & BasePoint(.BranchReference2).Name & ")"
                        'BasePoint(.BranchReference1).Name & BasePoint(.BranchReference2).Name
                    End If
                    Exit Function
                Case fAngle
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = Proper(fnAngle) & "($(" & .BranchReference1 & "),$(" & .BranchReference2 & "),$(" & .BranchReference3 & "))"
                    Else
                        RestoreExpressionFromBranch = Proper(fnAngle) & "(" & BasePoint(.BranchReference1).Name & "," & BasePoint(.BranchReference2).Name & "," & BasePoint(.BranchReference3).Name & ")"
                    End If
                    Exit Function
                Case fOAngle
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = Proper(fnOAngle) & "($(" & .BranchReference1 & "),$(" & .BranchReference2 & "),$(" & .BranchReference3 & "))"
                    Else
                        RestoreExpressionFromBranch = Proper(fnOAngle) & "(" & BasePoint(.BranchReference1).Name & "," & BasePoint(.BranchReference2).Name & "," & BasePoint(.BranchReference3).Name & ")"
                    End If
                    Exit Function
                Case fXAng
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = Proper(fnXAngle) & "($(" & .BranchReference1 & "),$(" & .BranchReference2 & "))"
                    Else
                        RestoreExpressionFromBranch = Proper(fnXAngle) & "(" & BasePoint(.BranchReference1).Name & "," & BasePoint(.BranchReference2).Name & ")"
                    End If
                    Exit Function
                Case fArea
                    S = Proper(fnArea) & "("
                    For Z = 1 To UBound(.Points)
                        If SubstituteNumbers Then
                            S = S & "$(" & .Points(Z) & IIf(Z = UBound(.Points), "))", "),")
                        Else
                            S = S & BasePoint(.Points(Z)).Name & IIf(Z = UBound(.Points), ")", ",")
                        End If
                    Next
                    RestoreExpressionFromBranch = S
                    Exit Function
                Case fNorm
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = Proper(fnNorm) & "($(" & .BranchReference1 & "))"
                    Else
                        RestoreExpressionFromBranch = Proper(fnNorm) & "(" & BasePoint(.BranchReference1).Name & ")"
                    End If
                    Exit Function
                Case fArg
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = Proper(fnArg) & "($(" & .BranchReference1 & "))"
                    Else
                        RestoreExpressionFromBranch = Proper(fnArg) & "(" & BasePoint(.BranchReference1).Name & ")"
                    End If
                    Exit Function
                Case fGetX
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = "$(" & .BranchReference1 & ").X"
                    Else
                        RestoreExpressionFromBranch = BasePoint(.BranchReference1).Name & ".X"
                    End If
                    Exit Function
                Case fGetY
                    If SubstituteNumbers Then
                        RestoreExpressionFromBranch = "$(" & .BranchReference1 & ").Y"
                    Else
                        RestoreExpressionFromBranch = BasePoint(.BranchReference1).Name & ".Y"
                    End If
                    Exit Function
                Case Else
                    ReportError "Неучтенная функция в модуле modEvaluator!!!"
                    FName = ""
            End Select
                
            RestoreExpressionFromBranch = Proper(FName) & "(" & RestoreExpressionFromBranch(tTree, .BranchReference1, SubstituteNumbers) & SubParam & ")"
       Case bOperatorUnary
            Select Case .UnaryOpType
                Case uMinus
                    RestoreExpressionFromBranch = "-(" & RestoreExpressionFromBranch(tTree, .BranchReference1, SubstituteNumbers) & ")"
                Case uLogicalNegation
                    RestoreExpressionFromBranch = unLogicalNegation & "(" & RestoreExpressionFromBranch(tTree, .BranchReference1, SubstituteNumbers) & ")"
            End Select
       Case bOperatorBinary
            Select Case .BinaryOpType
                Case bAddition
                    FName = "+"
                Case bSubtraction
                    FName = "-"
                Case bMultiplication
                    FName = "*"
                Case bDivision
                    FName = "/"
                Case bIntegerDivision
                    FName = "\"
                Case bPower
                    FName = "^"
                Case bModulus
                    FName = biModulus
                Case bLogicalAnd
                    FName = Proper(Trim(biLogicalAnd))
                Case bLogicalOr
                    FName = Proper(Trim(biLogicalOr))
                Case bLogicalXor
                    FName = Proper(Trim(biLogicalXor))
                Case bLogicalImp
                    FName = Trim(biLogicalImp)
                Case bLogicalEqv
                    FName = Trim(biLogicalEqv)
                Case bLogicalSch
                    FName = Trim(biLogicalScheffer)
                Case bLogicalPrc
                    FName = Trim(biLogicalPierce)
                Case bComparisonEqv
                    FName = biComparisonEqv
                Case bComparisonLess
                    FName = biComparisonLess
                Case bComparisonLessEqv
                    FName = biComparisonLessEqv
                Case bComparisonMore
                    FName = biComparisonMore
                Case bComparisonMoreEqv
                    FName = biComparisonMoreEqv
                Case bComparisonNotEqv
                    FName = biComparisonNotEqv
            End Select
            
            RestoreExpressionFromBranch = "(" & RestoreExpressionFromBranch(tTree, .BranchReference1, SubstituteNumbers) & ") " & FName & " (" & RestoreExpressionFromBranch(tTree, .BranchReference2, SubstituteNumbers) & ")"
        Case bWatch
            RestoreExpressionFromBranch = WatchExpressions(.BranchReference1).Name
    End Select
End With
End Function

Public Function TreeDependsOnPoint(ByVal Point1 As Long, tTree As Tree) As Boolean
Dim Z As Long
For Z = 1 To tTree.NumberOfPoints
    If tTree.Points(Z) = Point1 Then TreeDependsOnPoint = True: Exit Function
Next
For Z = 1 To tTree.BranchCount
    With tTree.Branches(Z)
        If .BranchType = bWatch Then
            If TreeDependsOnPoint(Point1, WatchExpressions(.BranchReference1).WatchTree) Then TreeDependsOnPoint = True: Exit Function
        End If
    End With
Next Z
End Function

Public Function TreeDependsOnWE(ByVal tWE As Long, tTree As Tree) As Boolean
Dim Z As Long
For Z = 1 To tTree.BranchCount
    With tTree.Branches(Z)
        If .BranchType = bWatch Then
            If .BranchReference1 = tWE Then TreeDependsOnWE = True: Exit Function
        End If
    End With
Next Z
End Function

Public Sub ReplacePointInTree(MainTree As Tree, ByVal Point1 As Long, ByVal Point2 As Long)
Dim Z As Long, Q As Long

With MainTree
    For Z = 1 To .BranchCount
        With .Branches(Z)
            If .BranchType = bFunction Then
                If .FuncType = fAngle Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                    If .BranchReference2 = Point1 Then .BranchReference2 = Point2
                    If .BranchReference3 = Point1 Then .BranchReference3 = Point2
                End If
                If .FuncType = fOAngle Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                    If .BranchReference2 = Point1 Then .BranchReference2 = Point2
                    If .BranchReference3 = Point1 Then .BranchReference3 = Point2
                End If
                If .FuncType = fDistance Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                    If .BranchReference2 = Point1 Then .BranchReference2 = Point2
                End If
                If .FuncType = fXAng Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                    If .BranchReference2 = Point1 Then .BranchReference2 = Point2
                End If
                If .FuncType = fGetX Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                End If
                If .FuncType = fGetY Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                End If
                If .FuncType = fNorm Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                End If
                If .FuncType = fArg Then
                    If .BranchReference1 = Point1 Then .BranchReference1 = Point2
                End If
                If .FuncType = fArea Then
                    For Q = 1 To UBound(.Points)
                        If .Points(Q) = Point1 Then .Points(Q) = Point2
                    Next
                End If
            End If
        End With
    Next
    For Z = 1 To .NumberOfPoints
        If .Points(Z) = Point1 Then .Points(Z) = Point2
    Next
End With
End Sub

Public Sub ReplaceWEInTree(Tree1 As Tree, ByVal OldValue As Long, ByVal NewValue As Long)
Dim Z As Long
With Tree1
    For Z = 1 To .BranchCount
        If .Branches(Z).BranchType = bWatch Then
            If .Branches(Z).BranchReference1 = OldValue Then .Branches(Z).BranchReference1 = NewValue
        End If
    Next
End With
End Sub

Public Function IsCustomVariable(ByVal Index As Long) As Boolean
If Index <= 0 Or Index > CVCount Then IsCustomVariable = False Else IsCustomVariable = True
End Function

Public Sub SubstitutePointNamesInExpression(ByRef sExpr As String, ByVal OldVal As Long, ByVal NewVal As Long)
sExpr = Replace(sExpr, "$(" & OldVal & ")", BasePoint(NewVal).Name)
End Sub

Public Function MaxValue(MainTree As Tree, Points() As Long)
Dim MV As Double, Z As Long
MV = -Infinity
For Z = 1 To UBound(Points)
    If MainTree.Branches(Points(Z)).CurrentValue > MV Then MV = MainTree.Branches(Points(Z)).CurrentValue
Next
MaxValue = MV
End Function

Public Function MinValue(MainTree As Tree, Points() As Long)
Dim MV As Double, Z As Long
MV = Infinity
For Z = 1 To UBound(Points)
    If MainTree.Branches(Points(Z)).CurrentValue < MV Then MV = MainTree.Branches(Points(Z)).CurrentValue
Next
MinValue = MV
End Function

Public Function DifferentiateTree(T As Tree) As String
DifferentiateTree = DifferentiateBranch(T, 1, "X")
End Function

Public Function DifferentiateBranch(T As Tree, ByVal B As Long, ByVal X As String) As String
Dim P1 As String, P2 As String, S As String, R1 As String, R2 As String

Select Case T.Branches(B).BranchType
Case bConstant
    DifferentiateBranch = "0"
Case bVariable
    If X = T.Branches(B).VariableName Then DifferentiateBranch = "1" Else DifferentiateBranch = "0"
Case bFunction
    P1 = DifferentiateBranch(T, T.Branches(B).BranchReference1, X)
    If P1 = "0" Then
        DifferentiateBranch = "0"
    Else
        P2 = RestoreExpressionFromBranch(T, T.Branches(B).BranchReference1, False)
        Select Case T.Branches(B).FuncType
        Case fCos
            S = "-Sin(" & P2 & ")"
            If P1 <> "1" Then S = S & " * (" & P1 & ")"
        Case fSin
            S = "Cos(" & P2 & ")"
            If P1 <> "1" Then S = S & " * (" & P1 & ")"
        End Select
    End If
Case bOperatorBinary
    Select Case T.Branches(B).BinaryOpType
    Case bAddition
    Case bSubtraction
    Case bMultiplication
    Case bDivision
    Case bPower
    End Select
Case bOperatorUnary
    If T.Branches(B).UnaryOpType = uMinus Then
        P1 = DifferentiateBranch(T, T.Branches(B).BranchReference1, X)
        If P1 = "0" Then
            DifferentiateBranch = "0"
        Else
            DifferentiateBranch = "-(" & P1 & ")"
        End If
    End If
End Select
End Function

Public Function Differentiate(ByVal S As String, ByVal X As String) As String

End Function
'
'Public Function Differentiate2(ByVal Expression As String) As String
'If Expression = "" Then Exit Function
'If DiffVariable = "" Then DiffVariable = "X"
'Dim ParenthesesBalance As Long, Z As Long, Priority As Long, CB As Long, LenS As Long
'Dim Op1 As String, Op2 As String, S As String, TempStr1 As String
'Dim MidStr As String, MidStr2 As String, MidStr3 As String, Param1 As String, Param2 As String
'Const LC_LB = 2
'Const LC_UB = 8
'Dim LeftCache(LC_LB To LC_UB) As String
'S = UCase(Trim(Expression))
'Z = InStr(S, " ")
'Do
'    If Z + 1 > Len(S) Then Exit Do
'    Do While Mid(S, Z + 1, 1) = " "
'        S = Left(S, Z + 1) & Right(S, Len(S) - Z - 1)
'    Loop
'    Z = InStr(Z + 1, S, " ")
'Loop Until Z = 0
'
'BinaryOpSearch:
'If S = DiffVariable Then Differentiate = "1": Exit Function
'
'For Priority = 1 To 3
'    ParenthesesBalance = 0
'    Op1 = ""
'    Op2 = ""
'    LenS = Len(S)
'    For Z = LenS To 1 Step -1
'        MidStr = Mid(S, Z, 1)
'        If MidStr = ")" Then ParenthesesBalance = ParenthesesBalance + 1
'        If MidStr = "(" Then ParenthesesBalance = ParenthesesBalance - 1
'        If ParenthesesBalance < 0 Then Exit Function
'        If ParenthesesBalance = 0 Then
'            If Op2 <> "" Then
'                Op1 = RTrim(Left(S, Z - 1))
'                TempStr1 = ""
'                Param1 = ""
'                Param2 = ""
'                If Op1 <> "" Then
'                    Select Case MidStr
'                        Case biAddition
'                            If Priority = 1 Then
'                                Param1 = Differentiate(Op1)
'                                Param2 = Differentiate(Op2)
'                                If Param1 = "0" Then
'                                    If Param2 = "0" Then TempStr1 = "0" Else TempStr1 = Param2
'                                Else
'                                    If Param2 = "0" Then TempStr1 = Param1 Else TempStr1 = Param1 & " + " & Param2
'                                End If
'                                Differentiate = TempStr1
'                                Exit Function
'                            End If
'                        Case biSubtraction
'                            If Priority = 1 And Mid(S, Z - 1, 1) <> unMinus Then
'                                Param1 = Differentiate(Op1)
'                                Param2 = Differentiate(Op2)
'                                If Param1 = "0" Then
'                                    If Param2 = "0" Then TempStr1 = "0" Else TempStr1 = "-" & Param2
'                                Else
'                                    If Param2 = "0" Then TempStr1 = Param1 Else TempStr1 = Param1 & " - " & Param2
'                                End If
'                                Differentiate = TempStr1
'                                Exit Function
'                            End If
'                        Case biMultiplication
'                            If Priority = 2 Then
'                                Param1 = Differentiate(Op1)
'                                Param2 = Differentiate(Op2)
'                                If Param1 <> "0" Then Param1 = IIf(Param1 = "1", "", Param1 & " * ") & Op2
'                                If Param2 <> "0" Then Param2 = IIf(Param2 = "1", "", Param2 & " * ") & Op1
'                                If Param1 = "0" Then
'                                    If Param2 = "0" Then TempStr1 = "0" Else TempStr1 = Param2
'                                Else
'                                    If Param2 = "0" Then TempStr1 = Param1 Else TempStr1 = Param1 & " + " & Param2
'                                End If
'                                Differentiate = TempStr1
'                                Exit Function
'                            End If
'                        Case biDivision
'                            If Priority = 2 Then
'                                Param1 = Differentiate(Op1)
'                                Param2 = Differentiate(Op2)
'                                If Param1 <> "0" Then Param1 = IIf(Param1 = "1", "", Param1 & " * ") & Op2
'                                If Param2 <> "0" Then Param2 = IIf(Param2 = "1", "", Param2 & " * ") & Op1
'                                If Param1 = "0" Then
'                                    If Param2 = "0" Then TempStr1 = "0" Else TempStr1 = Param2
'                                Else
'                                    If Param2 = "0" Then TempStr1 = Param1 Else TempStr1 = Param1 & " - " & Param2
'                                    TempStr1 = "(" & TempStr1 & ") / (" & Op2 & ") ^ 2"
'                                End If
'                                Differentiate = TempStr1
'                                Exit Function
'                            End If
'                        Case biPower
'                            If Priority = 3 Then
'                                If Op1 = DiffVariable Then
'                                    If InStr(Op2, DiffVariable) Then
'                                        Param2 = Differentiate("(" & Op2 & ") * " & fnLn & "(" & Op1 & ")")
'                                        TempStr1 = "e ^ ((" & Op2 & ") * " & fnLn & "(" & Op1 & "))"
'                                        If Param2 <> "0" Then TempStr1 = TempStr1 & IIf(Param2 = "1", "", " * (" & Param2 & ")") Else TempStr1 = "0"
'                                        Differentiate = TempStr1
'                                        Exit Function
'                                    Else
'                                        TempStr1 = "(" & Op2 & ") * " & DiffVariable & " ^ (" & Op2 & " - 1)"
'                                        Differentiate = TempStr1
'                                        Exit Function
'                                    End If
'                                ElseIf Op2 = DiffVariable Then
'                                    TempStr1 = S & " * " & fnLn & "(" & Op1 & ")"
'                                    Differentiate = TempStr1
'                                    Exit Function
'                                Else
'                                    Differentiate = "0"
'                                    Exit Function
'                                End If
'                            End If
'                    End Select
'                End If
'            End If
'        End If
'        Op2 = MidStr & Op2
'    Next
'Next
'
'If Left(S, 1) = unMinus Then Differentiate = "-" & Differentiate(Right(S, Len(S) - 1)): Exit Function
'
'If Left(S, 1) = "(" And Right(S, 1) = ")" Then
'    S = Mid(S, 2, Len(S) - 2)
'    GoTo BinaryOpSearch
'End If
'
'If S = "PI" Or S = "E" Or IsNumeric(S) Then Differentiate = "0": Exit Function
'
'For Z = LC_LB To LC_UB
'    LeftCache(Z) = Left(S, Z)
'Next
'
'Param1 = GetParameter(S)
'
'If LeftCache(3) = fnSin Then
'    TempStr1 = fnCos & "(" & Param1 & ")"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else If Param1 <> "1" Then TempStr1 = TempStr1 & " * " & Param1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnCos Then
'    TempStr1 = "-" & fnSin & "(" & Param1 & ")"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else If Param1 <> "1" Then TempStr1 = TempStr1 & " * " & Param1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnTan Or LeftCache(2) = fnTg Then
'    TempStr1 = " / " & fnCos & "(" & Param1 & ") ^ 2"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnCtg Or LeftCache(3) = fnCot Then
'    TempStr1 = " / " & fnSin & "(" & Param1 & ") ^ 2"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = "-" & Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnAbs Then
'    TempStr1 = fnSgn & "(" & Param1 & ")"
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnSgn Then Differentiate = "0": Exit Function
'If LeftCache(3) = fnExp Then
'    TempStr1 = S
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else If Param1 <> "1" Then TempStr1 = TempStr1 & " * " & Param1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnAtn Or LeftCache(4) = fnAtan Or LeftCache(6) = fnArctan Or LeftCache(5) = fnArctg Then
'    TempStr1 = " / (1 + (" & Param1 & ") ^ 2"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnSqr Then
'    TempStr1 = " / (2 * " & fnSqr & "(" & Param1 & "))"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(3) = fnLog Then
'    TempStr1 = " / (" & Param1 & " * " & fnLn & "(" & GetParameter(S, 2) & "))"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(2) = fnLn Then
'    TempStr1 = " / " & Param1
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(2) = fnLg Then
'    TempStr1 = " / (" & Param1 & " * " & fnLn & "(10))"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(4) = fnAsin Or LeftCache(6) = fnArcsin Then
'    TempStr1 = " / " & fnSqr & "(1 - (" & Param1 & ") ^ 2)"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(4) = fnAcos Or LeftCache(6) = fnArccos Then
'    TempStr1 = " / " & fnSqr & "(1 - (" & Param1 & ") ^ 2)"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = "-" & Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'If LeftCache(4) = fnAcot Or LeftCache(6) = fnArcctg Or LeftCache(6) = fnArccot Then
'    TempStr1 = " / (1 + (" & Param1 & ") ^ 2"
'    Param1 = Differentiate(Param1)
'    If Param1 = "0" Then TempStr1 = "0" Else TempStr1 = "-" & Param1 & TempStr1
'    Differentiate = TempStr1
'    Exit Function
'End If
'Differentiate = "0"
'End Function
'
Public Function IsTreeDynamic(T As Tree) As Boolean
Dim Z As Long

For Z = 1 To T.BranchCount
    If T.Branches(Z).BranchType = bFunction Then
        Select Case T.Branches(Z).FuncType
        Case fDistance, fAngle, fArea, fArg, fGetX, fGetY, fNorm, fOAngle, fXAng
            IsTreeDynamic = True
            Exit Function
        Case Else
            'do nothing
        End Select
    End If
Next

IsTreeDynamic = False
End Function
