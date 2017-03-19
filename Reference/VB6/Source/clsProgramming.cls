VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgramming"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 1

Const TotalGrounds = 3
Dim GroundOperations(TotalGrounds) As String
Private Type OperandType
    Part As String
    Operation As String
End Type
Dim MaxVarVisible As Long

Public Function Evaluate(ByVal Expression As String, Optional ByVal XValue As Double = EmptyVar) As Variant
Attribute Evaluate.VB_Description = "Evaluates an expression to its numeric value."
Attribute Evaluate.VB_UserMemId = 0

Dim ParenthesisBalance As Long
Dim Char As String
Dim OperandBuffer As String
Dim Sum As Variant
Dim Count As Long
Dim Ground As Integer
Dim Z As Long, Q As Long, O As Long, V As Double
Dim LongData1 As Long, LongData2 As Long, LongData3 As Long

Dim Operands() As OperandType
Dim TempOperand As OperandType

Expression = UCase(Replace(Expression, " ", ""))
If Expression = "" Then Evaluate = 0: Exit Function

If Left(Expression, 1) = "-" Then Mid(Expression, 1, 1) = "|"
For Z = 1 To TotalGrounds
    For O = 1 To Len(Expression)
        Char = Mid(Expression, O, 1)
        If Char = "(" Then ParenthesisBalance = ParenthesisBalance + 1
        If Char = ")" Then ParenthesisBalance = ParenthesisBalance - 1
        If ParenthesisBalance < 0 Then Err.Raise vbObjectError + 1, , "Invalid parentheses": Evaluate = Null: Exit Function
        If ParenthesisBalance = 0 Then
            If O + 1 < Len(Expression) Then If Char = "|" And Mid(Expression, O + 1, 1) = "-" Then Mid(Expression, O + 1, 1) = "|"
            For Q = 1 To Len(GroundOperations(Z))
                If Char = Mid(GroundOperations(Z), Q, 1) Then
                    Ground = Z
                    If O + 1 < Len(Expression) Then
                        If Mid(Expression, O + 1, 1) = "-" Then Mid(Expression, O + 1, 1) = "|"
                    End If
                    Exit For
                End If
            Next Q
        End If
    Next O
Next Z

If Ground <> 0 Then
    For Z = 1 To Len(Expression)
        Char = Mid$(Expression, Z, 1)
        If Char = "(" Then ParenthesisBalance = ParenthesisBalance + 1
        If Char = ")" Then ParenthesisBalance = ParenthesisBalance - 1
        
        If ParenthesisBalance = 0 And InStr(GroundOperations(Ground), Char) Then
            If OperandBuffer = "" Then
                Err.Raise vbObjectError + 1, , "Invalid expression": Evaluate = Null: Exit Function
            Else
                Count = Count + 1
                ReDim Preserve Operands(Count)
                Operands(Count).Operation = Char
                Operands(Count).Part = OperandBuffer
                OperandBuffer = ""
            End If
        Else
            OperandBuffer = OperandBuffer + Char
        End If
    Next
    
    If ParenthesisBalance <> 0 Then Err.Raise vbObjectError + 1, , "Invalid number of parentheses": Evaluate = Null: Exit Function
    If OperandBuffer = "" Then Err.Raise vbObjectError + 1, , "Expression not finished": Evaluate = Null: Exit Function
    
    Count = Count + 1
    ReDim Preserve Operands(Count)
    Operands(Count).Operation = ""
    Operands(Count).Part = OperandBuffer
    
    If Operands(1).Operation <> "" Then
        Sum = Evaluate(Operands(1).Part, XValue)
        For Z = 2 To Count
            Select Case Operands(Z - 1).Operation
                Case "+"
                    Sum = Sum + Evaluate(Operands(Z).Part, XValue)
                Case "-"
                    Sum = Sum - Evaluate(Operands(Z).Part, XValue)
                Case "*"
                    Sum = Sum * Evaluate(Operands(Z).Part, XValue)
                Case "/"
                    Sum = Sum / Evaluate(Operands(Z).Part, XValue)
                Case "\"
                    Sum = Sum \ Evaluate(Operands(Z).Part, XValue)
                Case "%"
                    Sum = Sum Mod Evaluate(Operands(Z).Part, XValue)
                Case "^"
                    Sum = Sum ^ Evaluate(Operands(Z).Part, XValue)
                Case Else
                    Exit For
            End Select
        Next
    Else
        lMinus = 1
        Do While Left(Operands(1).Part, 1) = "|"
            lMinus = -lMinus
            Operands(1).Part = Right(Operands(1).Part, Len(Operands(1).Part) - 1)
        Loop
        Sum = lMinus * Evaluate(Right$(Operands(1).Part, Len(Operands(1).Part) - 1), XValue)
    End If
    
Else
    
    If Left(Expression, 1) = "(" And Right(Expression, 1) = ")" Then
        Expression = Mid(Expression, 2, Len(Expression) - 2)
        Sum = Evaluate(Expression, XValue)
    Else
        If Left$(Expression, 1) = "|" Then
            lMinus = 1
            Do While Left(Expression, 1) = "|"
                lMinus = -lMinus
                Expression = Right(Expression, Len(Expression) - 1)
            Loop
            Sum = lMinus * Evaluate(Expression)
        Else
            If IsNumeric(Expression) Then
                Sum = Val(Expression)
            Else
                If Expression Like "MINUS(*)" Then
                    Sum = -Evaluate(GetParameter(Expression), XValue)
                ElseIf Expression Like "COS(*)" Then
                   Sum = Cos(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "SIN(*)" Then
                   Sum = Sin(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "EXP(*)" Then
                   Sum = Exp(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "ABS(*)" Then
                   Sum = Abs(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "INT(*)" Then
                   Sum = Int(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "SGN(*)" Then
                   Sum = Sgn(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "SQR(*)" Then
                   Sum = Sqr(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "TAN(*)" Or Expression Like "TG(*)" Then
                    Sum = Tan(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "ATN(*)" Or Expression Like "ARCTG(*)" Or Expression Like "ARCTAN(*)" Or Expression Like "ATAN(*)" Then
                    Sum = Atn(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "ARCSIN(*)" Or Expression Like "ASIN(*)" Then
                    V = Evaluate(GetParameter(Expression), XValue)
                    If Abs(V) >= 1 Then Sum = Sgn(V) * PI / 2 Else Sum = Atn(V / Sqr(1 - V * V))
                ElseIf Expression Like "ARCCOS(*)" Or Expression Like "ACOS(*)" Then
                    V = Evaluate(GetParameter(Expression), XValue)
                    If Abs(V) >= 1 Then Sum = PI / 2 - Sgn(V) * PI / 2 Else Sum = -Atn(V / Sqr(1 - V * V)) + PI / 2
                ElseIf Expression Like "ARCCTG(*)" Then
                    Sum = -Atn(Evaluate(GetParameter(Expression), XValue)) + PI / 2
                ElseIf Expression Like "SEC(*)" Then
                    Sum = 1 / Cos(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "COSEC(*)" Then
                    Sum = 1 / Sin(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "LN(*)" Or Expression Like "LOG(*)" Then
                    Sum = Log(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "LG(*)" Then
                    Sum = Log(Evaluate(GetParameter(Expression), XValue)) / Log(10)
                ElseIf Expression Like "ROUND(*)" Then
                    Sum = Round(Evaluate(GetParameter(Expression), XValue))
                ElseIf Expression Like "RADIANS(*)" Or Expression Like "RAD(*)" Then
                    Sum = Evaluate(GetParameter(Expression), XValue) * ToRadians
                ElseIf Expression Like "DEGREES(*)" Or Expression Like "DEG(*)" Then
                    Sum = Evaluate(GetParameter(Expression), XValue) * ToDegrees
                ElseIf Expression Like "RANDOM(*)" Then
                    Sum = Rnd * Evaluate(GetParameter(Expression), XValue)
                ElseIf Expression Like "DISTANCE(*)" Or Expression Like "DIST(*)" Then
                    LongData1 = GetPointByName(GetParameter(Expression, 1))
                    LongData2 = GetPointByName(GetParameter(Expression, 2))
                    If LongData1 <> 0 And LongData2 <> 0 Then
                        Sum = Distance(BasePoint(LongData1).X, BasePoint(LongData1).Y, BasePoint(LongData2).X, BasePoint(LongData2).Y)
                    Else
                        Sum = 0
                    End If
                ElseIf Expression Like "ANGLE(*)" Then
                    LongData1 = GetPointByName(GetParameter(Expression, 1))
                    LongData2 = GetPointByName(GetParameter(Expression, 2))
                    LongData3 = GetPointByName(GetParameter(Expression, 3))
                    If LongData1 <> 0 And LongData2 <> 0 And LongData3 <> 0 Then
                        Sum = Angle(BasePoint(LongData1).X, BasePoint(LongData1).Y, BasePoint(LongData2).X, BasePoint(LongData2).Y, BasePoint(LongData3).X, BasePoint(LongData3).Y)
                    Else
                        Sum = 0
                    End If
                Else
                    If Expression = "RND" Then
                        Sum = Rnd
                    ElseIf Expression = "TIMER" Then
                        Sum = Timer
                    ElseIf Expression = "PI" Then
                        Sum = 3.14159265358979
                    ElseIf Expression = "E" Then
                        Sum = 2.71828182845905
                    ElseIf Expression = "TORADIANS" Or Expression = "TORAD" Then
                        Sum = ToRadians
                    ElseIf Expression = "TODEGREES" Or Expression = "TODEG" Then
                        Sum = ToDegrees
                    ElseIf Expression = "X" And XValue <> EmptyVar Then
                        Sum = XValue
                    Else
                        Sum = FindVar(Expression, XValue)
                    End If
                End If
            End If
        End If
    End If
End If

Evaluate = CDec(Sum)

End Function

Public Function FindVar(ByVal Expression As String, Optional ByVal XValue As Double = EmptyVar) As Double
Dim Z As Long, tempC As Long, WN As Long
Dim OldMaxVarVisible As Long

WN = GetWatchByName(Expression)

'OldMaxVarVisible = MaxVarVisible
'If MaxVarVisible = 0 Then MaxVarVisible = WECount
'
'If MaxVarVisible < WN Then
'    FindVar = 0
'    Exit Function
'Else
'    MaxVarVisible = WN - 1
'    If MaxVarVisible = 0 Then MaxVarVisible = -1
If WN <> 0 Then
    If WatchExpressions(WN).Expression <> Expression Then
        FindVar = Evaluate(WatchExpressions(WN).Expression, XValue)
        Exit Function
    End If
End If
'End If
'
'MaxVarVisible = OldMaxVarVisible
'------------------------------------------------------------------------------------

For Z = 1 To PointCount
    If InStr(Expression, BasePoint(Z).Name) Then
        If tempC = 0 Then
            tempC = Z
        Else
            FindVar = Distance(BasePoint(tempC).X, BasePoint(tempC).Y, BasePoint(Z).X, BasePoint(Z).Y)
            Exit Function
        End If
    End If
Next

End Function

Public Function EvalWatch(ByVal Expression As String) As Double
On Local Error GoTo EH
EvalWatch = Evaluate(Expression)
'While NeedToRefresh
'   NeedToRefresh = False
'    EvalWatch = Evaluate(Expression)
'Wend
Exit Function

EH:
EvalWatch = 0
End Function

Private Sub Class_Initialize()
GroundOperations(1) = "^"
GroundOperations(2) = "*/\%"
GroundOperations(3) = "+-"
End Sub

Private Function GetParameter(ByVal Expression As String, Optional ByVal ParamNum As Long = 1) As String
Dim ParBalance As Long, Z As Long, S As String, Curr As Long, Responce As String
Expression = Left(Expression, Len(Expression) - 1)
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
Next

GetParameter = Responce

End Function
