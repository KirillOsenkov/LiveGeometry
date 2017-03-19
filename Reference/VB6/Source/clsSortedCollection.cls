VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSortedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Fast sorted collection."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Class1"
Attribute VB_Ext_KEY = "Member0" ,"Class1"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection
Public Sorted As Boolean

Public Sub Add(Value, Optional ByVal sKey As String)
Dim Z As Long, C As Long
Dim LB As Long, UB As Long

C = mCol.Count
If Sorted And C > 0 Then
    LB = 1
    UB = C
    Do
        Z = (UB + LB) \ 2
        If Value > mCol(Z) Then
            LB = Z + 1
        Else
            If Value = mCol(Z) Then Exit Sub
            UB = Z - 1
        End If
    Loop Until UB < LB
    
    If LB > Z Then
        If Len(sKey) = 0 Then mCol.Add Value, , , Z Else mCol.Add Value, sKey, , Z
    Else
        If Len(sKey) = 0 Then mCol.Add Value, , Z Else mCol.Add Value, sKey, Z
    End If
Else
    If Len(sKey) = 0 Then mCol.Add Value Else mCol.Add Value, sKey
End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Item = mCol(vntIndexKey)
End Property

Public Function FindItem(Value) As Long
Dim Z As Long, C As Long
Dim LB As Long, UB As Long

If mCol.Count > 0 Then
    LB = 1
    UB = mCol.Count
    Do
        Z = (UB + LB) \ 2
        If Value = mCol(Z) Then FindItem = Z: Exit Function
        If Value > mCol(Z) Then LB = Z + 1 Else UB = Z - 1
    Loop Until UB < LB
End If
End Function

Public Sub ReplaceItem(ByVal Index As Long, NewValue, Optional Key As String)
Dim Z As Long

If Index < 0 Or Index > mCol.Count Then Exit Sub
If Index = mCol.Count Then
    mCol.Remove Index
    If Key <> "" Then mCol.Add NewValue, Key Else mCol.Add NewValue
Else
    mCol.Remove Index
    If Sorted Then
        If Key <> "" Then Add NewValue, Key Else Add NewValue
    Else
        If Key <> "" Then mCol.Add NewValue, Key, Index Else mCol.Add NewValue, , Index
    End If
End If
End Sub

Public Property Get Count() As Long
Count = mCol.Count
End Property

Public Sub Remove(Index)
If Index > 0 And Index <= mCol.Count Then mCol.Remove Index
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
Sorted = True
Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
Set mCol = Nothing
End Sub

Public Sub Clear()
Do While mCol.Count > 0
    mCol.Remove 1
Loop
End Sub
