VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOfTwig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As Twig
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0)
End Sub

Friend Sub NewC(other As ListOfTwig)
    other.CloneTo m_Arr, m_Count
End Sub

Friend Sub CloneTo(Arr_out() As Twig, n_out As Long)
    Arr_out = m_Arr: n_out = m_Count
    ReDim Preserve Arr_out(0 To m_Count - 1)
End Sub

Public Function Add(Value As Twig) As Twig
    Set Add = Value
    Dim u As Long: u = UBound(m_Arr)
    If u < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2 - 1)
    Set m_Arr(m_Count) = Value
    m_Count = m_Count + 1
End Function

Public Sub Clear()
    Class_Initialize
    m_Count = 0
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Item(ByVal Index As Long) As Twig
    'If m_Count - 1 < Index Then Err.Raise 9
    If m_Count - 1 < Index Then MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1: Exit Property
    Set Item = m_Arr(Index)
End Property

Public Property Set Item(ByVal Index As Long, Value As Twig)
    'If m_Count - 1 < Index Then Err.Raise 9
    If m_Count - 1 < Index Then MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1: Exit Property
    Set m_Arr(Index) = Value
End Property

