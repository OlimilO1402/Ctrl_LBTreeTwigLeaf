Attribute VB_Name = "MArray"
Option Explicit

Public Declare Sub RtlMoveMemory Lib "kernel32" ( _
    ByRef Dst As Any, ByRef src As Any, ByVal bytLength As Long)
    
Public Declare Sub RtlZeroMemory Lib "kernel32" ( _
    ByRef Dst As Any, ByVal bytLength As Long)
    
' die Funktion ArrPtr geht bei allen Arrays außer bei String-Arrays
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" ( _
    ByRef Arr() As Any) As Long

' deswegen hier eine Hilfsfunktion für StringArrays
Public Function StrArrPtr(ByRef StrArr As Variant) As Long
    Call RtlMoveMemory(StrArrPtr, ByVal VarPtr(StrArr) + 8, 4)
End Function

' jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
' um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
' anderen Array zuzuweisen.
Public Property Get VSAPtr(VArr As Variant) As Long
    Call RtlMoveMemory(VSAPtr, ByVal VarPtr(VArr) + 8, 4)
End Property

Public Property Let VSAPtr(VArr As Variant, ByVal RHS As Long)
    Call RtlMoveMemory(ByVal VarPtr(VArr) + 8, ByVal RHS, 4)
End Property

' jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
' um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
' anderen Array zuzuweisen.
Public Property Get SAPtr(ByVal pArr As Long) As Long
    Call RtlMoveMemory(SAPtr, ByVal pArr, 4)
End Property

Public Property Let SAPtr(ByVal pArr As Long, ByVal RHS As Long)
    Call RtlMoveMemory(ByVal pArr, RHS, 4)
End Property

Public Sub ZeroSAPtr(ByVal pArr As Long)
    Call RtlZeroMemory(ByVal pArr, 4)
End Sub

Public Function Col_Contains(aCol As Collection, key As String) As Boolean
    On Error Resume Next
    If IsEmpty(aCol(key)) Then: 'DoNothing
    Col_Contains = (Err.Number = 0)
    On Error GoTo 0
End Function
    


