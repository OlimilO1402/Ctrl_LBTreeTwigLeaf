VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOfLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'speichert Leafs And/Or Twigs which are also Leafs
Private m_Arr() As Leaf
Private m_Count As Long
'
Private m_UseHashes As Boolean 'wenn UseHashes dann wird die Collection verwendet, sonst wird das Array verwendet
Private m_Col As Collection

Private Sub Init()
    If m_UseHashes Then
        Set m_Col = New Collection
    Else
        ReDim m_Arr(0)
        m_Count = 0
    End If
End Sub

Friend Sub New_(bUseHashes As Boolean)
    m_UseHashes = bUseHashes
    Init
End Sub

Friend Sub NewC(other As ListOfLeaf)
    m_UseHashes = other.UseHashes
    If m_UseHashes Then
        other.CloneToC m_Col
    Else
        other.CloneToA m_Arr, m_Count
    End If
End Sub

Public Function Add(Value As Leaf) As Leaf
    Set Add = Value
    If m_UseHashes Then
        m_Col.Add Value, Value.Name
    Else
        Dim u As Long: u = UBound(m_Arr)
        If u < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2 - 1)
        Set m_Arr(m_Count) = Value
        m_Count = m_Count + 1
    End If
End Function

Public Sub Clear()
    If m_UseHashes Then
        Set m_Col = New Collection
    Else
        ReDim m_Arr(0)
        m_Count = 0
    End If
End Sub

Friend Sub CloneToA(Arr_out() As Leaf, n_out As Long)
    'how to do a deep copy of alle data in the array?
    'all objects must have a copy-constructor by default
    Arr_out = m_Arr: n_out = m_Count
End Sub
Friend Sub CloneToC(Col_out As Collection)
    'how to do a deep copy of alle Data in the Collection?
    'all objects must have a copy-constructor by default
    If Col_out Is Nothing Then Set Col_out = New Collection
    Dim v
    For Each v In m_Col
        Col_out.Add v
    Next
End Sub

Friend Function Clone() As ListOfLeaf
    Set Clone = New ListOfLeaf
    Clone.NewC Me
End Function

Public Function Contains(aKey As String) As Boolean
    If m_UseHashes Then
        On Error Resume Next
        If IsEmpty(m_Col(aKey)) Then: 'DoNothing
        Contains = (Err.Number = 0)
        On Error GoTo 0
    Else
        Dim i As Long
        Dim lf As Leaf
        For i = 0 To m_Count - 1
            Set lf = m_Arr(i)
            If lf.Name = aKey Then Contains = True: Exit Function
        Next
    End If
End Function

Public Property Get Count() As Long
    If m_UseHashes Then Count = m_Col.Count Else Count = m_Count
End Property

Public Property Get IndexOf(aKey As String) As Long
    If m_UseHashes Then
        'error
    Else
        IndexOf = -1
        Dim i As Long, lf As Leaf
        For i = 0 To m_Count - 1
            Set lf = m_Arr(i)
            If lf.Name = aKey Then IndexOf = i: Exit Property
        Next
    End If
End Property
Public Property Get ItemK(aKey As String) As Leaf
    On Error Resume Next
    If m_UseHashes Then
        Set ItemK = m_Col.Item(aKey)
    Else
        Dim i As Long, lf As Leaf
        For i = 0 To m_Count - 1
            Set lf = m_Arr(i)
            If lf.Name = aKey Then Set ItemK = lf: Exit Property
        Next
'        For i = 1 To m_Col.Count
'            Set lf = m_Col.Item(i)
'            If lf.Name = aKey Then Set ItemK = lf: Exit Property
'        Next
    End If
    On Error GoTo 0
End Property

Public Property Get Item(ByVal Index As Long) As Leaf
    If m_UseHashes Then
        Set Item = m_Col.Item(Index + 1)
    Else
        'If m_Count - 1 < Index Then Err.Raise 9
        Set Item = m_Arr(Index)
    End If
End Property

Public Property Set Item(ByVal Index As Long, Value As Twig)
    If m_UseHashes Then
        m_Col.Remove Index
        m_Col.Add Value, , Index
    Else
        'If m_Count - 1 < Index Then Err.Raise 9
        If m_Count - 1 < Index Then MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1: Exit Property
        Set m_Arr(Index) = Value
    End If
End Property

Public Property Get UseHashes() As Boolean
    UseHashes = m_UseHashes
End Property

Public Function Insert(Value As Leaf, Index As Long) As Leaf
    If m_UseHashes Then
        'nope
    Else
        Dim u As Long: u = UBound(m_Arr)
        If u < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2 - 1)
        
        If Index < m_Count Then
            Dim p As Long: p = VarPtr(m_Arr(Index))
            RtlMoveMemory ByVal p + 4, ByVal p, (m_Count - Index) * 4
            RtlZeroMemory ByVal p, 4
            Set m_Arr(Index) = Value
            Set Insert = Value
            m_Count = m_Count + 1
        ElseIf u < Index Then
            Set Insert = Me.Add(Value)
        Else
            'Fehler
            'Err.Raise 9
            MsgBox "Fehler"
        End If
    End If
End Function

Public Function Remove(Index As Long) As Leaf
    If m_UseHashes Then
        Set Remove = m_Col.Item(Index)
        m_Col.Remove Index
    Else
        'hier Fehler?
        Set Remove = m_Arr(Index)
        Set m_Arr(Index) = Nothing
        'und alle darunter nach oben kopieren
        Dim u As Long: u = m_Count - 1
        Dim p As Long: p = VarPtr(m_Arr(Index))
        RtlMoveMemory ByVal p, ByVal p + 4, (u - Index) * 4
        'den letzten nullen
        p = VarPtr(m_Arr(u))
        RtlZeroMemory ByVal p, 4
        m_Count = m_Count - 1
    End If
End Function

'vb-tech.de
'Public Sub CollectionChangeKey(col As Collection, _
'    OldIndex As Variant, NewKey As String)
'
'  Dim Value As Variant
'
'  If IsObject(col(OldIndex)) Then
'    Set Value = col(OldIndex)
'  Else
'    Value = col(OldIndex)
'  End If
'  col.Add Value, NewKey 'erst hinzufügen
'  col.Remove OldIndex   'dann entfernen
'End Sub
'Public Function IsInCollection( _
'    ByRef col As Collection, _
'    ByRef elem As String _
'  ) As Boolean
'  On Error Resume Next
'    If IsEmpty(col(elem)) Then: 'DoNothing
'    IsInCollection = (Err.Number = 0)
'  On Error GoTo 0
'End Function

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
    'Debug.Print "Terminate"
    Erase m_Arr()
    m_Count = 0
    Set m_Col = Nothing
End Sub
