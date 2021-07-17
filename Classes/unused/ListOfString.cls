VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListOfString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Arr() As String
Private m_Count As Long

Private Sub Class_Initialize()
    ReDim m_Arr(0)
End Sub

Friend Sub NewC(other As ListOfString)
    other.CloneTo m_Arr, m_Count
End Sub

Friend Sub CloneTo(Arr_out() As String, n_out As Long)
    Arr_out = m_Arr: n_out = m_Count
    ReDim Preserve Arr_out(0 To m_Count - 1)
End Sub

Public Sub Add(Value As String)
    Dim u As Long: u = UBound(m_Arr)
    If u < m_Count Then ReDim Preserve m_Arr(0 To m_Count * 2 - 1)
    m_Arr(m_Count) = Value
    m_Count = m_Count + 1
End Sub

Public Sub Clear()
    Class_Initialize
    m_Count = 0
End Sub

Public Property Get Count() As Long
    Count = m_Count
End Property

Public Property Get Item(ByVal Index As Long) As String
    'If m_Count - 1 < Index Then Err.Raise 9
    If m_Count - 1 < Index Then MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1: Exit Property
    Item = m_Arr(Index)
End Property

Public Property Let Item(ByVal Index As Long, Value As String)
    'If m_Count - 1 < Index Then Err.Raise 9
    If m_Count - 1 < Index Then MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1: Exit Property
    m_Arr(Index) = Value
End Property

Public Sub Insert(ByVal Index As Long, Value As String)
    'Fügt an der Stelle Index ein neues Element Value ein
    'Index ist 0-basiert
    If m_Count - 1 < Index Then
        MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1
        Exit Sub
    End If
    Dim u As Long: u = UBound(m_Arr)
    If u < m_Count Then ReDim Preserve m_Arr(0 To (u + 1) * 2 - 1)
    'jetzt alle Elemente um eins nach unten kopieren
    Dim arrLng() As Long: SAPtr(ArrPtr(arrLng)) = SAPtr(StrArrPtr(m_Arr))
    Dim pElem As Long: pElem = VarPtr(arrLng(Index))
    RtlMoveMemory ByVal pElem + ptrSize, ByVal pElem, 4 * (m_Count - Index)
    arrLng(Index) = 0 'den Zeiger auf den String zu null setzen, sonst doppelt
    'das Hilfsarray wieder löschen
    ZeroSAPtr ArrPtr(arrLng)
    m_Arr(Index) = Value
    m_Count = m_Count + 1
End Sub

Public Sub Remove(ByVal Index As Long)
    'Entfernt das Element an der Stelle Index
    'Index ist 0-basiert
    If m_Count - 1 < Index Then
        MsgBox "Index: " & Index & " out of bounds: " & m_Count - 1
        Exit Sub
    End If
    'das Item an der Stelle Index zuerst zu "" setzen
    m_Arr(Index) = vbNullString
    'und alles danach um eins nach oben kopieren
    Dim arrLng() As Long: SAPtr(ArrPtr(arrLng)) = SAPtr(StrArrPtr(m_Arr))
    Dim pElem As Long: pElem = VarPtr(arrLng(Index))
    RtlMoveMemory ByVal pElem, ByVal pElem + ptrSize, 4 * (m_Count - Index - 1)
    ZeroSAPtr ArrPtr(arrLng)
    m_Count = m_Count - 1
End Sub

Public Sub ReadFromFile(FNm As String)
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary Access Read As FNr
    Dim s As String: s = Space(LOF(FNr))
    Get FNr, , s
    m_Arr = Split(s, GetLineSeparator(s))
Finally:
    Close FNr
End Sub

'hmm, ist das OK so?
Private Function GetLineSeparator(sFile As String) As String
    Dim u1 As Long: u1 = UBound(Split(sFile, vbCrLf))
    Dim u2 As Long: u2 = UBound(Split(sFile, vbCr))
    Dim u3 As Long: u3 = UBound(Split(sFile, vbLf))
    Dim u As Long: u = Max(u1, Max(u2, u3))
    GetLineSeparator = IIf(u = u1, vbCrLf, IIf(u = u2, vbCr, vbLf))
End Function

Friend Function ToStr() As String
    ToStr = Join(m_Arr, vbCrLf)
End Function

Public Sub WriteToFile(FNm As String)
Try: On Error GoTo Finally
    Dim FNr As Integer: FNr = FreeFile
    Open FNm For Binary Access Write As FNr
    Dim s As String: s = Join(m_Arr, vbCrLf)
    Put FNr, , s
Finally:
    Close FNr
End Sub

Public Sub Sort()
    'n.y.i.
End Sub

