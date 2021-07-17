VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IndentStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_Count As Long
Private m_Space As String
Public TabSize  As Long

Friend Sub New_(aTabSize As Long)
    TabSize = aTabSize
End Sub

Public Sub Clear()
    m_Count = 0
    m_Space = vbNullString
End Sub

Public Sub Push() '    'hinzufügen
    m_Count = m_Count + 1
    m_Space = Space(m_Count * TabSize)
End Sub

Public Sub Pop()
    m_Count = m_Count - 1
    m_Space = Space(m_Count * TabSize)
End Sub

Public Function Peek() As String
    Peek = m_Space
End Function

Public Property Get Count() As Long
    Count = m_Count
End Property

