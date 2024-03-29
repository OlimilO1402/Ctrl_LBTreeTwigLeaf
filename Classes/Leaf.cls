VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Leaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'z.B. f�r Controls in Forms oder f�r Dateien in einem vb-Projekt
'Klasse als Interface ben�tzen f�r Controls oder VB-Dateien
Public Name  As String
Public Twig  As Twig
Public Value As Variant
Public ViewIndex As Long

Friend Sub New_(aName As String, aValue)
    Name = aName
    If IsObject(aValue) Then Set Value = aValue Else Value = aValue
End Sub
'Copy to MNew:
'Public Function Leaf(aName As String, aValue) As Leaf
'    Set Leaf = New Leaf: Leaf.New_ aName, aValue
'End Function

Friend Sub NewC(other As Leaf)
    Me.ViewIndex = other.ViewIndex
    Set Me.Twig = other.Twig
    Me.Name = other.Name
    Me.Value = other.Value
End Sub

Friend Function Clone() As Leaf
    Set Clone = New Leaf: Clone.NewC Me
End Function

Public Property Get Path() As String
    Path = Twig.Path
End Property

Public Property Get Tree() As Twig
    Set Tree = Twig.Tree
End Property
Public Property Set Tree(ByVal Value As Twig)
    '
End Property

'Root->Tree->Twig->Leaf
'      Tree->Twig->Leaf

'aus der Sicht des Leafs ist der Twig der Zweig an dem das Blatt dranh�ngt
'aus der Sicht des Leafs ist der Tree der Baum an dem der Zweig dranh�ngt an dem das Blatt dranh�ngt
'aus der Sicht des Twigs ist der Twig er selber
'aus der Sicht des Twigs ist der Tree der Baum an dem er dranh�ngt

'wenn ein Leaf ein Twig ist dann liefert Twig sich selber und Tree den Tree des Tiwgs
