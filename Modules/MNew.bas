Attribute VB_Name = "MNew"
Option Explicit

Public Function IndentStack(aTabSize As Long) As IndentStack
    Set IndentStack = New IndentStack
    IndentStack.New_ aTabSize
End Function

Public Function Leaf(aName As String, aValue) As Leaf
    Set Leaf = New Leaf
    Leaf.New_ aName, aValue
End Function

Public Function Twig(aName As String) As Twig
    Set Twig = New Twig
    Twig.New_ aName
End Function

Public Function CTwig(aLeaf As Leaf) As Twig
    'a simple cast
    Set CTwig = aLeaf
End Function

'Public Function ListOfLeaf(bUseHashes As Boolean) As ListOfLeaf
'    Set ListOfLeaf = New ListOfLeaf
'    ListOfLeaf.New_ bUseHashes
'End Function
Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function


Public Function Splitter(BolMDI As Boolean, MyOwner As Object, MyContainer As Object, Name As String, LeftTop As Control, LeftBot As Control) As Splitter
    Set Splitter = New Splitter
    Splitter.New_ BolMDI, MyOwner, MyContainer, Name, LeftTop, LeftBot
End Function

Public Function TreeListBox(aLB As ListBox, aTabSize As Long, aRoot As Twig) As TreeListBox
    Set TreeListBox = New TreeListBox
    TreeListBox.New_ aLB, aTabSize, aRoot
End Function

