VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Twig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements Leaf
'BASIC = Beginner's All-purpose Symbolic Instruction Code
'BAPASIC
'Beginners an Professionals All-purpose Symbolic Instruction Code
'VB64free = Visual Basic 6 for free and also Visual Basic 64 is Free
'VB64U = Visual Basic 6 for You and also Visual Basic 64 is You
'VB64All = Visual Basic 6 for All and also Visual Basic 64 is for All
'Public ViewIndex As Long   ' der Index in der ListBox, ändert sich ständig
Private Base  As Leaf     ' der twig hat einen Namen und kann auch Content haben
'Public Tree  As Twig       ' der Baum/Ober-Zweig an dem dieser Zweig dranhängt
'Leaf hat schon:
'Public ViewIndex As Long
'Public Tree  As Twig
'das wird von Leaf gebraucht
'Public Name  As String
'Public Value As Variant
'Public Twig as Twig
Private m_IsOpen As Boolean

'Public Twigs As ListOfTwig ' alle Zweige darunter
'Twigs ist jetzt auch ListOfLeaf, weil ListOfLeaf ListOfTwig ersetzen kann
Public Twigs As ListOfLeaf ' alle Zweige darunter
Public Leafs As ListOfLeaf


Private Sub Class_Initialize()
    Set Base = New Leaf
    Set Twigs = MNew.ListOfLeaf(False) 'ListOfTwig
    Set Leafs = MNew.ListOfLeaf(False)
End Sub

Friend Sub New_(aName As String)
    Base.Name = aName
End Sub

Friend Sub NewC(other As Twig)
    Set Base = other.Clone 'NewC_Leaf(other)
    'Base.ViewIndex = other.ViewIndex
    'Set Base.Tree = other.Tree
    'so und die Listen kopieren
    m_IsOpen = other.IsOpen
    Set Twigs = other.Twigs.Clone '(other.Twigs)
    Set Leafs = other.Leafs.Clone 'NewC_ListOfLeaf(other.Leafs)
End Sub

Friend Function Clone() As Twig
    Set Clone = New Twig: Clone.NewC Me
End Function

Public Function Add(ByVal aPath As String) As Leaf
    'man gibt nur einen Path und es wird entschieden, ob ein Zweig oder ein Blatt hinzugefügt wird
    Dim isTwig As Boolean
    Dim nam As String: nam = ParsePath(aPath, isTwig)
    If isTwig Then
        Dim tw As Twig
        If nam = Me.Name Then
            'jetzt leider nochmal parsen und schauen ob istwig
            If Len(aPath) > 0 Then
                nam = ParsePath(aPath, isTwig)
                If isTwig Then
                    Set tw = AddOrGetTwig(nam)
                    Set Add = tw.Add(aPath)
                Else
                    Set Add = AddOrGetLeaf(nam)
                End If
            End If
        Else
            Set tw = AddOrGetTwig(nam)
            If Len(aPath) Then
                Set Add = tw.Add(aPath)
            Else
                Set Add = tw
            End If
        End If
    Else
        Set Add = AddOrGetLeaf(nam)
    End If
End Function

Public Function Remove(curLeafOrTwig As Leaf) As Leaf
    If Not curLeafOrTwig Is Nothing Then
        Dim nam As String: nam = curLeafOrTwig.Name
        Dim ll As ListOfLeaf
        If TypeOf curLeafOrTwig Is Twig Then
            Set ll = curLeafOrTwig.Tree.Twigs
        Else
            Set ll = curLeafOrTwig.Twig.Leafs
        End If
        'den Index im Array herausfinden
        Dim i As Long: i = ll.IndexOf(nam)
        If i >= 0 Then
            ll.Remove i
            Set curLeafOrTwig = Nothing
        End If
    Else
    
    End If
    Set Remove = curLeafOrTwig
End Function

Private Function ParsePath(aPath_inout As String, isTwig_out As Boolean) As String
    Dim pos As Long: pos = InStr(1, aPath_inout, "\")
    If pos > 0 Then
        isTwig_out = True
        ParsePath = Left(aPath_inout, pos - 1)
        aPath_inout = Mid(aPath_inout, pos + 1)
    Else
        ParsePath = aPath_inout
        aPath_inout = ""
    End If
End Function

'##############################'v    Leaf related    v'##############################'
Public Function AddLeaf(aLeaf As Leaf) As Leaf
    Set AddLeaf = Leafs.Add(aLeaf)
    Set AddLeaf.Twig = Me
End Function

Public Function InsertLeaf(aLeaf As Leaf, Index As Long) As Leaf
    Set InsertLeaf = Me.Leafs.Insert(aLeaf, Index)
    Set InsertLeaf.Twig = Me
End Function

Public Function AddOrGetLeaf(nam As String) As Leaf
    If Me.Leafs.Contains(nam) Then
        Set AddOrGetLeaf = Me.LeafByName(nam)
    Else
        Set AddOrGetLeaf = Me.AddLeaf(MNew.Leaf(nam, Empty))
    End If
End Function

'##############################'v    Twig related    v'##############################'
Public Function AddTwig(aTwig As Twig) As Twig
    Set AddTwig = Twigs.Add(aTwig)
    Set AddTwig.Tree = Me
End Function

Public Function InsertTwig(aTwig As Twig, Index As Long) As Twig
    Set InsertTwig = Me.Twigs.Insert(aTwig, Index)
    Set InsertTwig.Tree = Me
End Function

Private Function AddOrGetTwig(nam As String) As Twig
    If Me.Twigs.Contains(nam) Then
        Set AddOrGetTwig = Me.Twigs.ItemK(nam)
    Else
        Set AddOrGetTwig = Me.AddTwig(MNew.Twig(nam))
    End If
End Function

Public Sub Clear()
    Class_Initialize
End Sub

Public Property Get CountTwigs() As Long
    Dim c As Long
    Dim i As Long
    Dim tw As Twig
    For i = 0 To Twigs.Count - 1
        Set tw = Twigs.Item(i)
        c = c + 1 'für den Twig tw an der Position i
        c = c + tw.CountTwigs 'für alle twigs die an tw dranhängen
    Next
    CountTwigs = c
End Property

Public Property Get CountLeafs() As Long
    Dim c As Long
    Dim i As Long
    Dim tw As Twig
    For i = 0 To Twigs.Count - 1
        Set tw = Twigs.Item(i)
        c = c + tw.Leafs.Count '.CountLeafs 'für alle leafs die an tw dranhängen
        c = c + tw.CountLeafs
    Next
    CountLeafs = c
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = m_IsOpen
End Property
Public Property Let IsOpen(ByVal Value As Boolean)
    m_IsOpen = Value
    If m_IsOpen Then
        'auch den ganzen Tree aufklappen, sonst gehts nicht!
        If Not Tree Is Nothing Then Tree.IsOpen = True
    End If
End Property

'Split:
'""    lb:0; ub:-1 -> kein Element
'"a"   lb:0; ub:0  -> 1 Element:  "a"
'"\"   lb:0; ub:1  -> 2 Elemente: "", ""
'"a\"  lb:0; ub:1  -> 2 Elemente: "a", ""
Public Property Get TwigByPath(ByVal aPath As String) As Twig
    ' Funktion holt mithilfe eines Pfades aus einem Tree einen Zweig heraus
    ' hier wird auch "..\" unterstützt d.h. dass man hier zurückgehen kann ohne den davor kennen zu müssen
    ' das "..\" ist relativ
    'hier die funktion ParsePath verwenden!
    
    Dim isTwig As Boolean
    Dim nam As String: nam = ParsePath(aPath, isTwig)
    'Dim tw As Twig
    Dim p As String
    Dim pos As Long
    pos = InStr(1, aPath, "\")
    If pos > 0 Then
        p = Left(aPath, pos) ' - 1)
        If p = "..\" Then
            'nach unten gehen
            aPath = Mid(aPath, pos + 1)
            If Me.Twig Is Nothing Then
                MsgBox "Beware twig is nothing"
            Else
                If aPath = vbNullString Then
                    'einfach nur nach oben
                    Set TwigByPath = Tree
                Else
                    Set TwigByPath = Tree.Twig(aPath)
                End If
                'OK oder jetzt gleich in der Liste von Tree suchen
                'erst wenn nochmal "..\" dann weiter nach unten gehen
            End If
        ElseIf p = Me.Name & "\" Then
            aPath = Mid(aPath, pos + 1)
            pos = InStr(1, aPath, "\")
            If pos = 0 Then
                Set Twig = Me
            Else
                p = Left(aPath, pos) ' - 1)
                Dim T As Twig
                Dim i As Long
                For i = 0 To Twigs.Count - 1
                    Set T = Twigs.Item(i)
                    If p = T.Name & "\" Then
                        'aPath = Mid(aPath, pos + 1)
                        If T.Twigs.Count = 0 Then
                            Set TwigByPath = T
                        Else
                            Set TwigByPath = T.Twig(aPath)
                            Exit Property
                        End If
                    End If
                Next
            End If
        Else
            If Not Me.Tree Is Nothing Then
                If p = Me.Tree.Name Then
                    Set TwigByPath = Me.Tree
                Else
                    'jetzt die ganze Meute durchsuchen
                    For i = 0 To Twigs.Count - 1
                        Set T = Twigs.Item(i)
                        If p = T.Name & "\" Then
                            'aPath = Mid(aPath, pos + 1)
                            If T.Twigs.Count = 0 Then
                                Set TwigByPath = T
                            Else
                                Set TwigByPath = T.Twig(aPath)
                                Exit Property
                            End If
                        End If
                    Next
                End If
            Else
                'jetzt die ganze Meute durchsuchen
                For i = 0 To Twigs.Count - 1
                    Set T = Twigs.Item(i)
                    If p = T.Name & "\" Then
                        'aPath = Mid(aPath, pos + 1)
                        If T.Twigs.Count = 0 Then
                            Set TwigByPath = T
                        Else
                            Set TwigByPath = T.Twig(aPath)
                            Exit Property
                        End If
                    End If
                Next
            End If
        End If
    End If
End Property

Public Property Get LeafByName(aName As String) As Leaf
    Set Leaf = Leafs.ItemK(aName)
End Property

Public Property Get Path() As String
    Path = Me.Name & "\"
    If Not Me.Tree Is Nothing Then Path = Me.Tree.Path & Path
End Property

Public Sub ToListBox(aLB As ListBox, ind As IndentStack, aPath As String)
    Dim ve As String: ve = IIf(Me.IsOpen, "v-", ">-")
    Me.ViewIndex = aLB.ListCount '- 1
    aLB.AddItem ind.Peek & ve & Me.Name
    aLB.ItemData(aLB.ListCount - 1) = ObjPtr(Me)
    ind.Push
        If Me.IsOpen Then
            Dim i As Long
            Dim tw As Twig
            For i = 0 To Twigs.Count - 1
                Set tw = Twigs.Item(i)
                tw.ToListBox aLB, ind, aPath & Me.Name & "\"
            Next
            Dim lf As Leaf
            For i = 0 To Leafs.Count - 1
                Set lf = Leafs.Item(i)
                ve = IIf(i < Leafs.Count - 1, "|-", "'-")
                lf.ViewIndex = aLB.ListCount '- 1
                aLB.AddItem ind.Peek & ve & lf.Name
                aLB.ItemData(aLB.ListCount - 1) = ObjPtr(lf)
            Next
        End If
    ind.Pop
End Sub

Public Sub OpenAll()
    Dim tw As Twig, i As Long
    For i = 0 To Twigs.Count - 1
        Set tw = Twigs.Item(i)
        tw.IsOpen = True
        tw.OpenAll
    Next
End Sub
Public Sub CloseAll()
    Dim tw As Twig, i As Long
    For i = 0 To Twigs.Count - 1
        Set tw = Twigs.Item(i)
        tw.IsOpen = False
        tw.CloseAll
    Next
End Sub
Public Function CountOpenTwigs() As Long
    Dim tw As Twig, i As Long, c As Long
    For i = 0 To Twigs.Count - 1
        Set tw = Twigs.Item(i)
        If tw.IsOpen Then c = c + 1
        c = c + tw.CountOpenTwigs
    Next
    CountOpenTwigs = c
End Function

'########################################'    Public properties related to interface Leaf   '########################################'
'Alle Property Get und Let zur Vereinfachung
Public Property Get ViewIndex() As Long
    ViewIndex = Base.ViewIndex
End Property
Public Property Let ViewIndex(ByVal aValue As Long)
    Base.ViewIndex = aValue
End Property

Public Property Get Tree() As Twig
    Set Tree = Base.Twig
End Property
Public Property Set Tree(ByVal aValue As Twig)
    Set Base.Twig = aValue
End Property

Public Property Get Twig() As Twig
    Set Twig = Me
End Property
Public Property Set Twig(aValue As Twig)
    'der Zweig kann sich nicht selber verändern
End Property

Public Property Get Name() As String
    Name = Base.Name
End Property
Public Property Let Name(ByVal aValue As String)
    Base.Name = aValue
End Property

Public Property Get Value() As Variant
    Value = Base.Value
End Property
Public Property Let Value(aValue As Variant)
    Base.Value = aValue
End Property
Public Property Set Value(aValue)
    Set Base.Value = aValue
End Property

'########################################'    Interface Leaf    '########################################'
Private Property Get Leaf_Name() As String
    Leaf_Name = Me.Name
End Property
Private Property Let Leaf_Name(ByVal RHS As String)
    Me.Name = RHS
End Property

Private Property Get Leaf_Tree() As Twig
    Set Leaf_Tree = Me.Tree
End Property
Private Property Set Leaf_Tree(ByVal Value As Twig)
    Set Me.Tree = Value
End Property

Private Property Get Leaf_Twig() As Twig
    Set Leaf_Twig = Me
End Property
Private Property Set Leaf_Twig(ByVal RHS As Twig)
    'an object cannot change itself
End Property

Private Property Get Leaf_Value() As Variant
    Leaf_Value = Me.Value
End Property
Private Property Let Leaf_Value(RHS As Variant)
    Me.Value = RHS
End Property
Private Property Set Leaf_Value(RHS As Variant)
    Set Me.Value = RHS
End Property

Private Property Get Leaf_ViewIndex() As Long
    Leaf_ViewIndex = Me.ViewIndex
End Property
Private Property Let Leaf_ViewIndex(ByVal RHS As Long)
    Base.ViewIndex = RHS
End Property

Private Property Get Leaf_Path() As String
    Leaf_Path = Me.Path
End Property

