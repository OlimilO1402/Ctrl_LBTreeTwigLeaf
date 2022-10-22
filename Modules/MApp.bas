Attribute VB_Name = "MApp"
Option Explicit

Sub Main()
    FMain.Show
End Sub

'einen Baum in String darstellen
'ja jetzt die Frage sollen nur Treezweige oder auch Leafs dargestellt werden?
'v-Wurzelelement
'  v-Zweig 1
'   |-Blatt 1
'   |-Blatt 2
'   '-Blatt 3
'  v-Zweig 2
'    v-Zweig 3
'     | |-Blatt 4
'     | |-Blatt 5
'     | '-Blatt 6
'     >Zweig 4
'  >-Zweig 5
'  >-Zweig 6
'  v-Zweig 7
'   |-Blatt 7
'   |-Blatt 8
'   '-Blatt 9

'System
Public Function CreateDefaultTree() As Twig
    'this function creates a default tree as an example, and
    'at the same time it shows how easy it is to create one.
    Dim Root As Twig: Set Root = MNew.Twig("Wurzelelement")
    With Root
        .IsOpen = True
        .Value = "Dies ist die Graue Eminenz!"
        With .AddTwig(MNew.Twig("Zweig 1"))
            .IsOpen = True
            .AddLeaf MNew.Leaf("Blatt 1", "TestText eins Dingsdongs")
            .AddLeaf MNew.Leaf("Blatt 2", "TestText ZWEIEEEE")
            .AddLeaf MNew.Leaf("Blatt 3", "TestText hoibe drui")
        End With
        With .AddTwig(MNew.Twig("Zweig 2"))
            .IsOpen = True
            With .AddTwig(MNew.Twig("Zweig 3"))
                .IsOpen = True
                .AddLeaf MNew.Leaf("Blatt 4", "TestText Quttruorrror")
                .AddLeaf MNew.Leaf("Blatt 5", "TestText fIMpf")
                .AddLeaf MNew.Leaf("Blatt 6", "TestText XXXEEESSS")
            End With
            .AddTwig MNew.Twig("Zweig 4")
        End With
        .AddTwig MNew.Twig("Zweig 5")
        .AddTwig MNew.Twig("Zweig 6")
        With .AddTwig(MNew.Twig("Zweig 7"))
            .IsOpen = True
            .AddLeaf MNew.Leaf("Blatt 7", "TestText SIMMMM")
            .AddLeaf MNew.Leaf("Blatt 8", "TestText OCCCT")
            .AddLeaf MNew.Leaf("Blatt 9", "TestText Noinnn")
        End With
    End With
    Set CreateDefaultTree = Root
End Function

Public Function CreateRandomTree(ByVal n As Long) As Twig
    'OK so jetzt wollen wir einen Tree mit vielen Zweigen erzeugen
    'und die Zweige sollen alle irgendwie zuf‰llig angeordnet sein
    'auﬂerdem sollen sich in den Zweigen viele Bl‰tter tummeln
    Randomize
    Dim Root As Twig: Set Root = MNew.Twig(RandomName(20))
    Dim i As Long, j As Long, k As Long, L As Long
    For i = 1 To n
        With Root.AddTwig(MNew.Twig(RandomName(20)))
            For j = 1 To n
                With .AddTwig(MNew.Twig(RandomName(20)))
                    For k = 1 To n
                        With .AddTwig(MNew.Twig(RandomName(20)))
                            For L = 1 To (n + Rnd * 10)
                                .AddLeaf MNew.Leaf(RandomName(10), RandomName(100))
                            Next
                        End With
                    Next
                    For L = 1 To (n + Rnd * 10)
                        .AddLeaf MNew.Leaf(RandomName(10), RandomName(100))
                    Next
                End With
            Next
            For L = 1 To (n + Rnd * 10)
                .AddLeaf MNew.Leaf(RandomName(10), RandomName(100))
            Next
        End With
    Next
    Set CreateRandomTree = Root
End Function

Public Function RandomName(ByVal NameLen As Long) As String
    Randomize
    Dim i As Long
    Dim b As Boolean
    Dim c As Long
    Dim r As Double
    For i = 1 To NameLen
        r = Rnd
        b = CBool(CLng(r))
        If b Then
            If i = 1 Then
                c = RandomBetween(65, 90)
            Else
                If (Int(r * 100) Mod 5) = 0 Then c = 32 Else c = RandomBetween(65, 90)
            End If
        Else
            c = RandomBetween(97, 122)
        End If
        RandomName = RandomName & ChrW(c)
    Next
End Function

Public Function RandomBetween(ByVal minval_incl As Long, ByVal maxval_incl As Long) As Long
    RandomBetween = minval_incl + Rnd * (maxval_incl - minval_incl)
End Function


