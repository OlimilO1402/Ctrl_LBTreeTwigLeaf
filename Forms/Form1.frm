VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "TreeTwigLeaf, TreeView in a ListBox"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnFont 
      Caption         =   "Font"
      Height          =   375
      Left            =   9120
      TabIndex        =   18
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnStatistics 
      Caption         =   "Statistics"
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnCloseAllTwigs 
      Caption         =   "Close All"
      Height          =   375
      Left            =   9120
      TabIndex        =   19
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnOpenAllTwigs 
      Caption         =   "Open All"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnCountLeafs 
      Caption         =   "Count leafs"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnCountTwigs 
      Caption         =   "Count twigs"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton BtnRandomTree 
      Caption         =   "Random Tree"
      Height          =   375
      Left            =   6720
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnDefaultTree 
      Caption         =   "Default Tree"
      Height          =   375
      Left            =   5520
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton BtnAddTwigOrLeafByPath 
      Caption         =   "Add Twig Or Leaf By Path"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton BtnRename 
      Caption         =   "Rename"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnGoUp 
      Caption         =   "^_Go Up"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton BtnForw 
      Caption         =   "Forw->"
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton BtnBackw 
      Caption         =   "<-Backw"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton BtnDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnAddLeaf 
      Caption         =   "Add Leaf"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BtnAddTwig 
      Caption         =   "Add Twig"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Panel1 
      BorderStyle     =   0  'Kein
      Height          =   5175
      Left            =   0
      ScaleHeight     =   5175
      ScaleWidth      =   9015
      TabIndex        =   4
      Top             =   1320
      Width           =   9015
      Begin VB.TextBox TBValue 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   6
         Top             =   0
         Width           =   5895
      End
      Begin VB.ListBox LBTree 
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4920
         ItemData        =   "Form1.frx":1782
         Left            =   0
         List            =   "Form1.frx":1784
         MultiSelect     =   2  'Erweitert
         OLEDragMode     =   1  'Automatisch
         OLEDropMode     =   1  'Manuell
         TabIndex        =   5
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.TextBox TBPath 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   10455
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Tree As Twig
Private m_curLeaf As Leaf

Private mSplit1 As Splitter
Private WithEvents mTreeLB As TreeListBox
Attribute mTreeLB.VB_VarHelpID = -1

Private m_Hist As Collection 'die Historie, wie der user sich im Baum bewegt hat
Private m_HistIndex As Long
Private m_HistActionFlag  As Boolean
   
Private Sub Form_Load()
    Set m_Tree = CreateDefaultTree
    Set mTreeLB = MNew.TreeListBox(LBTree, 4, m_Tree)
    
    BtnBackw.Enabled = False
    BtnForw.Enabled = False
    Set m_Hist = New Collection
    Set mSplit1 = MNew.Splitter(False, Me, Panel1, "mSplit1", LBTree, TBValue)
    mSplit1.LeftTopPos = LBTree.Width
    mSplit1.BorderStyle = bsXPStyl
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    'LBTree.ItemData
    UpdateView
End Sub

Private Sub Form_Resize()
    Dim tppx As Single: tppx = IIf(Me.ScaleMode = vbTwips, Screen.TwipsPerPixelX, 1)
    Dim L As Single, T As Single, W As Single, H As Single
    Dim Brdr As Single ': Brdr = 8 * tppx
    L = 1 * Brdr:     W = Me.ScaleWidth - 1 * Brdr - L
    T = Panel1.Top:   H = Me.ScaleHeight - 1 * Brdr - T
    If W > 0 And H > 0 Then
        Panel1.Move L, T, W, H
    End If
End Sub

Private Sub BtnOpenAllTwigs_Click()
    m_Tree.OpenAll
    UpdateView
End Sub
Private Sub BtnCloseAllTwigs_Click()
    m_Tree.CloseAll
    UpdateView
End Sub

'##############################'    erste Zeile Schalter    '##############################'
Private Sub BtnAddTwig_Click()
    'dem aktuellen Zweig einen neuen Unterzweig hinzufügen
    Dim Name As String
    If GotNameAndValue("Neuer Ordner", "Neuer Ordner", Name) = vbOK Then
        m_curLeaf.Twig.AddTwig MNew.Twig(Name)
        UpdateView
    End If
End Sub

Private Sub BtnAddLeaf_Click()
    'dem aktuellen Ordner eine Datei hinzufügen
    Dim Name As String, Value As String
    If GotNameAndValue("Neue Datei", "Dateiname; TextderDatei", Name, Value) Then
        m_curLeaf.Twig.AddLeaf MNew.Leaf(Name, Value)
        UpdateView
    End If
End Sub

Private Sub BtnInsert_Click()
    'An aktueller Position ein Leaf-Element einfügen, alle darunter werden um 1 nach unten geschoben
    Dim Name As String, Value As String
    Dim i  As Long
    If TypeOf m_curLeaf Is Twig Then
        If GotNameAndValue("Neuer Zweig", "Zweigname", Name) = vbOK Then
            Dim tr As Twig: Set tr = m_curLeaf.Tree
            i = tr.Twigs.IndexOf(m_curLeaf.Name)
            tr.InsertTwig i, MNew.Twig(Name)
            UpdateView
        End If
    Else
        If GotNameAndValue("Neue Datei", "Dateiname; TextderDatei", Name, Value) = vbOK Then
            Dim tw As Twig: Set tw = m_curLeaf.Twig
            i = tw.Leafs.IndexOf(m_curLeaf.Name)
            tw.InsertLeaf i, MNew.Leaf(Name, Value)
            UpdateView
        End If
    End If
End Sub

Private Function GotNameAndValue(title As String, prompt As String, Name_out As String, Optional ByRef Value_out As String = "") As VbMsgBoxResult
    Dim s As String: s = InputBox(title, prompt, Name_out)
    If StrPtr(s) Then
        Dim Name As String, Value As String
        If InStr(1, s, ";") Then
            Dim sa() As String: sa = Split(s, ";")
            If UBound(sa) > -1 Then Name_out = sa(0)
            If UBound(sa) > 0 Then Value_out = sa(1)
        Else
            Name_out = s
            Value_out = ""
        End If
        GotNameAndValue = vbOK
    Else
        GotNameAndValue = vbCancel
    End If
End Function

Private Sub BtnDelete_Click()
    'löscht das aktuelle Element, Zweig oder Leaf
    If Not m_curLeaf Is Nothing Then
        'es wird erst im Twig nachgefragt ist es ein Twig oder Leaf
        If MsgBox("Are you sure to remove the current element: " & m_curLeaf.Name, vbOKCancel) = vbCancel Then Exit Sub
        Set m_curLeaf = m_curLeaf.Twig.Remove(m_curLeaf)
        If m_curLeaf Is Nothing Then
            LBTree.ListIndex = LBTree.ListIndex - 1
            UpdateView
        End If
    Else
        '
    End If
End Sub

Private Sub BtnRename_Click()
    Dim Name As String: Name = m_curLeaf.Name
    If GotNameAndValue("Name ändern:", "Anderer Name:", Name) = vbOK Then
        m_curLeaf.Name = Name
        UpdateView
    End If
End Sub

Private Sub BtnRandomTree_Click()
    If MsgBox("Are you sure to delete the current tree: " & m_Tree.Name & vbCrLf & "and create a new random tree?", vbOKCancel) = vbCancel Then Exit Sub
    Set m_Tree = CreateRandomTree(4)
    Set mTreeLB = MNew.TreeListBox(LBTree, 4, m_Tree)
    UpdateView
End Sub

Private Sub BtnDefaultTree_Click()
    If MsgBox("Are you sure to delete the current tree: " & m_Tree.Name & vbCrLf & "and create a new default tree?", vbOKCancel) = vbCancel Then Exit Sub
    Set m_Tree = CreateDefaultTree()
    Set mTreeLB = MNew.TreeListBox(LBTree, 4, m_Tree)
    UpdateView
End Sub

'##############################'    zweite Zeile Schalter    '##############################'
Private Sub BtnBackw_Click()
    'zum vorigen Element zurückkehren
    m_HistActionFlag = True
    If Not m_Hist Is Nothing Then
        Dim obj As Leaf: Set obj = m_Hist.Item(m_HistIndex)
        'OK, alles was mit LBTree zusammenhängt sollte in die TreeListBox wandern?
        'hier müßte also Selected(i) = True sein
        'oder Property Let ListIndex !!
        LBTree.ListIndex = obj.ViewIndex
        'If TypeOf obj Is Twig Then
        '    Set m_curTwig = obj
        '    'LBTree.ListIndex = m_curTwig.ViewIndex
        'ElseIf TypeOf obj Is Leaf Then
            Set m_curLeaf = obj
        'End If
        'm_curTwig = m_Hist
        m_HistIndex = m_HistIndex - 1
        'If Not m_History.Tree Is Nothing Then
        '    Set m_History = m_History.Tree
        '    'Private Sub TBPath_KeyUp(KeyCode As Integer, Shift As Integer)
        '    TBPath.Text = m_History.Path
        'End If
        UpdateView2
    End If
    m_HistActionFlag = False
End Sub

Private Sub BtnForw_Click()
    'zum vorigen Element zurückkehren
    m_HistActionFlag = True
    If Not m_Hist Is Nothing Then
        m_HistIndex = m_HistIndex + 1
        UpdateView2
    End If
    m_HistActionFlag = False
End Sub

Private Sub BtnGoUp_Click()
    'Go Up
    TBPath.Text = "..\"
    TBPath_KeyUp vbKeyReturn, 0
End Sub

Private Sub BtnAddTwigOrLeafByPath_Click()
    m_Tree.Add TBPath.Text
    UpdateView
End Sub

Private Sub BtnCountTwigs_Click()
    If m_Tree Is Nothing Then Exit Sub
    Dim s As String
    s = "There are " & m_Tree.CountTwigs & " twigs in the tree."
    's = "Im Baum befinden sich " & m_Tree.CountTwigs & " Zweige."
    MsgBox s
End Sub

Private Sub BtnCountLeafs_Click()
    If m_Tree Is Nothing Then Exit Sub
    Dim s As String
    s = "There are " & m_Tree.CountLeafs & " leafs in the tree."
    's = "Im Baum befinden sich " & m_Tree.CountTwigs & " Blätter."
    MsgBox s
End Sub

Private Sub BtnStatistics_Click()
    Dim nlf  As Long:  nlf = m_Tree.CountLeafs
    Dim ntw  As Long:  ntw = m_Tree.CountTwigs
    Dim n    As Long:    n = nlf + ntw
    Dim notw As Long: notw = m_Tree.CountOpenTwigs
    Dim nctw As Long: nctw = ntw - notw
    Dim s    As String
    s = "There are " & n & " elements in the tree:" & vbCrLf & _
        "Leafs: " & nlf & vbCrLf & _
        "Twigs: " & ntw & vbCrLf & _
        "  open (v): " & notw & vbCrLf & _
        "  close(>): " & nctw
    MsgBox s
End Sub

Sub UpdateView()
    mTreeLB.Render
End Sub

Sub UpdateView2()
    BtnBackw.Enabled = IIf(0 < m_HistIndex, True, False)
    BtnForw.Enabled = IIf(m_HistIndex < m_Hist.Count, True, False)
End Sub

Private Sub mTreeLB_Click()
    Set m_curLeaf = mTreeLB.Selection
    TBPath.Text = m_curLeaf.Path
    TBValue.Text = m_curLeaf.Value
    If m_Hist Is Nothing Then
        Set m_Hist = New Collection
    Else
        If Not m_HistActionFlag Then
            m_Hist.Add m_curLeaf
            m_HistIndex = m_HistIndex + 1
        End If
        UpdateView2
    End If
End Sub

Private Sub TBValue_LostFocus()
    If Not m_curLeaf Is Nothing Then
        m_curLeaf.Value = TBValue.Text
    End If
End Sub

Private Sub TBPath_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim tw As Twig
        Dim aPath As String: aPath = TBPath.Text
        If Left(aPath, 3) = "..\" Then
            'relativ zum aktuellen
            Set tw = m_curLeaf.Twig.TwigByPath(aPath)
        Else
            Set tw = m_Tree.TwigByPath(aPath)
        End If
        If tw Is Nothing Then MsgBox "Pfad nicht gefunden: " & vbCrLf & aPath: Exit Sub
        tw.IsOpen = True
        UpdateView
        LBTree.Selected(tw.ViewIndex) = True
    End If
End Sub

Private Sub BtnFont_Click()
    Dim f As StdFont: Set f = LBTree.Font
    Dim fd As New FontDialog
    Set fd.Font = f
    If fd.ShowDialog = vbOK Then
        Set f = fd.Font
        Set TBPath.Font = f
        Set LBTree.Font = f
        Set TBValue.Font = f
    End If
End Sub

