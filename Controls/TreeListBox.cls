VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type
Private Type RECTAPI
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
'Minimum supported client: Windows Vista [desktop apps only]
Private Declare Function LBItemFromPt Lib "comctl32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal bAutoScroll As Long) As Long
'hLB           Type: HWND    A handle to the list box to check.
'pt            Type: POINT   A POINT structure that contains the screen coordinates to check.
'bAutoScroll   Type: BOOL    A scroll flag. If this parameter is TRUE and the point is directly above or below the list box, the function scrolls the list box by one line and returns -1. Otherwise, the function does not scroll the list box.
'Return value  Type: int     Returns the item identifier if the point is over a list item, or -1 otherwise.

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECTAPI) As Long

Private m_ind As IndentStack
Private m_Tree As Twig

Private m_Selection As Leaf

Private WithEvents mLBTree As ListBox
Attribute mLBTree.VB_VarHelpID = -1

Public Event Click() 'man könnte auch hier gleich das aktuelle Element übergeben
'Public Event DblClick()


'für den Indentstack einfach einen Integer für maximale Einrückungstiefe für ein Element übergeben
Friend Sub New_(aLB As ListBox, aTabSize As Long, aRoot As Twig)
    Set mLBTree = aLB
    Set m_ind = MNew.IndentStack(aTabSize)
    Set m_Tree = aRoot
End Sub
'Copy to MNew:
'Public Function TreeListBox(aLB As ListBox, aTabSize As Long, aRoot As Twig) As TreeListBox
'    Set TreeListBox = New TreeListBox: TreeListBox.New_ aLB, aTabSize, aRoot
'End Function

Private Sub Class_Initialize()
    '
End Sub

Private Sub mLBTree_Click()
    Dim obj As Object: Set obj = MPtr.PtrToObject(mLBTree.ItemData(mLBTree.ListIndex))
    If Not obj Is Nothing Then
        Set m_Selection = obj
    End If
    ZeroObject obj
    RaiseEvent Click
End Sub

Public Property Get Selection() As Leaf
    Set Selection = m_Selection
End Property

Public Property Get Selected(Index As Long) As Boolean
    Selected = mLBTree.Selected(Index)
End Property
Public Property Let Selected(Index As Long, Value As Boolean)
    mLBTree.Selected(Index) = Value
End Property

Private Sub mLBTree_DblClick()
    'Doppelclick dient zum Auf- und Zuklappen der Zweige
    Dim s As String: s = Trim(mLBTree.List(mLBTree.ListIndex))
    Dim s1 As String: s1 = Left(s, 1)
    If s1 = ">" Or s1 = "v" Then
        
        Dim obj As Object: Set obj = MPtr.PtrToObject(mLBTree.ItemData(mLBTree.ListIndex))
        
        If Not obj Is Nothing Then
            If TypeOf obj Is Twig Then
                Dim tw As Twig: Set tw = obj
                tw.IsOpen = Not tw.IsOpen
                Render
                Set tw = Nothing
            Else
                'bei Leaf nix machen
            End If
        End If
        ZeroObject obj
    Else
        
    End If
End Sub

Private Function BottomIndex() As Long
    'liefert den ViewIndex des untersten angzeigten Elements
    'erst das Rectangle bzw die Screen-Koordinaten der ListBox herausfinden,
    'dann einen punkt angeben, damit der Punkt richtig übergeben werden kann
    Dim hr As Long
    Dim rect As RECTAPI
    hr = GetWindowRect(mLBTree.hwnd, rect)
    If hr <> 0 Then
        BottomIndex = LBItemFromPt(mLBTree.hwnd, rect.Left + 5, rect.Bottom - 5, 0)
    End If
End Function

Public Sub Render()
    'Alle offenen Zweige anzeigen
    'so hier jetzt das unterste Item herausfinden,
    Dim botIndex As Long: botIndex = BottomIndex
    
    Dim i As Long: i = mLBTree.ListIndex
    mLBTree.Visible = False
    mLBTree.Clear
    'Hja Mist, der Indentstack muss auch hier rein, haja nadürlisch,
    '
    m_Tree.ToListbox mLBTree, m_ind, ""
    If i < 0 Then i = 0
    If mLBTree.ListCount - 1 < i Then i = mLBTree.ListCount - 1
    If 0 < botIndex And botIndex < mLBTree.ListCount Then
        'zum schluß zuerst das unterste item markieren und erst dann das letzte selektierte item markieren
        'hier evtl auch ob zugeklappt oder aufgeklappt wird
        mLBTree.Selected(botIndex) = True
    End If
    mLBTree.Selected(i) = True
    mLBTree.Visible = True

End Sub




'gefunden hier: http://www.visual-basic5.de/vbclassic/dezem04/popmenue.htm
''Beispiel: VB - Visual Basic - Verwendung der Api Funktion LBItemFromPt zur Index-Ermittlung
''Hinweis : Zum Testen in die ListBox klicken. Linker Mouse-Button ermittelt Index, rechter
''          Mouse-Button öffnet PopUp-Menü.
''
'Option Explicit
'
''The GetCursorPos function retrieves the cursor's position, in screen coordinates.
''Return Values:
''   If the function succeeds, the return value is nonzero.
''   If the function fails, the return value is zero.
'Private Declare Function GetCursorPos Lib "user32.dll" _
'                                (ByRef lpPoint As POINTAPI) _
'                                As Long
'
'Private Type POINTAPI
'    x       As Long
'    y       As Long
'End Type
''The LBItemFromPt function retrieves the index of the item at the specified point in a
''list box.
''Return Values:
''   The return value is the item identifier if the point is over a list item,
''   or  - 1 otherwise.
'Private Declare Function LBItemFromPt Lib "comctl32.dll" _
'                                (ByVal hwnd As Long, _
'                                ByVal x As Long, _
'                                ByVal y As Long, _
'                                ByVal bAutoScroll As Long) _
'                                As Long
'
'Private Function GetListIndexFromPoint(ByVal hwnd As Long) As Long
'    Dim pa      As POINTAPI
'    Dim Result  As Long
'
'    Call GetCursorPos(pa)
'    With pa
'        Result = LBItemFromPt(hwnd, .x, .y, True)
'    End With
'    If Result <> -1 Then
'        GetListIndexFromPoint = Result
'    End If
'End Function
'
'Private Sub File1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    Label1.Caption = CStr(GetListIndexFromPoint(File1.hwnd))
'End Sub
'
'Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbRightButton And (GetListIndexFromPoint(File1.hwnd) <> -1) Then
'        PopupMenu mnuTestHeader
'    End If
'End Sub
'
'Private Sub Form_Load()
'    With mnuTestHeader
'        .Visible = False
'    End With
'End Sub
'
'Private Sub mnuPopUp1_Click()
'    Call Unload(Me)
'End Sub
'
'
'_OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE__OLDCODE_
'
'Private Sub mLBTree_Click()
'    Debug.Print "mLBTree_Click"
'    Dim obj As Object: Set obj = PtrToObject(mLBTree.ItemData(mLBTree.ListIndex))
'
'    If Not obj Is Nothing Then
'        Set m_Selection = obj
'        If TypeOf obj Is Twig Then
            'Set m_curTwig = Nothing
            'Set m_curTwig = obj
            'Set m_curLeaf = Nothing
'            TBPath.Text = m_curTwig.Path
'            TBValue.Text = m_curTwig.Value
            'Set t = Nothing
'            If m_Hist Is Nothing Then
'                Set m_Hist = New Collection 'm_curTwig
'            Else
'                m_Hist.Add m_curTwig
'                m_HistIndex = m_HistIndex + 1
'                'Set m_History.Tree = m_History 'zuerst
'                'Set m_History = m_curTwig 'und erst danach, sonst wirds keine Kette
'                UpdateView2
'            End If
'        ElseIf TypeOf obj Is Leaf Then
            'Set m_curLeaf = obj
            'Set m_Selection = obj
            'Set m_curTwig = m_curLeaf.Tree 'sollen wir das machen? warum nicht?
            'Set m_curTwig = m_curLeaf. Nothing
'            TBValue.Text = m_curLeaf.Value
'            TBPath.Text = m_curLeaf.Tree.Path
'            m_Hist.Add m_curLeaf
'            m_HistIndex = m_HistIndex + 1
'            'Set k = Nothing
'        End If
'    End If
'    ZeroObject obj
'    'Debug.Print "RaiseEvent Click"
'    RaiseEvent Click
'Elemente anschauen
    'Dim i As Integer: i = LBTree.ListIndex
    'Label1.Caption = Mid(Trim(LBTree.List(LBTree.ListIndex)), 3)
    'Dim t As Twig
    'Dim k As KeyValue
    'Set m_curLeaf = Nothing
    'ist da irgendetwas dabei, dass man das unbedingt hier machen muss,
'nööö sollte man besser in Form machen
'    Dim obj As Object: Set obj = PtrToObject(LBTree.ItemData(LBTree.ListIndex))
'
'    If Not obj Is Nothing Then
'        If TypeOf obj Is Twig Then
'            'Set m_curTwig = Nothing
'            Set m_curTwig = obj
'            Set m_curLeaf = Nothing
'            TBPath.Text = m_curTwig.Path
'            TBValue.Text = m_curTwig.Value
'            'Set t = Nothing
'            If m_Hist Is Nothing Then
'                Set m_Hist = New Collection 'm_curTwig
'            Else
'                m_Hist.Add m_curTwig
'                m_HistIndex = m_HistIndex + 1
'                'Set m_History.Tree = m_History 'zuerst
'                'Set m_History = m_curTwig 'und erst danach, sonst wirds keine Kette
'            End If
'        ElseIf TypeOf obj Is Leaf Then
'            Set m_curLeaf = obj
'            'Set m_curTwig = m_curLeaf.Tree 'sollen wir das machen? warum nicht?
'            'Set m_curTwig = m_curLeaf. Nothing
'            TBValue.Text = m_curLeaf.Value
'            TBPath.Text = m_curLeaf.Tree.Path
'            m_Hist.Add m_curLeaf
'            m_HistIndex = m_HistIndex + 1
'            'Set k = Nothing
'        End If
'    End If
'    ZeroObject obj
    'RaiseEvent Click
'End Sub

'##############################################################################################################'
'
'DRAG N DROP in einer ListBox und in einem Baum
'==============================================
'
'sind lange schon tot:
'    Private Sub LBTree_DragDrop(Source As Control, X As Single, Y As Single)
'        Debug.Print "LBTree_DragDrop"
'    End Sub
'    Private Sub LBTree_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
'        Debug.Print "LBTree_DragOver"
'    End Sub
'
'gibt es stattdessen
'1) LBTree_OLEStartDrag
'2) LBTree_OLEDragOver
'3) LBTree_OLEGiveFeedback
'   . . . beliebige Anzahl von Paaren von _OLEDragOver und _OLEGiveFeedback
'4) LBTree_OLEDragDrop
'5) LBTree_OLECompleteDrag



Private Sub mLBTree_OLEStartDrag(data As DataObject, AllowedEffects As Long)
    'Private Sub object_OLEStartDrag(data As DataObject, allowedeffects As Long)
    '
    'The OLEStartDrag event syntax has these parts:
    'Part           Description
    'object         An object expression that evaluates to an object in the Applies To list.
    'data           A DataObject object containing formats that the source will provide and, optionally, the data for those formats. If no data is contained in the DataObject, it is provided when the control calls the GetData method. The programmer should provide the values for this parameter in this event. The SetData and Clear methods cannot be used here.
    'allowedeffects A long integer containing the effects that the source component supports. The possible values are listed in Settings. The programmer should provide the values for this parameter in this event.
    '
    'Settings:
    'The settings for allowedeffects are:
    '   Constant            Value   Description
    '   vbDropEffectNone    0       Drop target cannot accept the data.
    '   vbDropEffectCopy    1       (Default) Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation.
    '                                Eine Maske, die anzeigt, daß eine Kopie erstellt wurde/erstellt werden würde.
    '   vbDropEffectMove    2       Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move.
    '                               Eine Maske, die anzeigt, daß eine Verschiebung stattgefunden hat/stattfinden würde.
    '   vbDropEffectScroll  -2147483648 (&H80000000) 'Eine Maske, die anzeigt, daß im Zielfenster für die Ablage ein Bildlauf stattgefunden hat bzw. stattfinden wird.
    '
    'Remarks
    '    'The source component should use the logical Or operator against the supported values and place the result in the allowedeffects parameter. The target component can use this value to determine the appropriate action (and what the appropriate user feedback should be).
    '    'The OLEStartDrag event also occurs if the components OLEDragMode property is set to Automatic. This allows you to add formats and data to the DataObject object after the component has done so. You can also override the default behavior of the component by clearing the DataObject object (using the Clear method) and then adding your data and formats.
    '    'You may wish to defer putting data into the DataObject object until the target component requests it. This allows the source component to save time by not loading multiple data formats. When the target performs the GetData method on the DataObject, the sources OLESetData event will occur if the requested data is not contained in the DataObject. At this point, the data can be loaded into the DataObject, which will in turn provide the data to the target.
    '
    'If the user does not load any formats into the DataObject, then the drag/drop operation is canceled.
    
    AllowedEffects = vbDropEffectCopy Or vbDropEffectMove 'Or vbDropEffectScroll
    'Debug.Print "1) LBTree_OLEStartDrag"
    
    'beim DragDrop im Datei-Explorer ist
    '* [Shift-Taste]: [->] Nach <Zweig> Verschieben
    '* [ Strg-Taste]: [+ ]  Nach <Zweig> Kopieren
    '* [Shift+Strg-Tasten]: Verknüpfung erstellen in <Zweig>



    Dim e As OLEDropEffectConstants
    e = vbDropEffectCopy Or vbDropEffectMove 'Or vbDropEffectScroll 'Scroll wird hier nicht unterstützt
    AllowedEffects = e
End Sub
Private Sub mLBTree_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'Private Sub object_OLEDragOver(data As DataObject, effect As Long, button As Integer, shift As Integer, x As Single, y As Single, state As Integer)
'
'The OLEDragOver event syntax has these parts:
'Part Description
'object  An object expression that evaluates to an object in the Applies To list.
'data    A DataObject object containing formats that the source will provide and, in addition, possibly the data for those formats. If no data is contained in the DataObject, it is provided when the control calls the GetData method. The SetData and Clear methods cannot be used here.
'effect  A long integer initially set by the source object identifying all effects it supports. This parameter must be correctly set by the target component during this event. The value of effect is determined by logically Or'ing together all active effects (as listed in Settings). The target component should check these effects and other parameters to determine which actions are appropriate for it, and then set this parameter to one of the allowable effects (as specified by the source) to specify which actions will be performed if the user drops the selection on the component. The possible values are listed in Settings.
'button  An integer which acts as a bit field corresponding to the state of a mouse button when it is depressed. The left button is bit 0, the right button is bit 1, and the middle button is bit 2. These bits correspond to the values 1, 2, and 4, respectively. It indicates the state of the mouse buttons; some, all, or none of these three bits can be set, indicating that some, all, or none of the buttons are depressed.
'shift   An integer which acts as a bit field corresponding to the state of the SHIFT, CTRL, and ALT keys when they are depressed.
'        The SHIFT key is bit 0, the CTRL key is bit 1, and the ALT key is bit 2. These bits correspond to the values 1, 2, and 4, respectively.
'        The shift parameter indicates the state of these keys; some, all, or none of the bits can be set, indicating that some, all, or none of
'        the keys are depressed. For example, if both the CTRL and ALT keys are depressed, the value of shift would be 6.
'x,y     A number that specifies the current horizontal (x) and vertical (y) position of the mouse pointer within the target form or control. The x and y values are always expressed in terms of the coordinate system set by the ScaleHeight, ScaleWidth, ScaleLeft, and ScaleTop properties of the object.
'state   An integer that corresponds to the transition state of the control being dragged in relation to a target form or control. The possible values are listed in Settings.
'
'settings
'
'The settings for effect are:
'Constant            Value   Description
'vbDropEffectNone    0       Drop target cannot accept the data.
'vbDropEffectCopy    1       Drop results in a copy of data from the source to the target. The original data is unaltered by the drag operation.
'vbDropEffectMove    2       Drop results in data being moved from drag source to drop source. The drag source should remove the data from itself after the move.
'vbDropEffectScroll -2147483648#
'
'(&H80000000)
'    Scrolling is occurring or about to occur in the target component. This value is used in conjunction with the other values. Note   Use only if you are performing your own scrolling in the target component.
'
'The settings for state are:
'Constant    Value   Description
'vbEnter     0       Source component is being dragged within the range of a target.
'vbLeave     1       Source component is being dragged out of the range of a target.
'vbOver      2       Source component has moved from one position in the target to another.
    
    Dim s As ShiftConstants: s = Shift
    If (s And vbShiftMask) = vbShiftMask Then 'Shift-Key
        'Debug.Print "Shift"
        Effect = vbDropEffectCopy
    End If
    If (s And vbCtrlMask) = vbCtrlMask Then 'Ctrl-Key
        'Debug.Print "Strg"
        Effect = vbDropEffectMove
    End If
    'If Shift = 3 Then
    '    Debug.Print "Shift+Strg"
    '    'effect = effect Or vbDropEffectScroll
    '    effect = vbDropEffectScroll 'nope geht nicht
    'End If
    'mit x und y kann man rausfinden über welchem Item sich die Maus befindet, dann falls ein Zweig und falls ca 500ms lang der gleiche Zweig , dann den Zweig aufklappen
    'Debug.Print "2) LBTree_OLEDragOver"
    'Dim c As ClipBoardConstants
    'Dim e As OLEDropEffectConstants
    'Dim s As ShiftConstants
    'hier noch keine Operationen ausführen, weil DragOver und GiveFeedback u.U. mehrmals aufgerufen werden
End Sub
Private Sub LBTree_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    'Dim e As OLEDropEffectConstants
    'hier Effect setzen
    'Effect kann sein:
    'vbDropEffectNone = 0
    'vbDropEffectCopy = 1
    'vbDropEffectMove = 2
    'vbDropEffectScroll = -2147483648 (&H80000000)
    'Debug.Print "3) LBTree_OLEGiveFeedback"
    'Effect = vbDropEffectNone
    'Effect = vbDropEffectCopy
    'Effect = vbDropEffectMove
    'Effect = vbDropEffectScroll
    'DefaultCursors = Not DefaultCursors
    'DefaultCursors
    'A boolean value which determines whether Visual Basic uses the default mouse cursor proved by the component, or uses a user-defined mouse cursor.
    'True (default) = use default mouse cursor.
    'False = do not use default cursor. Mouse cursor must be set with the MousePointer property of the Screen object.
End Sub

Private Sub LBTree_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    'If data.GetFormat(vbCFFiles) Then Debug.Print "vbCfFiles"
    If data.GetFormat(vbCFText) Then
        Dim s As String: s = data.GetData(vbCFText)
        Debug.Print "vbCFText: |" & s
    End If
    'Debug.Print "4) LBTree_OLEDragDrop"
End Sub

Private Sub LBTree_OLECompleteDrag(Effect As Long)
    'Debug.Print "5) LBTree_OLECompleteDrag"
End Sub

Private Sub LBTree_OLESetData(data As DataObject, DataFormat As Integer)
'    Debug.Print "LBTree_OLESetData"
End Sub


