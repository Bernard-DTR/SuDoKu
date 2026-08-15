'-------------------------------------------------------------------------------------------
'Mardi 11/08/2026
' Stratégie Nice_Loop Continuous & Discontinuous, et AIC (Alternating Inference Chains)
'  
' Préfixe NLC
'   Les stratégies NLC, NLD et AIC sont en famille 7 pour ne pas générer d'erreur dans la gestion du menu
'                                  et elles passeront en famille 4 !
' Quelles sont les ressources nécessaires ? 
'      Je les place dans le module G000_Base.vb
'
'
'
'
'Exemple 1
'5..9.........2.4.98..7...........6...2.......4.6315....1.67..42.....3.7..4.....86
'Hodoku 220 détermine une NLC:
'2/3/4/6/8/9 5= r3c5 =3= r1c5 -3- r1c2 -6- r8c2 =6= r8c1 =2= r8c4 -2- r9c6 -9- r7c6 -8- r2c6 =8= r2c4 =5= r3c5 =3
'=> r9c4<>2, r1c8<>3, r3c5<>4, r2c2,r3c5<>6, r45c6<>8, r45c6,r8c15,r9c5<>9
'5..9.........2.4.989.7...........6...2.......4.631529..1.67..42.....3.7..4.....86
'SDK arrive à ce résultat avec Plcy_Strg_Profondeur         : UOBTXYSJZKQ
'Automate Graphe ne trouve rien

'Exemple 2
'548.......3....8..1.......3.72..86........3......47..9.1.3...2.2569.........6.7..
'Hodoku 220 détermine une NLC:
'1/2/4/5/7/9 7= r1c4 =6= r1c8 =9= r1c5 -9- r3c5 -5- r3c7 -4- r8c7 =4= r8c6 -4- r9c4 =4= r2c4 =7= r1c4 =6
'=> r1c48,r2c4<>1, r12c4<>2, r9c6<>4, r2c4,r3c8<>5, r1c8<>7, r2c56,r3c6,r5c5<>9
'548..3...63....8..12.8....3972.386..4.....3..3...47..971438592625697..38893.6.7..
'SDK arrive à ce résultat avec Plcy_Strg_Profondeur         : UOBTXYSJZKQ
'Automate Graphe ne trouve rien
'NLC : Exclusion(s) détectée(s) pour le candidat 3 : 3

'Exemple 3
'6..91.52......7.....................5....4.16..86..4.5.4....1..93.86......5.4138.
'Hodoku 220 détermine une NLC:
'2/3/5/7/8/9 8= r4c6 =5= r7c6 -5- r7c4 -3- r5c4 -2- r6c5 =2= r6c2 =1= r6c1 =7= r3c1 -7- r1c2 -8- r1c6 =8= r4c6 =5
'=> r45c5<>2, r23c4,r4c6,r6c1<>3, r7c5<>5, r36c2<>7, r1c9<>8, r4c6,r6c2<>
'6.491.52......76.......6...4.6......5....4.16..86..4.5847...162931862754265741389
'SDK arrive à ce résultat avec Plcy_Strg_Profondeur         : UOBTXYSJZKQ
'Automate Graphe ne trouve rien

'NLC: Exclusion(s) détectée(s) pour le candidat 6 : 3
'NLC: Exclusion(s) détectée(s) pour le candidat 6 : 4

'-------------------------------------------------------------------------------------------

Module G300_Strategy_Nice_Loop_Continuous

  '-------------------------------------------------------------------------------------------
  ' Construction des Liens forts et faibles
  ' Il existe plusieurs procédures pour construire les liens forts et faibles,
  '    Elles placent les liens dans une liste de type GLink_Cls
  '    Cette classe GLink_Cls est conforme pour des liens forts et les liens faibles
  '    Il existe 3 listes de liens:
  '                GLinks_Strong As New List(Of GLink_Cls)       ' Liste des liens Forts
  '                GLinks_Weak As New List(Of GLink_Cls)         ' Liste des liens Faibles
  '                GLinks As New List(Of GLink_Cls)              ' Liste des liens Forts et Faibles
  '    il est nécessaire de clearer ces listes avant de les remplir
  '    Elles sont cumulées dans Glinks avec AddRange
  '-------------------------------------------------------------------------------------------
  Public Solver_NLC As New DFS_NLC()
  'Public Noeud_NLC As String = ""

  Public Sub Strategy_NLC(U_temp(,) As String)
    Plcy_Strg = "NLC"
    Jrn_Add_Yellow(Proc_Name_Get() & " " & Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte)
    ' Quel Candidat à traiter ?   
    Dim Titre As String = Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte
    Dim Texte As String = "Candidat à traiter pour afficher les liens S et W :"
    Dim Candidat_IB As String = InputBox(Texte, Titre)
    If Candidat_IB Is "" Then Exit Sub 'Cancel enfoncé


    For Cdd As Integer = 1 To 9
      If Cdd.ToString <> Candidat_IB Then Continue For

      Jrn_Add(, {"NLC : Analyse du candidat " & Cdd.ToString()})
      GRslt_Init()
      If GRslt.Productivité Then Exit For

      Dim Candidat As String = Cdd.ToString()
      GRslt.Candidat = {Candidat, "0"}

      ' Liste des liens forts et Faible pour un candidat donné
      GLinks_Strong.Clear()
      GLinks_Build_Strong(U, Candidat)

      GLinks_Weak.Clear()
      GLinks_Build_Weak(U, Candidat)

      GLinks.Clear()
      GLinks.AddRange(GLinks_Strong)
      GLinks.AddRange(GLinks_Weak)
      GLinks_Display_SW(GLinks)
      Jrn_Add("SDK_Space")

      ' Création des Noeuds du graphe. On passe de GLinks à Graph
      ' Solver_NLC As New DFS_NLC() est Public pour être exploité dans Paint
      Solver_NLC.Graph_NLC_Build(GLinks)
      Solver_NLC.Graph_NLC_Display()
      GRslt.Nb_Noeuds = Solver_NLC.Graph.Count

      'Titre = Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte
      'Texte = "le graphe comporte " & Solver_NLC.Graph.Count & " nœuds. "
      'Texte &= vbCrLf & "Entrez le nœud à traiter pour afficher les chemins :"
      'Noeud_NLC = InputBox(Texte, Titre)
      'If Noeud_NLC Is "" Then Exit Sub 'Cancel enfoncé


      'If GcnlRslt.Productivité Then
      '  For Each excl As GCel_Excl_Cls In GcnlRslt.CelExcl
      '    Jrn_Add(, {"CNL : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & excl.Cel & " - Unités : " & String.Join(", ", excl.Exc.Select(Function(c) U_Coord(c)))})
      '  Next excl

      '  For Each hs As Tuple(Of Integer, String) In GcnlRslt.CelExcl_hs
      '    Jrn_Add_Yellow("CNL : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & hs.Item1 & " - Unité(s) : " & hs.Item2)
      '  Next hs
      '  Exit For

      'End If
    Next
    GRslt.Productivité = True
    GRslt.Nb_Liens = GLinks.Count
    GRslt_Display()
    Frm_SDK.Invalidate()
  End Sub


  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  ' Les liens forts et faibles  
  Public Sub GLinks_Build_Strong(U_temp(,) As String, candidat As String)
    'La procédure remplie la liste Public GLinks_Strong As New List(Of GLink_Cls) des liens forts pour le candidat donné

    Dim i As Integer, j As Integer
    Dim gLink_Unité As String

    ' --- Lignes ---
    For i = 0 To 8
      Dim cells As New List(Of Integer)
      For j = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count = 2 Then
        gLink_Unité = "Row" + (U_Row(cells(0)) + 1).ToString()
        GLinks_Strong.Add(New GLink_Cls With {.Cel = New Integer() {cells(0), cells(1)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Composition = "024"})

      End If
    Next

    ' --- Colonnes ---
    For j = 0 To 8
      Dim cells As New List(Of Integer)
      For i = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count = 2 Then
        gLink_Unité = "Col" + (U_Col(cells(0)) + 1).ToString()
        GLinks_Strong.Add(New GLink_Cls With {.Cel = New Integer() {cells(0), cells(1)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Composition = "024"})
      End If
    Next

    ' --- Régions ---
    For b As Integer = 0 To 8
      Dim cells As New List(Of Integer)
      Dim r0 As Integer = (b \ 3) * 3
      Dim c0 As Integer = (b Mod 3) * 3
      For i = 0 To 2
        For j = 0 To 2
          Dim idx As Integer = (r0 + i) * 9 + (c0 + j)
          If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
        Next
      Next
      If cells.Count = 2 Then
        gLink_Unité = "Reg" + (U_Reg(cells(0)) + 1).ToString()
        GLinks_Strong.Add(New GLink_Cls With {.Cel = New Integer() {cells(0), cells(1)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Composition = "024"})
      End If
    Next

  End Sub
  Public Sub GLinks_Build_Weak(U_temp(,) As String, candidat As String)
    'La procédure remplie la liste Public GLinks_Weak As New List(Of GLink_Cls) des liens faibles pour le candidat donné

    Dim i As Integer, j As Integer
    Dim gLink_Unité As String

    ' --- Lignes ---
    For i = 0 To 8
      Dim cells As New List(Of Integer)
      For j = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b As Integer = a + 1 To cells.Count - 1
            gLink_Unité = "Row" + (U_Row(i) + 1).ToString()
            GLinks_Weak.Add(New GLink_Cls With {.Cel = New Integer() {cells(a), cells(b)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "W",
                                       .Unité = gLink_Unité, .Composition = "024"})
          Next
        Next
      End If
    Next

    ' --- Colonnes ---
    For j = 0 To 8
      Dim cells As New List(Of Integer)
      For i = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b As Integer = a + 1 To cells.Count - 1
            gLink_Unité = "Col" + (U_Col(j) + 1).ToString()
            GLinks_Weak.Add(New GLink_Cls With {.Cel = New Integer() {cells(a), cells(b)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "W",
                                       .Unité = gLink_Unité, .Composition = "024"})
          Next
        Next
      End If
    Next

    ' --- Régions ---
    For b As Integer = 0 To 8
      Dim cells As New List(Of Integer)
      Dim r0 As Integer = (b \ 3) * 3
      Dim c0 As Integer = (b Mod 3) * 3
      For i = 0 To 2
        For j = 0 To 2
          Dim idx As Integer = (r0 + i) * 9 + (c0 + j)
          If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
        Next
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b2 As Integer = a + 1 To cells.Count - 1
            gLink_Unité = "Reg" + (b + 1).ToString()
            GLinks_Weak.Add(New GLink_Cls With {.Cel = New Integer() {cells(a), cells(b2)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "W",
                                       .Unité = gLink_Unité, .Composition = "024"})
          Next
        Next
      End If
    Next
  End Sub
  Public Sub GLinks_Display_SW(L As List(Of GLink_Cls))
    ' Affichage de La liste L Liste des liens forts ou faibles
    Jrn_Add(, {Proc_Name_Get() & " Affichage de L As List(Of GLink_Cls) : " & L.Count & " Lignes."})
    L = L.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList()

    If L.Count <> 0 Then
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In L
        With gLink
          Nb += 1
          Dim S As String = .Cdd(4) & " " &
            U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ")" & " → " &
            U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
            " Lien " & .Type & "  Unité " & .Unité.PadRight(6) & " Comp " & .Composition &
            " Cellules n° " & CStr(.Cel(0)).PadLeft(2) & "-" & CStr(.Cel(1)).PadLeft(2)
          Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(3) & " " & S})
        End With
      Next gLink
    End If
    Jrn_Add("SDK_Space")
  End Sub
  Public Sub Exemple_01()
    'Le programme liste pour un candidat donné les liens forts et faibles,
    ' Les liens sont triés dans le programme GLinks_Display_SW

    Dim Titre As String = "Stratégie NLC"
    Dim Texte As String = "Entrez le candidat à traiter pour afficher les liens S et W :"
    Dim Candidat_IB As String = InputBox(Texte, Titre)
    If Candidat_IB Is "" Then Exit Sub 'Cancel enfoncé
    GRslt_Init()
    GRslt.Candidat = {Candidat_IB, "0"}
    ' Liste des liens forts et Faible pour un candidat donné
    GLinks_Strong.Clear()
    GLinks_Build_Strong(U, Candidat_IB)
    Jrn_Add("SDK_Space")

    GLinks_Weak.Clear()
    GLinks_Build_Weak(U, Candidat_IB)
    Jrn_Add("SDK_Space")

    GLinks.Clear()
    GLinks.AddRange(GLinks_Strong)
    GLinks.AddRange(GLinks_Weak)
    GLinks_Display_SW(GLinks)
    Jrn_Add("SDK_Space")
    GRslt.Productivité = True
    GRslt.Nb_Liens = GLinks.Count
    GRslt_Display()
    Frm_SDK.Invalidate()

  End Sub



  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  Public Class DFS_NLC

    Public Graph As New Dictionary(Of Integer, List(Of Edge))()
    Public AllPaths As New List(Of List(Of GLink_Cls))()

    ' le graph est représenté par un dictionnaire
    '  où chaque clé est un nœud (cellule)
    '  et la valeur est une liste d'arêtes (liens) connectées à ce nœud.
    Public Sub Graph_NLC_Build(L As List(Of GLink_Cls))
      Graph.Clear()

      For Each glink As GLink_Cls In L
        Dim a As Integer = glink.Cel(0)
        Dim b As Integer = glink.Cel(1)

        If Not Graph.ContainsKey(a) Then Graph(a) = New List(Of Edge)
        If Not Graph.ContainsKey(b) Then Graph(b) = New List(Of Edge)

        Graph(a).Add(New Edge With {.Neighbor = b, .Link = glink})
        Graph(b).Add(New Edge With {.Neighbor = a, .Link = glink})
      Next
    End Sub
    Public Sub Graph_NLC_Display()
      Jrn_Add(, {Graph.Count & " entrée(s)."})
      Dim l As Integer = 0
      For Each kvp As KeyValuePair(Of Integer, List(Of Edge)) In Graph
        l += 1
        Dim edges As List(Of Edge) = Graph(kvp.Key)
        Dim sb As New Text.StringBuilder()
        For Each edge As Edge In edges
          sb.AppendFormat(" Voisin = {0}  Lien {1} ({2} → {3}) " & vbCrLf & "                   ",
                          U_Coord(edge.Neighbor),
                          edge.Link.Type,
                          U_Coord(edge.Link.Cel(0)),
                          U_Coord(edge.Link.Cel(1)))
        Next
        Dim edgeCount As String = $" {edges.Count}"
        Jrn_Add(, {$"{l,3} De {U_Coord(kvp.Key)} _{edgeCount,3}_ {sb}"})
      Next
    End Sub
  End Class

  Public Class Edge
    Public Property Neighbor As Integer
    Public Link As GLink_Cls
  End Class

End Module
