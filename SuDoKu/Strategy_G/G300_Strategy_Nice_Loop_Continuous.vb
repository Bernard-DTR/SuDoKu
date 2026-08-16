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

  ' Il y a une similitude de variable et de code avec DFS_Coloration
  '-------------------------------------------------------------------------------------------
  Public Solver_NLC As New DFS()
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
      'GLinks = GLinks.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList()
      'GLinks = GLinks.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ToList()

      'GLinks_Display_SW(GLinks)
      Jrn_Add("SDK_Space")

      ' Création des Noeuds du graphe. On passe de GLinks à Graph
      Solver_NLC.Graph_Build(GLinks)
      Solver_NLC.Graph_Display_Light()
      GRslt.Nb_Noeuds = Solver_NLC.Graph.Count

      'Titre = Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte
      'Texte = "le graphe comporte " & Solver_NLC.Graph.Count & " nœuds. "
      'Texte &= vbCrLf & "Entrez le nœud à traiter pour afficher les chemins :"
      'Noeud_NLC = InputBox(Texte, Titre)
      'If Noeud_NLC Is "" Then Exit Sub 'Cancel enfoncé
      AllPaths_NLC_Build(U_temp)

      'If GRslt.Productivité Then
      '  For Each excl As GCel_Excl_Cls In GRslt.CelExcl
      '    Jrn_Add(, {"NLC : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & excl.Cel & " - Unités : " & String.Join(", ", excl.Exc.Select(Function(c) U_Coord(c)))})
      '  Next excl

      '  For Each hs As Tuple(Of Integer, String) In GRslt.CelExcl_hs
      '    Jrn_Add_Yellow("NLC : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & hs.Item1 & " - Unité(s) : " & hs.Item2)
      '  Next hs
      '  Exit For

      'End If
    Next
    GRslt.Productivité = True
    GRslt.Nb_Liens = GLinks.Count
    GRslt_Display()
    Frm_SDK.Invalidate()
  End Sub

  Public Sub AllPaths_NLC_Build(U_temp(,) As String)
    Solver_NLC.AllPaths.Clear()

    Dim visited As HashSet(Of GLink_Cls)
    Dim path As List(Of GLink_Cls)

    For Each node As Integer In Solver_NLC.Graph.Keys
      Jrn_Add_Red("NLC : Recherche de chemins à partir du nœud " & node.ToString().PadRight(2) & " " & U_Coord(node))
      visited = New HashSet(Of GLink_Cls)
      path = New List(Of GLink_Cls)
      DFS_NLC(U_temp, node, Nothing, visited, path)
    Next
  End Sub

  Private Sub DFS_NLC(U_temp(,) As String, current As Integer,
                    lastType As String,
                    visited As HashSet(Of GLink_Cls),
                    path As List(Of GLink_Cls))

    Dim e As Edge
    Dim ln As GLink_Cls
    Dim nextNode As Integer

    For Each e In Solver_NLC.Graph(current)

      ln = e.Link
      nextNode = e.Neighbor

      If visited.Contains(ln) Then Continue For
      If lastType IsNot Nothing AndAlso ln.Type = lastType Then Continue For

      visited.Add(ln)
      path.Add(ln)

      If Path_NLC_Is_Productive(U_temp, path, path(0).Cel(0), nextNode) Then
        Solver_NLC.AllPaths.Add(New List(Of GLink_Cls)(path))
        Exit Sub
      End If

      DFS_NLC(U_temp, nextNode, ln.Type, visited, path)

      path.RemoveAt(path.Count - 1)
      visited.Remove(ln)
    Next
  End Sub

  Private Function Path_NLC_Is_Productive(U_temp(,) As String, path As List(Of GLink_Cls),
                                   startCel As Integer,
                                   currentCel As Integer) As Boolean

    Dim isLoop As Boolean = (path.Count >= 3 AndAlso currentCel = startCel)
    If Not isLoop Then Return False

    Dim evenLength As Boolean = ((path.Count Mod 2) = 0)
    If Not evenLength Then Return False

    Dim candidat As String = GRslt.Candidat(0)
    Dim weakLinks As New List(Of GLink_Cls)
    Dim ln As GLink_Cls

    For Each ln In path
      If ln.Type = "Weak" Then weakLinks.Add(ln)
    Next

    If weakLinks.Count = 0 Then Return False

    Dim i As Integer
    Dim celA As Integer, celB As Integer

    For Each ln In weakLinks
      celA = ln.Cel(0)
      celB = ln.Cel(1)

      For i = 0 To 80
        If i = celA OrElse i = celB Then Continue For
        If Not U_temp(i, 3).Contains(candidat) Then Continue For

        If Is_Vu(i, celA) AndAlso Is_Vu(i, celB) Then
          Dim key As Tuple(Of Integer, String) = Tuple.Create(i, candidat)
          If Not GRslt.CelExcl_hs.Contains(key) Then
            GRslt.CelExcl.Add(New GCel_Excl_Cls With {.Cel = i, .Cdd = candidat, .Exc = {celA, celB}})
            GRslt.CelExcl_hs.Add(key)
          End If
        End If
      Next
    Next

    If GRslt.CelExcl.Count > 0 Then
      GRslt.Productivité = True
      Jrn_Add(, {"LNC : Exclusion(s) détectée(s) pour le candidat " & candidat & " : " & GRslt.CelExcl.Count.ToString()})
      Return True
    End If

    Return False
  End Function

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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

End Module
