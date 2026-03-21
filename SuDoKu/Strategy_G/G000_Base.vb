Option Strict On
Option Explicit On
Imports SuDoKu.DancingLink       ' Pour l'indication de mémoire

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Mise en place 19/12/2025
' Objectif      Module de base des stratégies à base de graphe
' Construction des liens Forts
' .26.....4....2..1...75.8....9.4......5.8....24..6...39.....1...9.5..4.2.7...8.6.5
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Friend Module G000_Base
  Public Class Edge
    Public Property Neighbor As Integer
  End Class

  Public Class DFS_Coloration

    ' Graphe: cellule → liste d'arêtes (voisin)
    'Public ReadOnly Graph As Dictionary(Of Integer, List(Of Edge)) = New Dictionary(Of Integer, List(Of Edge))()
    Public ReadOnly Graph As New Dictionary(Of Integer, List(Of Edge))()
    ' Liste de tous les chemins trouvés
    Public AllPaths As New List(Of List(Of Integer))

    ' -----------------------------------------------------------------------------------------------------------
    ' Construction du graphe et affichage          Graph
    ' Construction des chemins et affichage        AllPaths 
    ' -----------------------------------------------------------------------------------------------------------
    Public Sub Graph_Build(ByVal Glinks As List(Of GLink_Cls))
      Graph.Clear()

      For Each glink As GLink_Cls In Glinks
        Dim a As Integer = glink.Cel(0)
        Dim b As Integer = glink.Cel(1)

        If Not Graph.ContainsKey(a) Then Graph(a) = New List(Of Edge)()
        If Not Graph.ContainsKey(b) Then Graph(b) = New List(Of Edge)()

        Graph(a).Add(New Edge() With {.Neighbor = b})
        Graph(b).Add(New Edge() With {.Neighbor = a})
      Next

    End Sub

    Public Sub Graph_Display()
      Jrn_Add(, {Graph.Count & " entrée(s)."})
      Dim l As Integer = 0
      For Each kvp As KeyValuePair(Of Integer, List(Of Edge)) In Graph
        l += 1
        Dim edges As List(Of Edge) = Graph(kvp.Key)
        Dim sb As New System.Text.StringBuilder()
        For Each edge As Edge In edges
          sb.AppendFormat(" → (Cel {0} )", U_Coord(edge.Neighbor))
        Next
        Dim edgeCount As String = $" ({edges.Count})"
        Jrn_Add(, {$"{l,2} Cel {U_Coord(kvp.Key)}:{edgeCount}{sb}"})
      Next
    End Sub

    ' Fonction principale pour lancer l’exploration depuis chaque nœud
    Public Sub AllPaths_Build()
      AllPaths.Clear()
      For Each node As Integer In Graph.Keys
        Paths_Build(node, New HashSet(Of Integer)(), New List(Of Integer)())
      Next
    End Sub

    ' Fonction récursive pour construire les chemins
    Private Sub Paths_Build(current As Integer, visited As HashSet(Of Integer), path As List(Of Integer))
      ' Ajouter le nœud courant au chemin
      ' Un HashSet est une collection de données non ordonnée qui ne permet pas les doublons
      path.Add(current)
      visited.Add(current)

      ' Stocker une copie du chemin actuel dans la liste des chemins
      AllPaths.Add(New List(Of Integer)(path))

      ' Explorer les voisins
      For Each e As Edge In Graph(current)
        Dim nextNode As Integer = e.Neighbor
        If Not visited.Contains(nextNode) Then
          Paths_Build(nextNode, New HashSet(Of Integer)(visited), New List(Of Integer)(path))
        End If
      Next
    End Sub

    Public Sub AllPaths_Display()
      Jrn_Add(, {AllPaths.Count & " chemins."})
      Dim pathIndex As Integer = 0
      For Each path As List(Of Integer) In AllPaths
        pathIndex += 1
        Dim pathStrg As String = String.Join(" → ", path.ConvertAll(Function(n) U_Coord(n)))
        Jrn_Add(, {$"Path {pathIndex,3} ({path.Count,2}) : {pathStrg}"})
      Next
    End Sub
    Public Sub Node_To_Paths(Candidat As String)
      GAllRoads.Clear()

      For Each path As List(Of Integer) In AllPaths
        If path.Count < 4 Then Continue For
        Dim Road_List As New List(Of GLink_Cls)   ' liste pour ce chemin
        Dim node_prv As Integer = -1
        Dim node As Integer

        For Each node In path
          If node_prv <> -1 Then
            Road_List.Add(New GLink_Cls With {.Cel = New Integer() {node_prv, node},
                                            .Cdd = New String() {Candidat, "0", Candidat, "0", Candidat},
                                            .Type = "S", .Unité = Wh_Unité(node_prv, node), .Cdd_Composition = "024"})
          End If
          node_prv = node
        Next
        ' ajoute la liste complète de ce chemin
        GAllRoads.Add(Road_List)
      Next
      'Jrn_Add(, {GAllRoads.Count & " GAllRoads.Count."})
    End Sub

    Public Function Wh_Unité(ByRef cel1 As Integer, ByRef cel2 As Integer) As String
      If U_Row(cel1) = U_Row(cel2) Then Return "Row" & CStr(U_Row(cel1) + 1)
      If U_Col(cel1) = U_Col(cel2) Then Return "Col" & CStr(U_Col(cel1) + 1)
      If U_Reg(cel1) = U_Reg(cel2) Then Return "Reg" & CStr(U_Reg(cel1) + 1)
      Return "#"
    End Function
  End Class

#Region "GCels List des Cellules"
  Public Sub GCels_bv_Build(U_temp(,) As String)
    'Détection des Cellules bi_value  
    ' Les 2 candidats sont placés dans cdd(0) et Cdd(1) et ensuite triés
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = 2 Then
        Dim cdd As String() = U_temp(i, 3).Where(Function(ch As Char) ch >= "1"c AndAlso ch <= "9"c).Select(Function(ch As Char) ch.ToString()).ToArray()
        GCels.Add(New GCel_Cls With {.Cel = i, .Cdd = cdd})
      End If
    Next i
  End Sub

  Public Sub GCels_bv_OrderBy()
    ' Tri selon Cdd(0), puis Cdd(1)
    GCels = GCels.
            Where(Function(x) x.Cdd.Length > 1).  ' Vérifie qu’il y a au moins deux éléments
            OrderBy(Function(x) x.Cdd(0)).
            ThenBy(Function(x) x.Cdd(1)).ToList() ' Then: Alors, Puis
  End Sub
  Public Sub GCels_Display()
    ' Affichage de la liste  GCels comportant n candidats
    ' Display affiche la GCels de 1 à 9 candidats
    If GCels.Count = 0 Then Exit Sub
    Jrn_Add(, {"Affichage de GCels :  " & GCels.Count & " Lignes."})
    Dim Nb As Integer = 0
    For Each gCel As GCel_Cls In GCels
      Nb += 1
      With gCel
        Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & U_Coord(.Cel) & " Candidats : " & String.Join("-", .Cdd) & " Cellule " & CStr(.Cel).PadLeft(2)})
      End With
    Next gCel
  End Sub
#End Region

#Region "Les Liens Forts"
  Sub GLinks_Build(U_temp(,) As String, Candidat As String)
    ' Construire la liste des liens pour un candidat donné ou tous les candidats si Candidat = " "
    ' La construction des liens ne tient compte d'aucun résutat attendu. 
    GLinks.Clear()
    For Cdd As Integer = 1 To 9
      Dim Cdd_Str As String = Cdd.ToString()
      If Candidat = "0" Or Candidat = Cdd_Str Then
        GLinks_Build_Unité(U_temp, Cdd_Str, "Row", U_9CelRow) ' Analyse des rangées
        GLinks_Build_Unité(U_temp, Cdd_Str, "Col", U_9CelCol) ' Analyse des colonnes
        GLinks_Build_Unité(U_temp, Cdd_Str, "Reg", U_9CelReg) ' Analyse des régions
      End If
    Next Cdd
  End Sub
  Public Sub GLinks_Build_Unité(U_temp(,) As String, Candidat As String, Code_Unité As String, Grp()() As Integer)
    ' On recherche dans chaque Rangée, Colonne, Région les unités comportant seulement 2 candidats (Liens Forts)
    ' Il est possible de détecter un même lien dans une Rangée/Colonne et dans la Région.
    For lcr As Integer = 0 To 8
      Dim Cel1 As Integer = -1, Cel2 As Integer = -1
      Dim n As Integer = 0
      For Each Cel As Integer In Grp(lcr)
        If U_temp(Cel, 3).Contains(Candidat) Then  ' Rechercher les cellules contenant le candidat
          n += 1
          If n = 1 Then Cel1 = Cel
          If n = 2 Then Cel2 = Cel
        End If
      Next Cel
      If n = 2 Then  ' L'unité ne doit contenir que 2 candidats pour qu'un lien fort soit ajouté
        Dim gLink_Unité As String = "#"
        Select Case Code_Unité
          Case "Row" : gLink_Unité = Code_Unité + (U_Row(Cel1) + 1).ToString()
          Case "Col" : gLink_Unité = Code_Unité + (U_Col(Cel1) + 1).ToString()
          Case "Reg" : gLink_Unité = Code_Unité + (U_Reg(Cel1) + 1).ToString()
        End Select
        GLinks.Add(New GLink_Cls With {.Cel = New Integer() {Cel1, Cel2},
                                       .Cdd = New String() {Candidat, " ", Candidat, " ", Candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Cdd_Composition = "024"})
      End If
    Next lcr
  End Sub
  Public Sub GLinks_OrderBy()
    ' Tri de GLinks par ordre croissant sur Cel(0) et Cel(1)
    GLinks = GLinks.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList()
  End Sub
  Public Sub GLinks_Exclude_Isolés()
    ' La procédure élimine les liens indépendants
    If GLinks.Count = 0 Then Exit Sub
    Dim U_Glinks(80) As Integer

    For Each gLink As GLink_Cls In GLinks
      U_Glinks(gLink.Cel(0)) += 1
      U_Glinks(gLink.Cel(1)) += 1
    Next gLink

    Dim GLinks_ToKeep As New List(Of GLink_Cls)
    For Each gLink As GLink_Cls In GLinks
      If Not (U_Glinks(gLink.Cel(0)) = 1 And U_Glinks(gLink.Cel(1)) = 1) Then
        GLinks_ToKeep.Add(gLink)
      End If
    Next gLink

    GLinks = New List(Of GLink_Cls)(GLinks_ToKeep)
  End Sub
  Public Sub GLinks_Exclude_Doubles()
    'Jrn_Add(, {Proc_Name_Get() & " Affichage de GLinks : " & GLinks.Count & " Lignes."})
    'la procédure élimine les liens en double
    GLinks = GLinks.GroupBy(Function(gl) (gl.Cel(0), gl.Cel(1))) _
                   .Select(Function(g) g.First()) _
                   .ToList()
    ' (2,5) double de (5,2)
    GLinks = GLinks _
        .GroupBy(Function(gl)
                   Dim a As Integer = gl.Cel(0)
                   Dim b As Integer = gl.Cel(1)
                   If a < b Then Return (a, b) Else Return (b, a)
                 End Function) _
        .Select(Function(g) g.First()) _
        .ToList()
  End Sub
  Public Sub GLinks_Display()
    ' Affichage de La liste GLinks 
    If GLinks.Count <> 0 Then
      Jrn_Add(, {Proc_Name_Get() & " Affichage de GLinks : " & GLinks.Count & " Lignes."})
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GLinks
        With gLink
          Nb += 1
          Dim S As String = .Cdd(4) & " " &
            U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ")" & " → " &
            U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
            " Lien " & .Type & "  Unité " & .Unité.PadRight(6) & " Comp " & .Cdd_Composition &
            " Cellules n° " & CStr(.Cel(0)).PadLeft(2) & "-" & CStr(.Cel(1)).PadLeft(2)
          Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(2) & " " & S})
        End With
      Next gLink
    End If
    Jrn_Add("SDK_Space")
  End Sub
#End Region

#Region "GRslt: Les Résultats d'une stratégie"
  Public Sub GRslt_Init()
    With GRslt
      .Code = Plcy_Strg
      .Candidat = {"0", "0"}
      .Cellule = {-1, -1}
      .Nb_Liens = -1
      .Nb_Noeuds = -1
      .Nb_Paths = -1
      .Path_Number = -1
      .RoadRight = New List(Of GLink_Cls)
      .CelExcl = New List(Of GCel_Excl_Cls)
      .CelExcl_hs = New HashSet(Of Tuple(Of Integer, String))
      .Productivité = False
    End With
  End Sub
  Public Sub GRslt_Display()
    ' La procédure affiche la structure GRslt et VERIFIE si les candidats à enlever font partie de la Solution DL
    Dim Nb As Integer = 0

    Jrn_Add(, {"Affichage des données GRslt"})
    With GRslt
      Jrn_Add(, {"Code Stratégie " & .Code & ", " & Stg_Get(.Code).Texte})
      Jrn_Add(, {"Candidat       " & String.Join(", ", .Candidat)})
      Jrn_Add(, {"Nb Liens forts " & .Nb_Liens & " (Glinks.count)"})
      Jrn_Add(, {"Nb Noeuds      " & .Nb_Noeuds})
      Jrn_Add(, {"Nb Paths       " & .Nb_Paths})
      Jrn_Add(, {"Path_Number    " & .Path_Number})
      Jrn_Add(, {"Cellules       " & String.Join(", ", .Cellule.Select(Function(c) U_Coord(c)))})
      Jrn_Add(, {"Productivité   " & .Productivité.ToString()})
    End With
    If GRslt.RoadRight.Count <> 0 Then
      Jrn_Add(, {"Affichage de GRslt.RoadRight.Count : " & GRslt.RoadRight.Count & " tronçons."})
      Nb = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        With gLink
          Nb += 1
          Dim S As String = .Cdd(4) & " " &
            U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ")" & " → " &
            U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
            " Lien " & .Type & "  Unité " & .Unité.PadRight(6) & " Comp " & .Cdd_Composition &
            " Cellules n° " & CStr(.Cel(0)).PadLeft(2) & "↔" & CStr(.Cel(1)).PadLeft(2)
          Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(2) & " " & S})
        End With
      Next gLink
    End If

    If GRslt.CelExcl.Count > 0 Then
      Jrn_Add(, {"Affichage de GRslt.CelExcl : " & GRslt.CelExcl.Count & " Lignes."})
      Nb = 0
      For Each GCel As GCel_Excl_Cls In GRslt.CelExcl
        Nb += 1
        Dim key As Tuple(Of Integer, String) = Tuple.Create(GCel.Cel, GCel.Cdd)
        GRslt.CelExcl_hs.Add(key)
        Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & U_Coord(GCel.Cel) & " Candidat : " & GCel.Cdd & " Extrémités : " & U_Coord(GCel.Exc(0)) & " ←→ " & U_Coord(GCel.Exc(1))})
      Next GCel

      Nb = 0
      Dim S As String = ""
      'Jrn_Add_Yellow("soit : " & GRslt.CelExcl_hs.Count.ToString() & " candidat(s)")
      Jrn_Add(, {"soit : " & GRslt.CelExcl_hs.Count.ToString() & " candidat(s)"})
      For Each CelExcl_key As Tuple(Of Integer, String) In GRslt.CelExcl_hs
        Nb += 1
        Dim cellule As Integer = CelExcl_key.Item1
        Dim candidat As String = CelExcl_key.Item2
        S &= $"{CStr(Nb),2} {U_Coord(cellule)} Candidat : {candidat}{vbCrLf}"
      Next
      Jrn_Add(, {S})

      'XSolution est calculé dans Game_Load (Game_New_Game)
      Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U)
      Jrn_Add(, {"Dancing Link        : " & CStr(DL.Nb_Solution)})
      Select Case DL.Nb_Solution
        Case -1, 0
        Case 1
          Jrn_Add(, {"Dancing Link        : " & DL.DLCode})
          Jrn_Add(, {"DL_Solution(0) " & XSolution})
          For Each gCel As GCel_Excl_Cls In GRslt.CelExcl
            Nb += 1
            With gCel
              Dim DL_Cellule As Integer = .Cel
              If XSolution(DL_Cellule) = .Cdd Then
                Jrn_Add(, {"❗ La cellule " & U_Coord(DL_Cellule) & " doit prendre la valeur " & .Cdd & " !"}, "Erreur")
              End If
            End With
          Next gCel
        Case Else
          Jrn_Add(, {"Dancing Link        : " & DL.DLCode & " Solutions multiples."})
      End Select
    End If
  End Sub
#End Region

#Region "Divers"
#End Region

#Region "Autres"
  Public Function Pzzl_Slv_UO(U_temp(,) As String) As Boolean
    ' La procédure consiste à appliquer les stratégies CdU et CdO pour résoudre la grille U_Temp
    ' Retourne True si la grille est résolue, False sinon
    While Wh_Grid_Nb_Cellules_Remplies(U_temp) <> 81  ' Tant que la grille n'est pas remplie
      'slv_U += 1
      Dim nb_U As Integer = 0
      ' Stratégie CdU des candidats uniques
      For i As Integer = 0 To 80
        If U_temp(i, 2) = " " And Trim(U_temp(i, 3)).Length = 1 Then
          nb_U += 1
          Cdd_Placer(U_temp, i, Trim(U_temp(i, 3)))
        End If
      Next
      ' la stratégie CdO étant la plus simple, elle "tourne" tant qu'elle fait des placements
      If nb_U <> 0 Then Continue While

      'slv_O += 1
      Dim nb_O As Integer = 0
      ' Stratégie CdO des candidats onligatoires
      For i As Integer = 0 To 80
        If U_temp(i, 2) = " " AndAlso Wh_Cell_Nb_Candidats(U_temp, i) > 1 Then
          For cdd As Integer = 1 To 9
            If U_temp(i, 3).Contains(CStr(cdd)) Then
              Dim unites As Integer()() = New Integer()() {U_9CelRow(U_Row(i)), U_9CelCol(U_Col(i)), U_9CelReg(U_Reg(i))}
              For Each unite As Integer() In unites
                If Not Pzzl_Slv_O_LCR(U_temp, unite, i, CStr(cdd)) Then
                  nb_O += 1
                  Cdd_Placer(U_temp, i, CStr(cdd))
                  Exit For ' Sort de la boucle des unités
                End If
              Next
            End If
          Next
        End If
      Next
      If nb_O = 0 Then Exit While

    End While
    If Wh_Grid_Nb_Cellules_Remplies(U_temp) = 81 Then Return True Else Return False
  End Function
  Function Cdd_Placer(U_temp(,) As String, Cellule As Integer, Candidat As String) As Boolean
    ' Place le candidat dans la cellule et met à jour les autres cellules
    U_temp(Cellule, 2) = Candidat
    U_temp(Cellule, 3) = Cnddts_Blancs
    Cdd_Remove_Cell_Coll(U_temp, Cellule)
    Return True
  End Function
  Function Pzzl_Slv_O_LCR(U_temp(,) As String, Grp() As Integer, Cellule As Integer, Candidat As String) As Boolean
    For g As Integer = 0 To 8 ' Analyse de l'Unité
      ' Pas de traitement pour les cellules remplies, ni pour la cellule étudiée
      If U_temp(Grp(g), 2) <> " " Or Grp(g) = Cellule Then Continue For
      If U_temp(Grp(g), 3).Contains(Candidat) = True Then Return True
    Next g                    '/Analyse de l'Unité
    Return False
  End Function
#End Region
End Module