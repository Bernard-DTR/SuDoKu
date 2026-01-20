Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
'
' GCs Coloration Simple
' Mise en place 30/12/2025
'-------------------------------------------------------------------------------------------
Module G200_Strategy_Coloration_Simple

  Public Sub Strategy_GCs(U_Temp(,) As String)
    ' 1 Initialisation de GRslt avec Plcy_Strg = "GCs"
    Plcy_Strg = "GCs"
    Jrn_Add_Red(Stg_Get(Plcy_Strg).Code & " " & Stg_Get(Plcy_Strg).Texte)

    GRslt_Init()

    For Cdd As Integer = 1 To 9   ' Les candidats sont traités au fur et à mesure de 1 à 9
      Dim Candidat As String = Cdd.ToString()
      If GRslt.Productivité Then Exit For
      GRslt_Init()
      GRslt.Candidat(0) = Candidat
      If Xap Then Jrn_Add(, {"Traitement du Candidat : " & Cdd.ToString})

      ' 20 Création des Liens Forts, il faut seulement 2 candidats par unité pour faire un lien fort
      '    Les liens isolés sont exclus
      '    Les liens en double ne seront pas traités dans le graphe
      GLinks_Build(U_Temp, Candidat)
      GLinks_Exclude_Isolés()
      GLinks_OrderBy()
      If Xap Then GLinks_Display()
      GRslt.Nb_Liens = GLinks.Count

      ' 30 Création des Noeuds du graphe. On passe de GLinks à Graph
      Dim Solver As New DFS_Coloration()
      Solver.Graph_Build(GLinks)
      If Xap Then Solver.Graph_Display()
      GRslt.Nb_Noeuds = Solver.Graph.Count

      '  Liste du graphe et copie pour affichage des noeuds et des arêtes 

      Solver.AllPaths_Build()
      If Xap Then Solver.AllPaths_Display()
      Solver.Node_To_Paths(Candidat)
      GRslt.Nb_Paths = Solver.AllPaths.Count

      ' 40 Vérification des chemins
      If GRoads_Vérification_dfs(U_Temp, GAllRoads) Then Exit For
    Next Cdd

    Stratégies_G_End()
  End Sub
  Public Function GRoads_Vérification_dfs(U_temp(,) As String,
                                          gAllRoads As List(Of List(Of GLink_Cls))) As Boolean
    Dim Road_Numéro As Integer = 0
    Dim Candidat As String = GRslt.Candidat(0)

    For Each Road As List(Of GLink_Cls) In gAllRoads
      If GRslt.Productivité Then Exit For
      Road_Numéro += 1

      If Road_Vérification_dfs(U_temp, Road, Candidat, Road_Numéro) Then
        Exit For
      End If
    Next

    Return GRslt.Productivité
  End Function
  Private Function Road_Vérification_dfs(U_temp(,) As String,
                                            Road As List(Of GLink_Cls),
                                            Candidat As String,
                                            Road_Numéro As Integer) As Boolean

    Dim Link_Fin_Prv As Integer = -1
    Dim LienIndex As Integer = 0
    Dim Couleurs As New Dictionary(Of Integer, Boolean)

    GRslt.RoadRight.Clear()
    GRslt.CelExcl.Clear()
    GRslt.Productivité = False

    ' Construction du chemin
    For Each Link As GLink_Cls In Road
      LienIndex += 1

      ' Vérification de continuité
      If LienIndex > 1 AndAlso Link.Cel(0) <> Link_Fin_Prv Then Exit For

      ' Ajout du lien
      GRslt.RoadRight.Add(New GLink_Cls With {
          .Cel = New Integer() {Link.Cel(0), Link.Cel(1)},
          .Cdd = New String() {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), "0"},
          .Type = Link.Type,
          .Unité = Link.Unité,
          .Cdd_Composition = Link.Cdd_Composition
      })

      ' Attribution des couleurs (alternance)
      Dim couleurDéb As Boolean = (LienIndex Mod 2 = 1) ' impair = Green
      Couleurs(Link.Cel(0)) = couleurDéb
      Couleurs(Link.Cel(1)) = Not couleurDéb

      ' Détection cycle
      If Link.Cel(1) = Link_Fin_Prv Then
        Dim cyclePair As Boolean = (LienIndex Mod 2 = 0)
        If Not cyclePair Then
          ' Cycle impair = contradiction => wrap
          For Each kvp As KeyValuePair(Of Integer, Boolean) In Couleurs
            If kvp.Value Then
              GRslt.CelExcl.Add(New GCel_Excl_Cls With {.Cel = kvp.Key, .Cdd = Candidat})
            End If
          Next
          GRslt.Productivité = True
        End If
        Exit For
      End If

      Link_Fin_Prv = Link.Cel(1)
    Next

    ' Analyse des traps (cycle pair)
    If Not GRslt.Productivité AndAlso GRslt.RoadRight.Count >= 3 Then
      For Each celG As Integer In Couleurs.Where(Function(c) c.Value).Select(Function(c) c.Key)
        For Each celB As Integer In Couleurs.Where(Function(c) Not c.Value).Select(Function(c) c.Key)
          Dim i As Integer
          For i = 0 To 80
            If i = celG OrElse i = celB Then Continue For
            If Not U_temp(i, 3).Contains(Candidat) Then Continue For
            If Is_Vu(i, celG) AndAlso Is_Vu(i, celB) Then
              If Not GRslt.CelExcl.Any(Function(ex) ex.Cel = i AndAlso ex.Cdd = Candidat) Then
                GRslt.CelExcl.Add(New GCel_Excl_Cls With {
                    .Cel = i, .Cdd = Candidat, .Exc = {celB, celG}})
              End If
            End If
          Next
        Next
      Next

      If GRslt.CelExcl.Count > 0 Then
        GRslt.Path_Number = Road_Numéro
        GRslt.Productivité = True
        'If Xap Then XAllRoads_List_Display_New(Road_Numéro)
      End If
    End If

    Return GRslt.Productivité
  End Function

End Module