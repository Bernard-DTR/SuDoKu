'-------------------------------------------------------------------------------------------
' Mercredi 06/08/2025
' Stratégie de Remote_Pairs
' 2 Candidats identiques
' Préfixe XRp Rp
' la stratégie propose également les cellules ayant l'un au l'autre des candidats
'-------------------------------------------------------------------------------------------

Friend Module Q050_Remote_Pairs
  Public Sub Strategy_XRp(U_temp(,) As String)
    ' Stratégie Remote_Pairs
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "XRp"
    XRslt_Init()

    ' 2 Inventaire des Cellules Bivalues triées sur les candidats 
    GCels.Clear()
    XCels_List_Get_XRp(U_temp)
    If Xap Then XCels_List_Display()

    ' 3 Création des liens forts, il faut au moins 2 cellules bivalues pour faire un lien
    XLinks_List.Clear()
    XLinks_List_Generate_XRp()
    If Xap Then XLinks_List_Display("1")

    ' 4 Création des chemins
    XAllRoads_List.Clear()

    For Each Link As XLink_Cls In XLinks_List
      'Construnction de tous les chemins dans XAllRoads_List
      Dim Cell_F As Integer = Link.Cel(0)
      Dim Cdd() As String = {Link.Cdd(0), Link.Cdd(1)}
      For Each Candidat As String In Cdd
        Dim XRoad As New List(Of XLink_Cls)
        Dim Hachage As New HashSet(Of Integer) From {Cell_F}
        If XAllRoads_List.Count < XRoads_Max Then XRoads_Inventory_XRp(XRoad, Hachage, Cell_F, Candidat, XAllRoads_List)
      Next Candidat
    Next Link
    XRslt.XRoads_Nombre = XAllRoads_List.Count
    If Xap Then Jrn_Add(, {CStr(XAllRoads_List.Count) & " XAllRoads_List.Count"})
    'XAllRoads_List_Display(-1)

    Dim Road_Numéro As Integer = 0
    Dim Link_Numéro As Integer = 0

    For Each Road As List(Of XLink_Cls) In XAllRoads_List
      If XRslt.Productivité Then Exit For
      Road_Numéro += 1
      Link_Numéro = 0
      XRslt.RoadRight.Clear()
      Dim Road_CelDéb As Integer = -1
      Dim Road_CelFin As Integer = -1
      For i As Integer = 0 To 80 : U_Road(i) = False : Next i
      Dim Cdd() As String = {"0", "0"}
      XRslt.Productivité = False

      For Each Link As XLink_Cls In Road
        If XRslt.Productivité Then Exit For
        Link_Numéro += 1

        U_Road(Link.Cel(0)) = True
        U_Road(Link.Cel(1)) = True

        ' Traitement du premier lien : on mémorise les 2 candidats Cdd et la cellule de départ Road_CelDéb 
        If Link_Numéro = 1 Then
          Cdd(0) = Link.Cdd(0)
          Cdd(1) = Link.Cdd(1)
          Road_CelDéb = Link.Cel(0)
        End If

        Road_CelFin = Link.Cel(1)
        ' Ajout préventif du lien dans le Chemin correct
        XRslt.RoadRight.Add(New XLink_Cls With {.Cel = New Integer() {Link.Cel(0), Link.Cel(1)},
                                                .Cdd = New String() {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), "0"},
                                                .Type = Link.Type, .Unité = Link.Unité})

        ' Les règles pour un Remote_pairs
        ' Test de la liaison
        '  x  1 à/partir du 3 èm lien et un nombre impair de liens
        If Link_Numéro >= 3 And Link_Numéro Mod 2 <> 0 Then                                     ' Test 1
          If Candidats_Exclure_XRp(U_temp, Cdd, Road_CelDéb, Road_CelFin) > 0 Then
            XRslt.Candidat(0) = Cdd(0)
            XRslt.Candidat(1) = Cdd(1)
            XRslt.XRoads_Numéro = Road_Numéro
            XRslt.XLinks_Nombre = Road.Count
            XRslt.Productivité = True
            If Xap Then XAllRoads_List_Display(Road_Numéro)
          End If
        End If
      Next Link
    Next Road

    Stratégies_G_End()
  End Sub

  Public Function Candidats_Exclure_XRp(U_temp(,) As String, Cdd() As String, Road_CelDéb As Integer, Road_CelFin As Integer) As Integer
    ' La cellule doit être vue des 2 extrémités 
    '            ne doit pas être sur le chemin
    '            doit comporter les 2 candidats ou l'un des deux
    XRslt.CelExcl.Clear()
    For i As Integer = 0 To 80
      If Not Is_Vu(i, Road_CelDéb) OrElse Not Is_Vu(i, Road_CelFin) Then Continue For
      If U_Road(i) Then Continue For

      If U_temp(i, 3).Contains(Cdd(0)) Then
        XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = i, .Cdd = Cdd(0), .Exc = {Road_CelDéb, Road_CelFin}})
      End If
      If U_temp(i, 3).Contains(Cdd(1)) Then
        XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = i, .Cdd = Cdd(1), .Exc = {Road_CelDéb, Road_CelFin}})
      End If

    Next
    Return XRslt.CelExcl.Count
  End Function

  Public Sub XCels_List_Get_XRp(U_temp(,) As String)
    'Détection des Cellules bi_value et de leurs 2 candidats
    ' Les 2 candidats sont placés dans cdd(0) et Cdd(1) et ensuite triés
    'Jrn_Add(, {Proc_Name_Get()})
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = 2 Then
        Dim cdd As String() = U_temp(i, 3).Where(Function(ch As Char) ch >= "1"c AndAlso ch <= "9"c).Select(Function(ch As Char) ch.ToString()).ToArray()
        GCels.Add(New GCel_Cls With {.Cel = i, .Cdd = cdd})
      End If
    Next i
    ' Tri selon Cdd(0), puis Cdd(1)
    GCels = GCels.
        Where(Function(x) x.Cdd.Length > 1).  ' Vérifie qu’il y a au moins deux éléments
        OrderBy(Function(x) x.Cdd(0)).
        ThenBy(Function(x) x.Cdd(1)).ToList() ' Then: Alors, Puis
  End Sub

  Public Sub XLinks_List_Generate_XRp()
    ' Création des liens forts, il faut au moins 2 cellules bivalues pour faire un lien
    ' Avec la boucle iXCel et la boucle jXCel, les liens sont correctement faits a --> b et b --> a
    ' Cel()       comporte les 2 cellules
    ' Cdd(0 et 1) sont les candidats de cel(0) 
    ' Cdd(2 et 3) sont les candidats de cel(1) on a Cdd_i(0) = Cdd_j(0), Cdd_j(1) = Cdd_j(1), et Cdd(4) = 0
    '                                          soit Cdd(0)   = Cdd(2) et Cdd((1)  = Cdd(3)    et Cdd(4) = 0
    ' Cdd(4)      est le candidat commun inutilisé

    For Each iXCel As GCel_Cls In GCels
      Dim iCdd As String() = {iXCel.Cdd(0), iXCel.Cdd(1)}
      For Each jXCel As GCel_Cls In GCels
        Dim jCdd As String() = {jXCel.Cdd(0), jXCel.Cdd(1)}

        If iXCel.Cel = jXCel.Cel Then Continue For
        ' Il faut que les 2 candidats soient les mêmes et dans une même unité
        If iCdd(0) = jCdd(0) AndAlso iCdd(1) = jCdd(1) Then
          Dim Unité As String() = Is_SameUnité(iXCel.Cel, jXCel.Cel) ' {Code Row, Col, Reg, le numéro}
          If Mid$(Unité(0), 1, 1) <> "#" Then
            XLinks_List.Add(New XLink_Cls With {.Cel = New Integer() {iXCel.Cel, jXCel.Cel},
                                                .Cdd = New String() {iXCel.Cdd(0), iXCel.Cdd(1), jXCel.Cdd(0), jXCel.Cdd(1), "0"},
                                                .Type = "S",
                                                .Unité = Unité(0) & Unité(1)})
          End If
        End If

      Next jXCel
    Next iXCel
  End Sub


  Public Sub XRoads_Inventory_XRp(XRoad As List(Of XLink_Cls), Hachage As HashSet(Of Integer), Cell_F As Integer, Candidat As String, XAllRoads_List As List(Of List(Of XLink_Cls)))
    ' Exploration récursive
    If XAllRoads_List.Count > XRoads_Max Then Exit Sub
    For Each Link As XLink_Cls In XLinks_List
      If XAllRoads_List.Count > XRoads_Max Then Exit For

      If Link.Cel(0) <> Cell_F Then Continue For
      If Hachage.Contains(Link.Cel(1)) Then Continue For

      ' On vérifie simplement que les deux cellules du lien sont compatibles (déjà garanti à la construction)
      Dim RoadNew As New List(Of XLink_Cls)(XRoad) From {Link}
      Hachage.Add(Link.Cel(1))

      ' Vérifier si la cellule d’arrivée contient le candidat cible
      Dim Cdd_Cell_T() As String = {Link.Cdd(2), Link.Cdd(3)}
      If Cdd_Cell_T.Contains(Candidat) Then
        XAllRoads_List.Add(RoadNew)
      End If

      ' Continuer récursivement
      If XAllRoads_List.Count < XRoads_Max Then XRoads_Inventory_XRp(RoadNew, Hachage, Link.Cel(1), Candidat, XAllRoads_List)
      Hachage.Remove(Link.Cel(1))
    Next Link
  End Sub



End Module
