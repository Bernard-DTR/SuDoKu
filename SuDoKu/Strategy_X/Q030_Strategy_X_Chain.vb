'-------------------------------------------------------------------------------------------
'Préfixe XCx
'Mercredi 09/07/2025 Stratégie X-Chain
'  Un seul candidat
'  Codification S/W Lien fort-Strong / Lien faible-Weak 
'...6..51........3937..9......45....1.6...87..2..731..668..........1...........8.5
'..3...9.2..5.9..3.6.2..3..8....5....3.126..9....3.851641.....2......5.........6..
'-------------------------------------------------------------------------------------------
Friend Module Q030_Strategy_X_Chain

  Public Sub Strategy_XCx(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "XCx"
    XRslt_Init()

    ' 2 Inventaire des candidats par ordre décroissant
    XCdds_List.Clear()
    XCdds_List_Get(U_temp)
    If Xap Then XCdds_List_Display()

    For Each XCdd As XCdd_Cls In XCdds_List
      If XRslt.Productivité Then Exit For

      ' 10 Initialisation de XRslt avec Plcy_Strg = "XCx"
      XRslt_Init()
      XRslt.Candidat(0) = XCdd.Cdd
      If Xap Then Jrn_Add("SDK_Space")
      If Xap Then Jrn_Add(, {"Candidat traité : " & XCdd.Cdd & " ,Qté : " & XCdd.Nb})

      ' 20 Création des Liens Forts, il faut seulement 2 candidats par unité pour faire un lien fort
      XLinks_List.Clear()
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Row", U_9CelRow) ' Analyse des rangées
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Col", U_9CelCol) ' Analyse des colonnes
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Reg", U_9CelReg) ' Analyse des régions
      If Xap Then XLinks_List_Display("1")

      ' 30 Construction des chemins
      XAllRoads_List.Clear()
      XRoads_Inventory_XCs_XCx_XNl(XLinks_List, New List(Of XLink_Cls)(), XAllRoads_List)
      If Xap Then Jrn_Add(, {CStr(XAllRoads_List.Count) & " XAllRoads_List.Count"})
      XRslt.XRoads_Nombre = XAllRoads_List.Count

      ' 40 Vérification des chemins
      If XRoads_Vérification_XCx(U_temp, XAllRoads_List) Then Exit For

    Next XCdd

    Stratégies_G_End()
  End Sub

  Public Function XRoads_Vérification_XCx(U_temp(,) As String, XAllRoads_List As List(Of List(Of XLink_Cls))) As Boolean
    Dim U_Cx(0 To 80) As Boolean    ' Tableau des cellules à omettre pour éviter les chemins en boucle
    Dim Road_Numéro As Integer = 0
    Dim Link_Numéro As Integer

    For Each Road As List(Of XLink_Cls) In XAllRoads_List
      If XRslt.Productivité Then Exit For

      Road_Numéro += 1
      Link_Numéro = 0
      XRslt.RoadRight.Clear()
      Dim Road_CelDéb As Integer = -1
      Dim Road_CelFin As Integer = -1
      Dim Link_Déb_Ec As Integer = -1
      Dim Link_Fin_Ec As Integer = -1
      Dim Link_Déb_Prv As Integer = -1, Link_Fin_Prv As Integer = -1
      Dim Candidat As String = XRslt.Candidat(0)
      XRslt.Productivité = False

      For i As Integer = 0 To 80 : U_Cx(i) = False : Next i

      For Each Link As XLink_Cls In Road
        Link_Numéro += 1

        If Link_Numéro = 1 Then
          ' Premier lien
          Road_CelDéb = Link.Cel(0)
          Link_Déb_Ec = Link.Cel(0) : Link_Fin_Ec = Link.Cel(1)
          ' Ajout préventif du lien Fort dans le Chemin correct
          XRslt.RoadRight.Add(New XLink_Cls With {.Cel = New Integer() {Link.Cel(0), Link.Cel(1)},
                                                  .Cdd = New String() {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), "0"},
                                                  .Type = Link.Type, .Unité = Link.Unité})
        Else
          ' Liens suivants
          Road_CelFin = Link.Cel(1)
          Link_Déb_Ec = Link.Cel(0) : Link_Fin_Ec = Link.Cel(1)
          If Link_Fin_Ec = Link_Fin_Prv Then Exit For
          If U_Cx(Link_Déb_Ec) Or U_Cx(Link_Fin_Ec) Then Exit For

          Dim Unité As String() = Is_SameUnité(Link_Fin_Prv, Link_Déb_Ec)
          If Unité(0) = "#" Then Exit For
          ' Ajout du lien faible, Quel est vraiment le type de lien ?  
          Dim Type_Réel As String
          Type_Réel = Is_LinkType(U_temp, Link_Fin_Prv, Link_Déb_Ec, Link.Cdd(0))
          XRslt.RoadRight.Add(New XLink_Cls With {.Cel = New Integer() {Link_Fin_Prv, Link_Déb_Ec},
                                        .Cdd = New String() {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), "0"},
                                        .Type = "W",
                                        .Unité = Unité(0) & Unité(1) & Type_Réel})
          ' Ajout du lien Fort
          XRslt.RoadRight.Add(New XLink_Cls With {.Cel = New Integer() {Link_Déb_Ec, Link_Fin_Ec},
                                        .Cdd = New String() {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), "0"},
                                        .Type = Link.Type,
                                        .Unité = Link.Unité})
        End If
        U_Cx(Link_Déb_Ec) = True
        U_Cx(Link_Fin_Ec) = True

        Link_Déb_Prv = Link_Déb_Ec : Link_Fin_Prv = Link_Fin_Ec
        If Link_Numéro >= 3 AndAlso Link_Numéro Mod 2 <> 0 Then      ' à/p de 3 liens et un nombre de liens impairs 
          If Candidats_Exclure_XCx_XCy(U_temp, Candidat, Road_CelDéb, Road_CelFin) > 0 Then
            XRslt.XRoads_Numéro = Road_Numéro
            XRslt.XLinks_Nombre = Road.Count
            XRslt.Productivité = True
            If Xap Then XAllRoads_List_Display(Road_Numéro)
            Exit For
          End If
        End If
      Next Link
    Next Road
    Return XRslt.Productivité
  End Function

  Public Function Candidats_Exclure_XCx_XCy(U_temp(,) As String, Candidat As String, Road_CelDéb As Integer, Road_CelFin As Integer) As Integer
    For i As Integer = 0 To 80
      If Not U_temp(i, 3).Contains(Candidat) Then Continue For
      If i = Road_CelDéb Or i = Road_CelFin Then Continue For
      If Is_Vu(i, Road_CelDéb) AndAlso Is_Vu(i, Road_CelFin) Then
        XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = i, .Cdd = Candidat, .Exc = {Road_CelDéb, Road_CelFin}})
      End If
    Next i
    Return XRslt.CelExcl.Count
  End Function

  Public Sub XRoads_Inventory_XCs_XCx_XNl(XRoad As List(Of XLink_Cls),
                                          Road As List(Of XLink_Cls),
                                          XAllRoads_List As List(Of List(Of XLink_Cls)))
    ' La procédure établit TOUS les Chemins possibles à/partir des Liens trouvés
    ' Le nombre de liens peut être important, il est donc limité à XRoads_Max
    If XAllRoads_List.Count >= XRoads_Max Then Exit Sub

    ' Si tous les Lnks sont utilisés
    If XRoad.Count = 0 Then
      Dim CompletedRoad As List(Of XLink_Cls)
      CompletedRoad = New List(Of XLink_Cls)(Road)
      XAllRoads_List.Add(CompletedRoad) 'Ajout du chemin dans la liste des chemins
      Return
    End If

    Dim i As Integer, Link As XLink_Cls

    ' Parcourir chaque Link
    For i = 0 To XRoad.Count - 1

      Link = XRoad(i)
      ' Direction Item1 -> Item2
      Road.Add(Link)
      Dim RestantRoad As List(Of XLink_Cls) = XRoad.Where(Function(t, index) index <> i).ToList()
      If XAllRoads_List.Count < XRoads_Max Then XRoads_Inventory_XCs_XCx_XNl(RestantRoad, Road, XAllRoads_List)
      Road.RemoveAt(Road.Count - 1)
      ' Direction Item2 -> Item1 
      Road.Add(New XLink_Cls With {.Cel = New Integer() {Link.Cel(1), Link.Cel(0)},
                                   .Cdd = New String() {Link.Cdd(2), Link.Cdd(3), Link.Cdd(0), Link.Cdd(1), Link.Cdd(4)},
                                   .Type = "S",
                                   .Unité = Link.Unité})
      If XAllRoads_List.Count < XRoads_Max Then XRoads_Inventory_XCs_XCx_XNl(RestantRoad, Road, XAllRoads_List)
      Road.RemoveAt(Road.Count - 1)
    Next i
  End Sub

  Public Sub XLinks_List_Generate_XCs_XCx_XNl(U_temp(,) As String, Ccd As String, Code_LCR As String, GrpArray()() As Integer)
    ' On recherche dans chaque rangée, colonne, région les unités comportant seulement 2 candidats (Liens Forts)
    ' Contrairement à Remote Pairs, les liens sont établis uniquement dans un sens
    For Lcr As Integer = 0 To 8
      Dim Grp() As Integer = GrpArray(Lcr)
      Dim Cel1 As Integer = -1, Cel2 As Integer = -1
      Dim n As Integer = 0
      ' Combien de candidats comporte l'unité ?
      For Each Cel As Integer In Grp
        If U_temp(Cel, 3).Contains(CStr(Ccd)) Then      ' Rechercher les cellules contenant le candidat
          n += 1
          If n = 1 Then Cel1 = Cel
          If n = 2 Then Cel2 = Cel
        End If
      Next Cel
      ' L'unité ne doit contenir que 2 candidats pour qu'un lien fort soit ajouté
      If n = 2 Then
        Select Case Code_LCR
          Case "Row"
            XLinks_List.Add(New XLink_Cls With {.Cel = New Integer() {Cel1, Cel2},
                                              .Cdd = New String() {Ccd, "0", Ccd, "0", Ccd},
                                              .Type = "S",
                                              .Unité = "Row" & U_Row(Cel1) + 1})
          Case "Col"
            XLinks_List.Add(New XLink_Cls With {.Cel = New Integer() {Cel1, Cel2},
                                              .Cdd = New String() {Ccd, "0", Ccd, "0", Ccd},
                                              .Type = "S",
                                              .Unité = "Col" & U_Col(Cel1) + 1})
          Case "Reg"
            ' il ne faut pas que les 2 cellules soient dans la même ligne / Colonne
            ' l'analyse des lignes/colonne a déjà été faite auparavant
            If U_Row(Cel1) <> U_Row(Cel2) And U_Col(Cel1) <> U_Col(Cel2) Then
              XLinks_List.Add(New XLink_Cls With {.Cel = New Integer() {Cel1, Cel2},
                                                .Cdd = New String() {Ccd, "0", Ccd, "0", Ccd},
                                                .Type = "S",
                                                .Unité = "Reg" & U_Reg(Cel1) + 1})
            End If
        End Select
      End If
    Next Lcr
  End Sub
End Module
