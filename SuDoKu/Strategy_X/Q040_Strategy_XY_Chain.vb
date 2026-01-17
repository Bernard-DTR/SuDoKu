Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
'Mercredi 23/07/2025
'Stratégie de XY-Xhain
'Préfixe XCy, Cy
'.52......8....4.12.1.7...3..3...94569...7....1.....9.....5...6.3...615...4.8.....
'-------------------------------------------------------------------------------------------
Friend Module Q040_Strategy_XY_Chain


  Public Sub Strategy_XCy(U_temp(,) As String)
    ' Stratégie XY-Chain
    If Xap Then Jrn_Add(, {Procédure_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "XCy"
    XRslt_Init()

    ' 2 Inventaire des Cellules Bivalues triées sur les candidats 
    GCels.Clear()
    XCels_List_Get(U_temp, 2)
    If Xap Then XCels_List_Display()

    ' 3 Création des liens forts, il faut au moins 2 cellules bivalues pour faire un lien
    XLinks_List.Clear()
    XLinks_List_Generate_XCy()
    If Xap Then XLinks_List_Display("1")

    ' 4 Création des chemins
    XAllRoads_List.Clear()

    For Each Link As XLink_Cls In XLinks_List
      'Construnction de tous les chemins dans XAllRoads_List
      Dim Cell_F As Integer = Link.Cel(0)
      Dim cdd() As String = {Link.Cdd(0), Link.Cdd(1)}
      For Each candidat As String In cdd
        Dim XRoad As New List(Of XLink_Cls)
        Dim Hachage As New HashSet(Of Integer) From {Cell_F}
        If XAllRoads_List.Count < XRoads_Max Then XRoads_Inventory_XRp(XRoad, Hachage, Cell_F, candidat, XAllRoads_List)
      Next candidat
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
      Dim Cdd_Y As String = "0"
      ' Pour un lien, on a CddA - Cdd(4) ---> Cdd(4) - CddB
      Dim CddA As String = "0"
      Dim CddB As String = "0"
      Dim CddB_Prv As String = "0"

      For Each Link As XLink_Cls In Road
        If XRslt.Productivité Then Exit For
        Link_Numéro += 1

        ' Il faut distinguer les traitement du premier lien
        '                    les traitements des liens ( premier et suivants)
        '                    les refus de liens
        '                    les raisons du test de profitabilité
        ' Correction du 27/07/2025
        '   Il doit y avoir une alternace des candidats (ChatGPT seule indique cette règle et l'explique)

        ' CddA Premier candidat  <> Cdd_Y du lien
        If Link.Cdd(0) = Link.Cdd(4) Then CddA = Link.Cdd(0)
        If Link.Cdd(1) = Link.Cdd(4) Then CddA = Link.Cdd(1)
        ' CddB Deuxième candidat <> Cdd_Y du lien
        If Link.Cdd(2) = Link.Cdd(4) Then CddB = Link.Cdd(2)
        If Link.Cdd(3) = Link.Cdd(4) Then CddB = Link.Cdd(3)

        ' Ajout préventif du lien dans le Chemin correct
        XRslt.RoadRight.Add(Link)

        If Link_Numéro = 1 Then           '  Traitement pour le premier lien
          ' CddA Premier candidat  <> Cdd_Y du lien
          ' Au premier lien, on mémorise le candidat Y et la cellule de départ
          ' Link.Cdd(4) est le Cdd_Y du lien, donc Cdd_Y est l'autre candidat
          If Link.Cdd(0) = Link.Cdd(4) Then Cdd_Y = Link.Cdd(1)
          If Link.Cdd(1) = Link.Cdd(4) Then Cdd_Y = Link.Cdd(0)
          Road_CelDéb = Link.Cel(0)
        Else                               '  Traitement des liens suivants
          Road_CelFin = Link.Cel(1)

          If CddA = CddB_Prv Then Exit For '  Test de l'alternance des candidats
        End If                             '/ Traitement Premier Lien et des Liens suivants  
        CddB_Prv = CddA

        ' Test de la liaison
        '  x  1 à/partir du 3 èm lien (Il ne faut pas un nombre pair ou impair de liens)
        '       Copilot et ChatGPT donne 4 liens minimum. Mistral dit que 3 suffisent. Domino ne parle pas de limite !
        '  x  2 Road_CelFin doit contenir Cdd_Y (Road_CelDéb contient évidemment Cdd_Y)
        '  x  3 Road_CelDéb et Road_CelFin ne doivent pas contenir la même paire de candidats
        '  x  4 CddB ne doit pas être Cdd_Y (Il ne faut pas arriver sur Cdd_Y)

        If Link_Numéro >= 3 Then                                        ' Test 1
          If U_temp(Road_CelFin, 3).Contains(Cdd_Y) Then                ' Test 2
            If U_temp(Road_CelDéb, 3) <> U_temp(Road_CelFin, 3) Then    ' Test 3
              If CddB <> Cdd_Y Then                                     ' Test 4 
                If Candidats_Exclure_XCx_XCy(U_temp, Cdd_Y, Road_CelDéb, Road_CelFin) > 0 Then
                  XRslt.Candidat(0) = CStr(Cdd_Y)
                  XRslt.XRoads_Numéro = Road_Numéro
                  XRslt.XLinks_Nombre = Road.Count
                  XRslt.Productivité = True
                  If Xap Then XAllRoads_List_Display(Road_Numéro)
                End If
              End If
            End If
          End If
        End If
      Next Link
    Next Road

    Stratégies_G_End()
  End Sub

  Public Sub XLinks_List_Generate_XCy()
    'Il s'agit de 2 boucles permettant i vers j et j vers i
    '   i = j est éliminé
    '   i et j doivent être dans la même unité L, C, R et partager AU moins 1 candidat commun placé dans Cdd(4)
    '  (i et j dans une même Reg ne sont pas dans la même Row ou Col)

    For Each iXCel As GCel_Cls In GCels
      Dim iCdd As String() = {iXCel.Cdd(0), iXCel.Cdd(1)}

      For Each jXCel As GCel_Cls In GCels
        If iXCel.Cel = jXCel.Cel Then Continue For
        Dim Unité As String() = Is_SameUnité(iXCel.Cel, jXCel.Cel) ' {Code Row, Col, Reg, le numéro}
        If Mid$(Unité(0), 1, 1) = "#" Then Continue For

        Dim jCdd As String() = {jXCel.Cdd(0), jXCel.Cdd(1)}
        Dim Common As New List(Of String)
        Dim cdd As String
        For Each cdd In iCdd
          If jCdd.Contains(cdd) AndAlso Not Common.Contains(cdd) Then
            Common.Add(cdd)
          End If
        Next cdd

        ' Si le lien n'a qu'un candidat commun, alors Cdd(4) est ce candidat commun
        ' Si le lien a 2 candidats communs, alors il est créé 2 liens
        '    un lien avec Cdd(4) = Premier candidat
        '    un lien avec Cdd(4) = Second  candidat
        If Common.Count = 1 Then
          XLinks_List.Add(New XLink_Cls With {.Cel = New Integer() {iXCel.Cel, jXCel.Cel},
                                              .Cdd = New String() {iXCel.Cdd(0), iXCel.Cdd(1), jXCel.Cdd(0), jXCel.Cdd(1), Common(0)},
                                              .Type = "S",
                                              .Unité = Unité(0) & Unité(1)})
        ElseIf Common.Count = 2 Then
          For Each SharedC As String In Common
            XLinks_List.Add(New XLink_Cls With {.Cel = New Integer() {iXCel.Cel, jXCel.Cel},
                                                .Cdd = New String() {iXCel.Cdd(0), iXCel.Cdd(1), jXCel.Cdd(0), jXCel.Cdd(1), SharedC},
                                                .Type = "S",
                                                .Unité = Unité(0) & Unité(1)})
          Next SharedC
        End If
      Next jXCel
    Next iXCel
  End Sub

End Module