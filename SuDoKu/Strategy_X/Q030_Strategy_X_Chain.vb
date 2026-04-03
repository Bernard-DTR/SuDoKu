'-------------------------------------------------------------------------------------------
'Préfixe XCx
'Mercredi 09/07/2025 Stratégie X-Chain
'  Un seul candidat
'  Codification S/W Lien fort-Strong / Lien faible-Weak 
'-------------------------------------------------------------------------------------------
Friend Module Q030_Strategy_X_Chain

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
