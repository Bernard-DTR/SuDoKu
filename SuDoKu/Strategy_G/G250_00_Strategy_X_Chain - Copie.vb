Option Strict On
Option Explicit On
'-------------------------------------------------------------------------------------------
'Stg_Profondeur: UOBTxysjzkq
'....5..8...7.293..1...74..25.91.2..4....4.........816.718....2...........6......7
'...6..51........3937..9......45....1.6...87..2..731..668..........1...........8.5
'..3...9.2..5.9..3.6.2..3..8....5....3.126..9....3.851641.....2......5.........6..
'-------------------------------------------------------------------------------------------
Module G250_Strategy_X_Chain
  Public Sub Strategy_GCx(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})
    GRslt_Init()
    For Cdd As Integer = 1 To 9
      Dim Candidat As String = Cdd.ToString()
      If GRslt.Productivité Then Exit For
      GRslt_Init()
      GRslt.Candidat(0) = Candidat
      If Xap Then Jrn_Add(, {"Traitement du Candidat : " & Candidat})

      ' 20 Création des liens
      GLinks_Build(U_temp, Candidat)
      GLinks_Exclude_Doubles()
      GLinks_OrderBy()
      If Xap Then GLinks_Display()

      GRslt.Nb_Liens = GLinks.Count
      'For Each gLink As GLink_Cls In GLinks
      '  Dim Current As New List(Of GLink_Cls)
      '  Current.Add(gLink)

      '  Array.Clear(U_Road, 0, 81)
      '  U_Road(gLink.Cel(0)) = True
      '  U_Road(gLink.Cel(1)) = True

      '  If DFS_XChain(GLinks, Current, U_Road, U_temp) Then
      '    Exit For
      '  End If

      'Next

      For Each gLink As GLink_Cls In GLinks

        ' Orientation 1 : telle quelle
        Dim Current1 As New List(Of GLink_Cls)
        Current1.Add(gLink)

        Array.Clear(U_Road, 0, 81)
        U_Road(gLink.Cel(0)) = True
        U_Road(gLink.Cel(1)) = True

        If DFS_XChain(GLinks, Current1, U_Road, U_temp) Then
          Exit For
        End If

        ' Orientation 2 : lien inversé
        Dim gLinkInv As GLink_Cls = ReverseLink(gLink)

        Dim Current2 As New List(Of GLink_Cls)
        Current2.Add(gLinkInv)

        Array.Clear(U_Road, 0, 81)
        U_Road(gLinkInv.Cel(0)) = True
        U_Road(gLinkInv.Cel(1)) = True

        If DFS_XChain(GLinks, Current2, U_Road, U_temp) Then
          Exit For
        End If

      Next



    Next
    Stratégies_G_End()
  End Sub

  Public Function ReverseLink(ByVal L As GLink_Cls) As GLink_Cls
    Dim R As New GLink_Cls With {
      .Cel = New Integer() {L.Cel(1), L.Cel(0)},
      .Cdd = New String() {L.Cdd(0), L.Cdd(1), L.Cdd(2), L.Cdd(3), L.Cdd(4)},
      .Type = L.Type,
      .Unité = L.Unité}
    Return R
  End Function

  Public Function IsCompatible(ByVal A As GLink_Cls,
                               ByVal B As GLink_Cls,
                               ByVal U_temp(,) As String) As Boolean

    If A.Cdd(0) <> B.Cdd(0) Then Return False
    If A.Cel(1) = B.Cel(0) Then Return False

    Dim unité As String() = Is_SameUnité(A.Cel(1), B.Cel(0))
    If unité(0) = "#" Then Return False

    Dim typeRéel As String = Is_LinkType(U_temp, A.Cel(1), B.Cel(0), A.Cdd(0))
    If typeRéel = "#" Then Return False

    Return True
  End Function

  Private Function ExploreLink(ByVal L As GLink_Cls,
                               ByVal AllLinks As List(Of GLink_Cls),
                               ByVal Current As List(Of GLink_Cls),
                               ByVal U_Road As Boolean(),
                               ByVal U_temp(,) As String) As Boolean

    Dim last As GLink_Cls = Current(Current.Count - 1)

    If Not IsCompatible(last, L, U_temp) Then Return False

    Dim c0 As Integer = L.Cel(0)
    Dim c1 As Integer = L.Cel(1)

    If U_Road(c0) OrElse U_Road(c1) Then Return False

    Current.Add(L)
    U_Road(c0) = True
    U_Road(c1) = True

    If DFS_XChain(AllLinks, Current, U_Road, U_temp) Then
      Return True
    End If

    Current.RemoveAt(Current.Count - 1)
    U_Road(c0) = False
    U_Road(c1) = False

    Return False
  End Function

  Public Function DFS_XChain(ByVal AllLinks As List(Of GLink_Cls),
                             ByVal Current As List(Of GLink_Cls),
                             ByVal U_Road As Boolean(),
                             ByVal U_temp(,) As String) As Boolean
    If Current.Count >= 3 Then
      If IsProductive(Current, U_temp) Then
        GRslt_RoadRight_Copie(Current, U_temp)
        GRslt.Productivité = True
        Return True
      End If
    End If

    For Each L0 As GLink_Cls In AllLinks
      If ExploreLink(L0, AllLinks, Current, U_Road, U_temp) Then
        Return True
      End If

      Dim Linv As GLink_Cls = ReverseLink(L0)
      If ExploreLink(Linv, AllLinks, Current, U_Road, U_temp) Then
        Return True
      End If
    Next

    Return False
  End Function

  Sub GRslt_RoadRight_Copie(ByVal Current As List(Of GLink_Cls), ByVal U_temp(,) As String)
    Dim Link_Numéro As Integer = 0
    Dim Link_Déb_Ec As Integer = -1
    Dim Link_Fin_Ec As Integer = -1
    Dim Link_Déb_Prv As Integer = -1, Link_Fin_Prv As Integer = -1

    For Each gLink As GLink_Cls In Current
      Link_Numéro += 1
      If Link_Numéro = 1 Then
        ' Premier lien
        Link_Déb_Ec = gLink.Cel(0) : Link_Fin_Ec = gLink.Cel(1)
        ' Ajout préventif du lien Fort dans le Chemin correct
        GRslt.RoadRight.Add(New GLink_Cls With {.Cel = New Integer() {gLink.Cel(0), gLink.Cel(1)},
                                                .Cdd = New String() {gLink.Cdd(0), gLink.Cdd(1), gLink.Cdd(2), gLink.Cdd(3), "0"},
                                                .Type = gLink.Type, .Unité = gLink.Unité})
      Else
        ' Liens suivants
        Link_Déb_Ec = gLink.Cel(0) : Link_Fin_Ec = gLink.Cel(1)
        Dim Unité As String() = Is_SameUnité(Link_Fin_Prv, Link_Déb_Ec)
        If Unité(0) = "#" Then Exit For
        'Ajout du lien faible, Quel est vraiment le type de lien ?  
        Dim Type_Réel As String
        Type_Réel = Is_LinkType(U_temp, Link_Fin_Prv, Link_Déb_Ec, gLink.Cdd(0))
        GRslt.RoadRight.Add(New GLink_Cls With {.Cel = New Integer() {Link_Fin_Prv, Link_Déb_Ec},
                                                .Cdd = New String() {gLink.Cdd(0), gLink.Cdd(1), gLink.Cdd(2), gLink.Cdd(3), "0"},
                                                .Type = "W", .Unité = Unité(0) & Unité(1) & Type_Réel})
        ' Ajout du lien Fort
        GRslt.RoadRight.Add(New GLink_Cls With {.Cel = New Integer() {Link_Déb_Ec, Link_Fin_Ec},
                                                .Cdd = New String() {gLink.Cdd(0), gLink.Cdd(1), gLink.Cdd(2), gLink.Cdd(3), "0"},
                                                .Type = gLink.Type, .Unité = gLink.Unité})
      End If

      Link_Déb_Prv = Link_Déb_Ec : Link_Fin_Prv = Link_Fin_Ec
    Next gLink
  End Sub

  Public Function IsProductive(ByVal Road As List(Of GLink_Cls),
                               ByVal U_temp(,) As String) As Boolean

    Dim celDeb As Integer = Road(0).Cel(0)
    Dim celFin As Integer = Road(Road.Count - 1).Cel(1)
    Dim candidat As String = Road(0).Cdd(0)

    GRslt.CelExcl.Clear()

    Dim nb As Integer = Candidats_Exclure_GCx(U_temp, candidat, celDeb, celFin)
    Return nb > 0
  End Function

  Public Function Candidats_Exclure_GCx(U_temp(,) As String,
                                        Candidat As String,
                                        Road_CelDéb As Integer,
                                        Road_CelFin As Integer) As Integer

    For i As Integer = 0 To 80
      If Not U_temp(i, 3).Contains(Candidat) Then Continue For
      If i = Road_CelDéb Or i = Road_CelFin Then Continue For

      If Is_Vu(i, Road_CelDéb) AndAlso Is_Vu(i, Road_CelFin) Then
        GRslt.CelExcl.Add(New GCel_Excl_Cls With {
          .Cel = i,
          .Cdd = Candidat,
          .Exc = {Road_CelDéb, Road_CelFin}
        })
      End If
    Next

    Return GRslt.CelExcl.Count
  End Function
End Module