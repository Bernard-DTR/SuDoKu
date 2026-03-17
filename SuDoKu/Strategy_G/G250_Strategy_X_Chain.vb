Option Strict On
Option Explicit On
Module G250_Strategy_X_Chain

  '-------------------------------------------------------------------------------------------
  'Stg_Profondeur: UOBTxysjzkq
  '....5..8...7.293..1...74..25.91.2..4....4.........816.718....2...........6......7
  '...6..51........3937..9......45....1.6...87..2..731..668..........1...........8.5
  '..3...9.2..5.9..3.6.2..3..8....5....3.126..9....3.851641.....2......5.........6..
  '-------------------------------------------------------------------------------------------
  '   Intégration de la procédure Candidats_Exclure_GCx       OK
  '   MaxRoads ne sert pas.                                   OK
  '   Passage de X à G
  '   XLinks_List     devient GLinks                          OK
  '   XLink_Cls       devient GLink_Cls                       OK
  '   Visited         est remplacé par Public U_Road          OK
  '   Result          est remplacé par Public GAllRoads As New List(Of List(Of GLink_Cls)) OK
  Public Sub Strategy_GCx(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de GRslt avec Plcy_Strg = "GCx"
    GRslt_Init()

    For Cdd As Integer = 1 To 9   ' Les candidats sont traités au fur et à mesure de 1 à 9
      Dim Candidat As String = Cdd.ToString()

      If Candidat <> "3" Then Continue For 'je teste pour le moment sur le candidat 3
      If GRslt.Productivité Then Exit For
      GRslt_Init()
      GRslt.Candidat(0) = Candidat
      If Xap Then Jrn_Add(, {"Traitement du Candidat : " & Cdd.ToString})

      ' 20 Création des Liens Forts, il faut seulement 2 candidats par unité pour faire un lien fort
      '    Les liens en double sont exclus
      GLinks_Build(U_temp, Candidat)
      GLinks_Exclude_Doubles()
      GLinks_OrderBy()
      If Xap Then GLinks_Display()
      GRslt.Nb_Liens = GLinks.Count

      For Each gLink As GLink_Cls In GLinks
        If gLink.Type <> "S" Then Continue For

        Dim Current As New List(Of GLink_Cls)
        Current.Add(gLink)

        Array.Clear(U_Road, 0, 81)                 ' For i As Integer = 0 To 80 : U_Road(i) = False : Next i
        U_Road(gLink.Cel(0)) = True
        U_Road(gLink.Cel(1)) = True

        DFS_XChain(GLinks, Current, U_Road, GAllRoads, U_temp)
      Next

      Jrn_Add_Orange("GAllRoads.Count " & GAllRoads.Count.ToString())
      GRslt_Display() 'Pour contrôle


    Next Cdd

    Stratégies_G_End()
  End Sub

  ' 1. Inversion d’un lien (on garde, ça peut servir)
  Public Function ReverseLink(ByVal L As GLink_Cls) As GLink_Cls
    Dim R As New GLink_Cls With {
        .Cel = New Integer() {L.Cel(1), L.Cel(0)},
        .Cdd = New String() {L.Cdd(0), L.Cdd(1), L.Cdd(2), L.Cdd(3), L.Cdd(4)},
        .Type = L.Type,
        .Unité = L.Unité
      }
    Return R
  End Function

  Public Function IsCompatible(ByVal A As GLink_Cls,
                             ByVal B As GLink_Cls,
                             ByVal U_temp(,) As String) As Boolean

    ' Même candidat
    If A.Cdd(0) <> B.Cdd(0) Then Return False

    ' Pas de lien faible entre une cellule et elle-même
    If A.Cel(1) = B.Cel(0) Then Return False

    ' Doivent voir l’une l’autre (même unité)
    Dim unité As String() = Is_SameUnité(A.Cel(1), B.Cel(0))
    If unité(0) = "#" Then Return False

    ' Vérifier que le lien faible est réellement possible
    Dim typeRéel As String = Is_LinkType(U_temp, A.Cel(1), B.Cel(0), A.Cdd(0))
    If typeRéel = "#" Then Return False

    Return True
  End Function
  ' 3. Extension du chemin
  Private Sub ExploreLink(ByVal L As GLink_Cls,
                        ByVal AllLinks As List(Of GLink_Cls),
                        ByVal Current As List(Of GLink_Cls),
                        ByVal U_Road As Boolean(),
                        ByVal GAllRoads As List(Of List(Of GLink_Cls)),
                        ByVal U_temp(,) As String)

    Dim last As GLink_Cls = Current(Current.Count - 1)

    If Not IsCompatible(last, L, U_temp) Then Exit Sub

    Dim c0 As Integer = L.Cel(0)
    Dim c1 As Integer = L.Cel(1)

    If U_Road(c0) OrElse U_Road(c1) Then Exit Sub

    Current.Add(L)
    U_Road(c0) = True
    U_Road(c1) = True

    DFS_XChain(AllLinks, Current, U_Road, GAllRoads, U_temp)

    Current.RemoveAt(Current.Count - 1)
    U_Road(c0) = False
    U_Road(c1) = False
  End Sub
  ' 4. DFS principale
  Public Sub DFS_XChain(ByVal AllLinks As List(Of GLink_Cls),
                      ByVal Current As List(Of GLink_Cls),
                      ByVal U_Road As Boolean(),
                      ByVal GAllRoads As List(Of List(Of GLink_Cls)),
                      ByVal U_temp(,) As String)


    ' À partir de 3 liens forts, on teste la productivité
    If Current.Count >= 3 Then
      If IsProductive(Current, U_temp) Then
        GAllRoads.Add(New List(Of GLink_Cls)(Current))
        Exit Sub
      End If
    End If

    Dim L0 As GLink_Cls
    For Each L0 In AllLinks

      ' Sens naturel
      ExploreLink(L0, AllLinks, Current, U_Road, GAllRoads, U_temp)

      ' Sens inversé (pour obtenir 70→62, 58→49, etc.)
      Dim Linv As GLink_Cls = ReverseLink(L0)
      ExploreLink(Linv, AllLinks, Current, U_Road, GAllRoads, U_temp)

    Next L0
  End Sub
  ' 5. Productivité : mêmes extrémités que ta version
  Public Function IsProductive(ByVal Road As List(Of GLink_Cls),
                               ByVal U_temp(,) As String) As Boolean

    Dim celDeb As Integer = Road(0).Cel(0)
    Dim celFin As Integer = Road(Road.Count - 1).Cel(1)
    Dim candidat As String = Road(0).Cdd(0)

    Return Candidats_Exclure_GCx(U_temp, candidat, celDeb, celFin) > 0
  End Function

  Public Function Candidats_Exclure_GCx(U_temp(,) As String, Candidat As String, Road_CelDéb As Integer, Road_CelFin As Integer) As Integer
    For i As Integer = 0 To 80
      If Not U_temp(i, 3).Contains(Candidat) Then Continue For
      If i = Road_CelDéb Or i = Road_CelFin Then Continue For
      If Is_Vu(i, Road_CelDéb) AndAlso Is_Vu(i, Road_CelFin) Then
        GRslt.CelExcl.Add(New GCel_Excl_Cls With {.Cel = i, .Cdd = Candidat, .Exc = {Road_CelDéb, Road_CelFin}})
        Jrn_Add_Yellow(Proc_Name_Get() & " " & U_Coord(i) & " " & Candidat)
      End If
    Next i
    Return GRslt.CelExcl.Count
  End Function

End Module