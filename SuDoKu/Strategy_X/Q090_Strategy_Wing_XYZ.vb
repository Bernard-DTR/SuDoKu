'-------------------------------------------------------------------------------------------
' Vendredi 22/08/2025
' Stratégie Wing_XYZ
' Préfixe WgZ, 
' Les stratégies sont ou seront X, XY, XYZ et W, soit WgX, WgY, WgZ et WgW
'6...527..4....35......1.............37.1.8.4...574.6.3..89..42.......3....2.86..9
'
'📌 Définition de la stratégie XYZ-Wing
'On a trois cases qui forment un motif :
'une cellule pivot qui contient trois candidats distincts (XYZ),
'deux ailes, qui sont des cellules bivalue partageant deux des trois candidats avec le pivot.
'
'Les conditions :
'Le pivot contient les trois candidats X, Y, Z.
'Chaque aile est bivalue et partage avec le pivot une paire de candidats 
'Aile 1 contient X et Y,
'Aile 2 contient Y et Z (ou toute permutation).
'Les trois cases se "voient" de façon à créer une relation logique :
'Le pivot voit chacune des deux ailes.
'Les ailes n'ont pas besoin de se voir entre elles.
'
'🔍 Principe logique
'Si le pivot est X, alors l'aile (X,Y) devient Y forcé → donc Z est exclu.
'Si le pivot est Y, alors les deux ailes se réduisent à X et Z → donc Z est encore exclu.
'Si le pivot est Z, alors directement Z est placé.
'👉 Dans tous les cas, le candidat Z peut être éliminé de toutes les cases qui voient à la fois les 2 ailes et le pivot.
'
' La stratégie XYZ_Wing, comme la stratégie W_Wing nécessite des structures plus larges en terme de candidats
' Adaptation des Classes
'   GCel_Cls  : Classe des Cellules avec ses 9 candidats Cdd As String() = Enumerable.Repeat("0", 9).ToArray()
'   XCels_List_Get(U_temp(,) As String, n As Integer)
'   XCels_Display
'   XLink_Cls : Classe structurant un lien avec 9 candidats Cdd As String() = Enumerable.Repeat("0", 9).ToArray() 
'   XLink_Str1  Utilisée dans X* et WgX, WgY
'   XLink_Str2  Utilisée dans WgZ
'-------------------------------------------------------------------------------------------
Friend Module Q090_Strategy_Wing_XYZ

  Public Sub Strategy_WgZ(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "WgZ" 
    XRslt_Init()

    ' 2 Inventaire des Cellules Trivalues 
    GCels.Clear()
    XCels_List_Get(U_temp, 3)
    If Xap Then XCels_List_Display()

    'If Plcy_XAllRoads_List_Rnd Then GCels = GCels.OrderBy(Function(x) RdX.Next()).ToList()
    '- OrderBy(Function(x) rnd.Next()) trie chaque élément selon un nombre aléatoire généré.
    '- .ToList() recrée la liste avec cet ordre mélangé.

    ' 3 Pour chaque cellule XYZ, on a les Permutations suivantes des 3 candidats 
    '   GCel_Cls Cel , cdd(0), Cdd(1), Cdd(2)
    '   X   XY et XZ
    '   Y   YX et YZ
    '   Z   ZX et ZY

    ' 3 Parcours des cellules trivaluées (pivot)
    For Each Cell_TV As GCel_Cls In GCels
      Dim Permutations As (Integer, Integer)() = {(1, 2), (0, 2), (1, 0), (2, 0), (0, 1), (2, 1)}
      For i As Integer = 0 To Permutations.Length - 1 Step 2
        XLinks_List.Clear()
        XLinks_List_Generate_WgZ(U_temp, Cell_TV.Cel, Cell_TV.Cdd(Permutations(i).Item1), Cell_TV.Cdd(Permutations(i).Item2))
        XLinks_List_Generate_WgZ(U_temp, Cell_TV.Cel, Cell_TV.Cdd(Permutations(i + 1).Item1), Cell_TV.Cdd(Permutations(i + 1).Item2))

        If Xap Then XLinks_List_Display("2")
        Candidats_Exclure_WgZ(U_temp)
        If XRslt.Productivité Then Exit For
      Next
      If XRslt.Productivité Then Exit For
    Next Cell_TV

    Stratégies_G_End()
  End Sub

  Public Sub XLinks_List_Generate_WgZ(U_temp(,) As String, CelXYZ As Integer, CddA As String, CddB As String)
    ' On recherche les Cellules à 2 candidats (Cdda, Cddb)  dans la Row, Col, Reg 
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) <> 2 Then Continue For
      Dim Unité As String() = Is_SameUnité(CelXYZ, i)
      If Unité(0) = "#" Then Continue For
      Dim Candidats As String = Cnddts_Blancs
      Mid$(Candidats, CInt(CddA), 1) = CddA
      Mid$(Candidats, CInt(CddB), 1) = CddB
      If U_temp(i, 3) = Candidats Then
        Dim CddXYZ As String = U_temp(CelXYZ, 3).Replace(" ", "")
        Dim CddAB As String = U_temp(i, 3).Replace(" ", "")
        Dim Link As New XLink_Cls With {.Cel = {CelXYZ, i}, .Cdd = {CddXYZ(0), CddXYZ(1), CddXYZ(2), CddAB(0), CddAB(1), CddB}, .Type = "W", .Unité = Unité(0) & Unité(1)}
        XLinks_List.Add(Link)
        Exit For
      End If
    Next i
  End Sub

  Public Function Candidats_Exclure_WgZ(U_temp(,) As String) As Integer
    XRslt.Productivité = False
    XRslt.RoadRight.Clear()
    XRslt.CelExcl.Clear()
    If XLinks_List.Count <> 2 Then Return 0
    Dim Candidat As String = XLinks_List.First.Cdd(5)
    Dim CelXYZ As Integer = XLinks_List.First.Cel(0)
    Dim CelXY As Integer = XLinks_List.First.Cel(1)
    Dim CelXZ As Integer = XLinks_List.Last.Cel(1)
    If U_temp(CelXY, 3) = U_temp(CelXZ, 3) Then Return 0
    XRslt.RoadRight.Add(XLinks_List.First)
    XRslt.RoadRight.Add(XLinks_List.Last)
    For i As Integer = 0 To 80
      If i = CelXYZ Or i = CelXY Or i = CelXZ Then Continue For
      If Not U_temp(i, 3).Contains(Candidat) Then Continue For
      If Is_Vu_Tous(i, {CelXYZ, CelXY, CelXZ}) Then
        XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = i, .Cdd = Candidat, .Exc = {CelXY, CelXZ}})
      End If
    Next
    If XRslt.CelExcl.Count > 0 Then
      XRslt.Candidat = {Candidat, "0"}
      XRslt.Productivité = True
    End If
    Return XRslt.CelExcl.Count
  End Function

End Module