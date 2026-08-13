'-------------------------------------------------------------------------------------------
'Mardi 11/08/2026
' Stratégie Nice_Loop Continuous & Discontinuous, et AIC (Alternating Inference Chains)
'  
' Préfixe NLC
'5..9.........2.4.98..7...........6...2.......4.6315....1.67..42.....3.7..4.....86
'   Les stratégies NLC, NLD et AIC sont en famille 7 pour ne pas générer d'erreur dans la gestion du menu
'                                  et elles passeront en famille 4 !
'548.......3....8..1.......3.72..86........3......47..9.1.3...2.2569.........6.7..
'NLC : Exclusion(s) détectée(s) pour le candidat 3 : 3
'6..91.52......7.....................5....4.16..86..4.5.4....1..93.86......5.4138.
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
  '-------------------------------------------------------------------------------------------


  Public Sub Strategy_NLC(U_temp(,) As String)


    Dim cand As String = "2"
    GLinks_Strong.Clear()
    GLinks_Build_Strong(U_temp, cand)
    GLinks_Strong = GLinks_Strong.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList()
    GLinks_Display_SW(GLinks_Strong)
    Jrn_Add("SDK_Space")

    GLinks_Weak.Clear()
    GLinks_Build_Weak(U_temp, cand)
    GLinks_Weak = GLinks_Weak.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList()
    GLinks_Display_SW(GLinks_Weak)
    Jrn_Add("SDK_Space")

    GLinks.AddRange(GLinks_Strong)
    GLinks.AddRange(GLinks_Weak)
    GLinks = GLinks.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList()
    GLinks_Display_SW(GLinks)
    Jrn_Add("SDK_Space")



    Exit Sub

    Plcy_Strg = "NLC"

    Dim Candidat As String
    GLinks.Clear()

    For Cdd As Integer = 1 To 9
      If GRslt.Productivité Then Exit For

      Jrn_Add(, {"NLC : Analyse du candidat " & Cdd.ToString()})
      GRslt_Init()

      Candidat = Cdd.ToString()
      GRslt_Init()
      GRslt.Candidat(0) = Candidat


      GLinks_Strong.Clear()
      GLinks_Build_Strong(U_temp, Candidat)
      GLinks_Display_SW(GLinks_Strong)

      GLinks_Weak.Clear()
      GLinks_Build_Weak(U_temp, Candidat)
      GLinks_Display_SW(GLinks_Weak)



      GLinks.AddRange(GLinks_Strong)
      GLinks.AddRange(GLinks_Weak)
      GLinks_Display_SW(GLinks)
      'Dim Solver As New DFS_CNL()
      'Solver.Graph_Build_cnl(Lall)
      ''Solver.Graph_cnl_Display()
      'Solver.AllPaths_Build_cnl(U_temp)

      'If GcnlRslt.Productivité Then
      '  For Each excl As GCel_Excl_Cls In GcnlRslt.CelExcl
      '    Jrn_Add(, {"CNL : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & excl.Cel & " - Unités : " & String.Join(", ", excl.Exc.Select(Function(c) U_Coord(c)))})
      '  Next excl

      '  For Each hs As Tuple(Of Integer, String) In GcnlRslt.CelExcl_hs
      '    Jrn_Add_Yellow("CNL : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & hs.Item1 & " - Unité(s) : " & hs.Item2)
      '  Next hs
      '  Exit For

      'End If
    Next
    GLinks_Display_SW(GLinks)
    GRslt_Display()

  End Sub
  Public Sub GLinks_Build_Strong(U_temp(,) As String, candidat As String)
    'La procédure remplie la liste Public GLinks_Strong As New List(Of GLink_Cls) des liens forts pour le candidat donné

    Dim i As Integer, j As Integer
    Dim gLink_Unité As String = "#"

    ' --- Lignes ---
    For i = 0 To 8
      Dim cells As New List(Of Integer)
      For j = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count = 2 Then
        gLink_Unité = "Row" + (U_Row(cells(0)) + 1).ToString()
        GLinks_Strong.Add(New GLink_Cls With {.Cel = New Integer() {cells(0), cells(1)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Composition = "024"})

      End If
    Next

    ' --- Colonnes ---
    For j = 0 To 8
      Dim cells As New List(Of Integer)
      For i = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count = 2 Then
        gLink_Unité = "Col" + (U_Col(cells(0)) + 1).ToString()
        GLinks_Strong.Add(New GLink_Cls With {.Cel = New Integer() {cells(0), cells(1)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Composition = "024"})
      End If
    Next

    ' --- Régions ---
    For b As Integer = 0 To 8
      Dim cells As New List(Of Integer)
      Dim r0 As Integer = (b \ 3) * 3
      Dim c0 As Integer = (b Mod 3) * 3
      For i = 0 To 2
        For j = 0 To 2
          Dim idx As Integer = (r0 + i) * 9 + (c0 + j)
          If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
        Next
      Next
      If cells.Count = 2 Then
        gLink_Unité = "Reg" + (U_Reg(cells(0)) + 1).ToString()
        GLinks_Strong.Add(New GLink_Cls With {.Cel = New Integer() {cells(0), cells(1)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "S",
                                       .Unité = gLink_Unité, .Composition = "024"})
      End If
    Next

  End Sub
  Public Sub GLinks_Build_Weak(U_temp(,) As String, candidat As String)
    'La procédure remplie la liste Public GLinks_Weak As New List(Of GLink_Cls) des liens faibles pour le candidat donné

    Dim i As Integer, j As Integer
    Dim gLink_Unité As String = "#"

    ' --- Lignes ---
    For i = 0 To 8
      Dim cells As New List(Of Integer)
      For j = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b As Integer = a + 1 To cells.Count - 1
            gLink_Unité = "Row" + (U_Row(i) + 1).ToString()
            GLinks_Weak.Add(New GLink_Cls With {.Cel = New Integer() {cells(a), cells(b)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "W",
                                       .Unité = gLink_Unité, .Composition = "024"})
          Next
        Next
      End If
    Next

    ' --- Colonnes ---
    For j = 0 To 8
      Dim cells As New List(Of Integer)
      For i = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b As Integer = a + 1 To cells.Count - 1
            gLink_Unité = "Col" + (U_Col(j) + 1).ToString()
            GLinks_Weak.Add(New GLink_Cls With {.Cel = New Integer() {cells(a), cells(b)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "W",
                                       .Unité = gLink_Unité, .Composition = "024"})
          Next
        Next
      End If
    Next

    ' --- Régions ---
    For b As Integer = 0 To 8
      Dim cells As New List(Of Integer)
      Dim r0 As Integer = (b \ 3) * 3
      Dim c0 As Integer = (b Mod 3) * 3
      For i = 0 To 2
        For j = 0 To 2
          Dim idx As Integer = (r0 + i) * 9 + (c0 + j)
          If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
        Next
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b2 As Integer = a + 1 To cells.Count - 1
            gLink_Unité = "Reg" + (b + 1).ToString()
            GLinks_Weak.Add(New GLink_Cls With {.Cel = New Integer() {cells(a), cells(b2)},
                                       .Cdd = New String() {candidat, " ", candidat, " ", candidat}, .Type = "W",
                                       .Unité = gLink_Unité, .Composition = "024"})
          Next
        Next
      End If
    Next
  End Sub
  Public Sub GLinks_Display_SW(L As List(Of GLink_Cls))
    ' Affichage de La liste L Liste des liens forts ou faibles
    Jrn_Add(, {Proc_Name_Get() & " Affichage de L As List(Of GLink_Cls) : " & L.Count & " Lignes."})
    If L.Count <> 0 Then
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In L
        With gLink
          Nb += 1
          Dim S As String = .Cdd(4) & " " &
            U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ")" & " → " &
            U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
            " Lien " & .Type & "  Unité " & .Unité.PadRight(6) & " Comp " & .Composition &
            " Cellules n° " & CStr(.Cel(0)).PadLeft(2) & "-" & CStr(.Cel(1)).PadLeft(2)
          Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(2) & " " & S})
        End With
      Next gLink
    End If
    Jrn_Add("SDK_Space")
  End Sub

End Module
