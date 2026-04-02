Imports SuDoKu.DancingLink

'-------------------------------------------------------------------------------------------
'Mardi 08/07/2025
'Code commun aux Stratégies X
' 
'Lien fort   S Strong Link entre 2 candidats: Si l'un est faux, l'autre est obligatoirement vrai
'              Dans une cellule bivaluée
'              Dans une unité, si un candidat n'apparait que 2 fois   
'Lien faible W Weak Link: Si l'un est vrai, l'autre est faux
'              Dans une ligne, si 2 cellules ont le même candidat, si l'une le prend, l'autre ne le peut pas.
'-------------------------------------------------------------------------------------------
Friend Module Q000_Strategy_X

  Function Is_Vu(Cel1 As Integer, Cel2 As Integer) As Boolean
    ' Usage: If Is_Vu(i, Road_CelDéb) AndAlso Is_Vu(i, Road_CelFin) Then
    Dim Row1 As Integer = Cel1 \ 9, Col1 As Integer = Cel1 Mod 9
    Dim Row2 As Integer = Cel2 \ 9, Col2 As Integer = Cel2 Mod 9
    Dim Reg1 As Integer = ((Row1 \ 3) * 3) + (Col1 \ 3)
    Dim Reg2 As Integer = ((Row2 \ 3) * 3) + (Col2 \ 3)
    Return Row1 = Row2 OrElse Col1 = Col2 OrElse Reg1 = Reg2
  End Function

  Function Is_Vu_Tous(Cellule As Integer, Autres_Cellules As Integer()) As Boolean
    ' Usage: If Is_Vu_Tous(i, {CelXYZ, CelXY, CelXZ}) Then
    For Each Cel As Integer In Autres_Cellules
      If Not Is_Vu(Cellule, Cel) Then Return False
    Next Cel
    Return True
  End Function

  Private Function ExisteLienFaible(U_temp(,) As String, cel1 As Integer, cel2 As Integer, cdd As String) As Boolean
    Return Is_Vu(cel1, cel2) AndAlso
           U_temp(cel1, 3).Contains(cdd) AndAlso
           U_temp(cel2, 3).Contains(cdd)
  End Function
#Region "GCels"

  Public Sub XCels_List_Get(U_temp(,) As String, n As Integer)
    ' Détection des Cellules ayant n candidats
    ' La GCels s'adapte de 1 à 9 candidats
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = n Then
        ' Take(n) n'est pas nécessaire si Wh_Cell_Nb_Candidats(U_temp, i)
        Dim cdd_i As String() = U_temp(i, 3).Where(Function(ch As Char) Char.IsDigit(ch)).Select(Function(ch As Char) ch.ToString()).Take(n).ToArray()
        GCels.Add(New GCel_Cls With {.Cel = i, .Cdd = cdd_i})
      End If
    Next i
  End Sub

  Public Sub XCels_List_Display()
    ' Affichage de la liste  GCels comportant n candidats
    ' Display affiche la GCels de 1 à 9 candidats
    Jrn_Add(, {"Affichage de XWCels_List :  " & GCels.Count & " Lignes."})
    If GCels.Count = 0 Then Exit Sub
    Dim Nb As Integer = 0
    For Each XWCel As GCel_Cls In GCels
      Nb += 1
      With XWCel
        Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & U_Coord(.Cel) & " Candidats : " & String.Join("-", .Cdd) & " Cellule " & CStr(.Cel).PadLeft(2)})
      End With
    Next XWCel
  End Sub

  ' Pour enlever des cellules de la liste:
  ' Dim Exclusions As Integer() = {52, 64, 41, 43}
  ' GCels.RemoveAll(Function(x) Exclusions.Contains(x.Cel))
  ' Il est délicat d'enlever des éléments d'une liste.

#End Region

#Region "XLinks_List"

  Public Function XLink_Str_New(XLink As XLink_Cls) As String
    'Retourne une chaîne lisible de XLink_Cls en utilisant la composition
    Dim S As String = ""
    Try
      With XLink
        Select Case .Composition
          Case "024", "0123"
            ' Les stratégies X et Wing (sauf XYZ-Wing)
            'Cel()        comporte les 2 cellules
            'Cdd(0 et 1)  sont les candidats de cel(0)
            'Cdd(2 et 3)  sont les candidats de cel(1)
            'Cdd(4)       est le candidat commun
            S = .Cdd(4) & " " &
            U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ")" & " → " &
            U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
            " Lien " & .Type & "  Unité " & .Unité.PadRight(6) & " Comp " & .Composition &
            " N° Cellules " & CStr(.Cel(0)).PadLeft(2) & " " & CStr(.Cel(1)).PadLeft(2)
          Case Else
            ' Permet l'affichage de X_Chain
            ' Les stratégies X et Wing (sauf XYZ-Wing)
            'Cel()        comporte les 2 cellules
            'Cdd(0 et 1)  sont les candidats de cel(0)
            'Cdd(2 et 3)  sont les candidats de cel(1)
            'Cdd(4)       est le candidat commun
            S = .Cdd(4) & " " &
            U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ")" & " → " &
            U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
            " Lien " & .Type & "  Unité " & .Unité.PadRight(6) & " Comp " & .Composition &
            " N° Cellules " & CStr(.Cel(0)).PadLeft(2) & " " & CStr(.Cel(1)).PadLeft(2)
        End Select
      End With
    Catch ex As Exception
    End Try
    Return S
  End Function

  Public Function XLink_Str1(XLink As XLink_Cls) As String
    'Retourne une chaîne lisible de XLink_Cls
    Dim S As String = ""
    Try
      With XLink
        ' Les stratégies X et Wing (sauf XYZ-Wing)
        'Cel()        comporte les 2 cellules
        'Cdd(0 et 1)  sont les candidats de cel(0)
        'Cdd(2 et 3)  sont les candidats de cel(1)
        'Cdd(4)       est le candidat commun
        S = .Cdd(4) & " " &
      U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & ") " &
      U_Coord(.Cel(1)) & " (" & .Cdd(2) & "-" & .Cdd(3) & ") " &
      " Type-Lien " & .Type & " Unité " & .Unité.PadRight(6) &
      " N° Cellules " & CStr(.Cel(0)).PadLeft(2) & " " & CStr(.Cel(1)).PadLeft(2)
      End With
    Catch ex As Exception
    End Try
    Return S
  End Function

  Public Function XLink_Str2(XLink As XLink_Cls) As String
    'Retourne une chaîne lisible de XLink_Cls
    Dim S As String = ""
    Try
      With XLink
        '(Stratégie XYZ-Wing)
        'Cel()        comporte les 2 cellules
        'Cdd(0, 1, 2) sont les candidats de cel(0)
        'Cdd(3 et  4) sont les candidats de cel(1)
        'Cdd(5)       est le candidat commun
        S = .Cdd(5) & " " &
        U_Coord(.Cel(0)) & " (" & .Cdd(0) & "-" & .Cdd(1) & "-" & .Cdd(2) & ") " &
        U_Coord(.Cel(1)) & " (" & .Cdd(3) & "-" & .Cdd(4) & ") " &
        " Type-Lien " & .Type & " Unité " & .Unité.PadRight(6) &
        " N° Cellules " & CStr(.Cel(0)).PadLeft(2) & " " & CStr(.Cel(1)).PadLeft(2)
      End With
    Catch ex As Exception
    End Try
    Return S
  End Function

  Public Sub XLinks_List_Display(Type As String)
    ' Affichage de La liste XLinks_List
    Jrn_Add(, {Proc_Name_Get()})
    If XLinks_List.Count <> 0 Then
      Jrn_Add(, {"Affichage de XLinks_List : " & XLinks_List.Count & " Lignes, Type : " & Type})
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XLinks_List
        With Link
          Nb += 1
          Select Case Type
            Case "1" : Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & XLink_Str1(Link)})
            Case "2" : Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & XLink_Str2(Link)})
          End Select
        End With
      Next Link
    End If
  End Sub

  Public Sub XLinks_List_Display_New()
    ' Affichage de La liste XLinks_List
    If XLinks_List.Count <> 0 Then
      Jrn_Add(, {Proc_Name_Get() & " Affichage de XLinks_List : " & XLinks_List.Count & " Lignes."})
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XLinks_List
        Nb += 1 : Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & XLink_Str_New(Link)})
      Next Link
    End If
    Jrn_Add("SDK_Space")
  End Sub

#End Region

#Region "XCdds_List"
  Public Sub XCdds_List_Get(U_temp(,) As String)
    ' La fonction dénombre les candidats existants de la grille et les place en ordre décroissant
    XCdds_List.Clear()
    For Cdd As Integer = 1 To 9
      Dim Cdd_Str As String = Cdd.ToString()
      Dim Nb As Integer = Enumerable.Range(0, 80) _
                         .Count(Function(i) U_temp(i, 2) = " " AndAlso U_temp(i, 3).Contains(Cdd_Str))
      If Nb <> 0 Then XCdds_List.Add(New XCdd_Cls With {.Cdd = Cdd_Str, .Nb = Nb})
    Next Cdd

    'XCdds_List = XCdds_List.OrderBy(Function(s) s.Nb).ToList() 'par ordre croissant
    XCdds_List = XCdds_List.OrderBy(Function(s) -s.Nb).ToList() 'par ordre décroissant
  End Sub

  Public Sub XCdds_List_Display()
    ' Affichage de La liste XCdds_List
    Jrn_Add(, {Proc_Name_Get()})
    If XCdds_List.Count <> 0 Then
      Jrn_Add(, {"Affichage de XCdds_List : " & XCdds_List.Count & " Lignes."})
      Dim Nb As Integer = 0
      For Each XCdd As XCdd_Cls In XCdds_List
        With XCdd
          Nb += 1
          Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & " Candidat " & XCdd.Cdd & " Nombre : " & CStr(XCdd.Nb).PadLeft(2)})
        End With
      Next XCdd
    End If
  End Sub
#End Region

#Region "XRslt"
  Public Sub XRslt_Init()
    With XRslt
      .Code = Plcy_Strg
      .Candidat = {"0", "0"}
      .Cellule = {-1, -1}
      .CelExcl = New List(Of XCel_Excl_Cls)
      .Productivité = False
      .XRoads_Nombre = -1
      .XRoads_Numéro = -1
      .XLinks_Nombre = -1
      .RoadRight = New List(Of XLink_Cls)
    End With
  End Sub

  Public Sub XRslt_Display_New()
    ' La procédure affiche la structure XRslt et VERIFIE si les candidats à enlever font partie de la Solution DL
    Dim Nb As Integer = 0

    Jrn_Add(, {"Affichage des données XRslt"})
    With XRslt
      Jrn_Add(, {"Code Stratégie " & .Code & ", " & Stg_Get(.Code).Texte})
      Jrn_Add(, {"Candidat       " & String.Join(", ", .Candidat)})
      Jrn_Add(, {" La valeur 0 d'un candidat signifie que ce candidat n'est pas pris en compte."}, "Italique")
      Jrn_Add(, {"Cellules       " & String.Join(", ", .Cellule.Select(Function(c) U_Coord(c)))})
      Jrn_Add(, {" La valeur # -1# d'une cellule signifie que la stratégie n'utilise pas cette information."}, "Italique")

      Jrn_Add(, {"XRoads_Nombre  " & .XRoads_Nombre.ToString()})
      Jrn_Add(, {"XRoads_Numéro  " & .XRoads_Numéro.ToString()})
      Jrn_Add(, {"XLinks_Nombre  " & .XLinks_Nombre.ToString()})
      Jrn_Add(, {"Productivité   " & .Productivité.ToString()})
      Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(.Durée_ms)
      Jrn_Add(, {"Durée          " & CStr(.Durée_ms).PadLeft(5) & " ms, soit " & String.Format("{0:00}:{1:00}:{2:00}:{3:000}", Ts.Hours, Ts.Minutes, Ts.Seconds, Ts.Milliseconds)})
    End With

    ' Chemin Correct
    If XRslt.RoadRight.Count <> 0 Then
      Jrn_Add(, {"Affichage de XRslt.RoadRight.Count : " & XRslt.RoadRight.Count & " Lignes."})
      Nb = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1 : Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(2) & " " & XLink_Str_New(Link)})
      Next Link
    End If

    Jrn_Add(, {"Affichage de XRslt.CelExcl : " & XRslt.CelExcl.Count & " Lignes."})
    Nb = 0
    For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
      With XCel
        Nb += 1
        Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & U_Coord(.Cel) & " Candidat : " & .Cdd & " Cellules extrémités : " & U_Coord(.Exc(0)) & " <--> " & U_Coord(.Exc(1))})
      End With
    Next XCel

    'XSolution est calculé dans Game_Load (Game_New_Game)
    Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U)
    Jrn_Add(, {"Dancing Link        : " & CStr(DL.Nb_Solution)})
    Select Case DL.Nb_Solution
      Case -1, 0
      Case 1
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode})
        Jrn_Add(, {"DL_Solution(0) " & XSolution})
        For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
          Nb += 1
          With XCel
            Dim DL_Cellule As Integer = .Cel
            If XSolution(DL_Cellule) = .Cdd Then
              Jrn_Add(, {" La cellule " & U_Coord(DL_Cellule) & " doit prendre la valeur " & .Cdd & " !"}, "Erreur")
            End If
          End With
        Next XCel
      Case Else
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode & " Solutions multiples."})
    End Select

  End Sub

  Public Sub XRslt_Display(Type As String)
    Dim Nb As Integer = 0

    Jrn_Add(, {"Affichage des données XRslt"})
    With XRslt
      Jrn_Add(, {"Code Stratégie " & .Code & ", " & Stg_Get(.Code).Texte})
      Jrn_Add(, {"Candidat       " & String.Join(", ", .Candidat)})
      Jrn_Add(, {" La valeur 0 d'un candidat signifie que ce candidat n'est pas pris en compte."}, "Italique")
      Jrn_Add(, {"Cellules       " & String.Join(", ", .Cellule.Select(Function(c) U_Coord(c)))})
      Jrn_Add(, {" La valeur # -1# d'une cellule signifie que la stratégie n'utilise pas cette information."}, "Italique")

      Jrn_Add(, {"XRoads_Nombre  " & .XRoads_Nombre.ToString()})
      Jrn_Add(, {"XRoads_Numéro  " & .XRoads_Numéro.ToString()})
      Jrn_Add(, {"XLinks_Nombre  " & .XLinks_Nombre.ToString()})
      Jrn_Add(, {"Productivité   " & .Productivité.ToString()})
      Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(.Durée_ms)
      Jrn_Add(, {"Durée: " & CStr(.Durée_ms).PadLeft(5) & " ms, soit " & String.Format("{0:00}:{1:00}:{2:00}:{3:000}", Ts.Hours, Ts.Minutes, Ts.Seconds, Ts.Milliseconds)})
    End With

    ' Chemin Correct
    If XRslt.RoadRight.Count <> 0 Then
      Jrn_Add(, {"Affichage de XRslt.RoadRight.Count : " & XRslt.RoadRight.Count & " Lignes."})
      Nb = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        Select Case Type
          Case "1" : Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(2) & " " & XLink_Str1(Link)})
          Case "2" : Jrn_Add(, {ChrW(Nb + Lettre_Flèche_ChrW) & " " & CStr(Nb).PadLeft(2) & " " & XLink_Str2(Link)})
        End Select
      Next Link
    End If

    Jrn_Add(, {"Affichage de XRslt.CelExcl : " & XRslt.CelExcl.Count & " Lignes."})
    Nb = 0
    For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
      With XCel
        Nb += 1
        Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & U_Coord(.Cel) & " Candidat : " & .Cdd & " Cellules extrémités : " & U_Coord(.Exc(0)) & " <--> " & U_Coord(.Exc(1))})
      End With
    Next XCel

    'XSolution est calculé dans Game_Load (Game_New_Game)
    Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U)
    Jrn_Add(, {"Dancing Link        : " & CStr(DL.Nb_Solution)})
    Select Case DL.Nb_Solution
      Case -1, 0
      Case 1
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode})
        Jrn_Add(, {"DL_Solution(0) " & XSolution})
        For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
          Nb += 1
          With XCel
            Dim DL_Cellule As Integer = .Cel
            If XSolution(DL_Cellule) = .Cdd Then
              Jrn_Add(, {" La cellule " & U_Coord(DL_Cellule) & " doit prendre la valeur " & .Cdd & " !"}, "Erreur")
            End If
          End With
        Next XCel
      Case Else
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode & " Solutions multiples."})
    End Select

  End Sub

  Private Sub Road_Display(Road_Numéro As Integer, Road As List(Of XLink_Cls))
    Jrn_Add(, {"Chemin: " & Road_Numéro.ToString().PadLeft(4) & " (tronçons : " & Road.Count.ToString() & ")"})
    Jrn_Add(, {(String.Join(vbCrLf, Road.Select(Function(f, Idx) (Idx + 1).ToString().PadLeft(2) & " " & XLink_Str1(f))))})
  End Sub
  Private Sub Road_Display2(Road_Numéro As Integer, Road As List(Of XLink_Cls))
    Jrn_Add(, {"Chemin: " & Road_Numéro.ToString().PadLeft(4) & " (tronçons : " & Road.Count.ToString() & ")"})
    Jrn_Add(, {(String.Join(vbCrLf, Road.Select(Function(f, Idx) (Idx + 1).ToString().PadLeft(2) & " " & XLink_Str_New(f))))})
  End Sub

  Public Sub Road_Display_New(Road As List(Of XLink_Cls))
    'Affichage d'un chemin Road de Liens
    Jrn_Add(, {"Le chemin comporte : " & Road.Count & " lien(s)."})
    Jrn_Add(, {(String.Join(vbCrLf, Road.Select(Function(f, Idx) (Idx + 1).ToString().PadLeft(2) & " " & XLink_Str_New(f))))})
  End Sub

  Public Sub XAllRoads_List_Display(Optional Id_Road As Integer = -1)
    ' Affichage de La liste XAllRoads_List
    Jrn_Add(, {Proc_Name_Get()})
    Dim Road_Numéro As Integer = 0
    If XAllRoads_List.Count <> 0 Then
      Jrn_Add(, {"Affichage de XAllRoads_List : " & XAllRoads_List.Count & " Chemins."})
      For Each Road As List(Of XLink_Cls) In XAllRoads_List
        Road_Numéro += 1
        If Id_Road = -1 OrElse Id_Road = Road_Numéro Then
          Road_Display(Road_Numéro, Road)
        End If
      Next Road
    End If
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub

  Public Sub XAllRoads_List_Display_New(Optional Id_Road As Integer = -1)
    ' Affichage de La liste XAllRoads_List
    Jrn_Add(, {Proc_Name_Get()})
    Dim Road_Numéro As Integer = 0
    If XAllRoads_List.Count <> 0 Then
      Jrn_Add(, {"Affichage de XAllRoads_List : " & XAllRoads_List.Count & " Chemins."})
      For Each Road As List(Of XLink_Cls) In XAllRoads_List
        Road_Numéro += 1
        If Id_Road = -1 OrElse Id_Road = Road_Numéro Then
          Road_Display2(Road_Numéro, Road)
        End If
      Next Road
    End If
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub


#End Region

  Public Sub Stratégies_G_End()
    ' Traitement commun aux stratégies G

    Select Case Plcy_Strg
      Case "Gbl", "Gbv", "GCs", "GCx"
        If GRslt.Productivité Then
          Frm_SDK.Mnu0902.Text = "Supprimer " & GRslt.CelExcl_hs.Count & " Candidat(s) : " & Stg_Get(Plcy_Strg).Texte & "."
          Frm_SDK.Mnu0902.Enabled = True
          GRslt_Display() 'Pour contrôle
        Else
          Jrn_Add(, {Stg_Get(Plcy_Strg).Texte & " sans résultat."})
          Frm_SDK.Mnu0902.Enabled = False
        End If
      Case "XCx", "XCy", "Xrp", "XNl", "WgX", "WgY", "WgZ", "WgW"
        If XRslt.Productivité Then
          Frm_SDK.Mnu0902.Text = "Supprimer " & XRslt.CelExcl.Count & " Candidat(s) : " & Stg_Get(Plcy_Strg).Texte & "."
          Frm_SDK.Mnu0902.Enabled = True
          XRslt_Display_New() 'Pour contrôle
        Else
          Jrn_Add(, {Stg_Get(Plcy_Strg).Texte & " sans résultat."})
          Frm_SDK.Mnu0902.Enabled = False
        End If
    End Select
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub
End Module