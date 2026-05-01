Friend Module M03_Sélection
  '-------------------------------------------------------------------------------
  ' Traitement de la Sélection 
  '-------------------------------------------------------------------------------
  Sub Cell_Val_Insert(Val As String, Cellule As Integer, Origine As String)
    ' 3 Utilisations : Frm_SDK_MouseClick     La cellule et le candidat sont testés
    '                  Mnu_Cel_Val_Insérer    La cellule et le candidat sont testés dans le menu
    '                  Cell_Slv_Result        les tests sont effectués dans le calcul de la résolution
    ' La cellule est obligatoirement une cellule, une cellule vide et le candidat est correct
    ' 011 Cellule et Val   
    ' 012 Les Policy
    If Plcy_Gnrl <> "Nrm" Then Exit Sub
    If Plcy_Strg <> "Sai" AndAlso Not Cell_Cdd_Controle(Val, Cellule, "Include") Then Exit Sub
    ' 013 Undo-Redo
    Game_Undo_Redo = "Normal" ' Peut prendre la valeur "Normal" ou Action
    Dim Av_Jeu As String = Act_Jeu()
    Dim Av_AllCdd As String = Act_Candidats()
    Dim Candidats_Avant As String = U(Cellule, 3)
    Pbl_Cell_Select = Cellule

    ' 02  L'insertion dans les ressources
    U(Cellule, 2) = Val : U(Cellule, 3) = Cnddts_Blancs
    U_CddExc(Cellule) = Cnddts_Blancs
    U_nb(0) += 1                         ' Décompte les cellules saisies
    U_nb(CInt(Val)) += 1                 ' Décompte les cellules par valeur

    ' 021 Traitement des cellules collatérales
    Cdd_Remove_Cell_Coll_Opt(U, Cellule)

    ' 022 Traitement divers
    Act_Add(Cellule, "Ajouter", Val, Candidats_Avant, Origine, Av_Jeu, Av_AllCdd)
    Pbl_Valeur_CdS = Val
    'CalculDernieresVides()
    Build_Bmp_Valeurs()
    Mnu_Mngt_Barre_Outils_Filtres_Enabled()
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00112", {U_nb(10).ToString(), (81 - U_nb(0)).ToString()})
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()

    ' 03  L'affichage du résultat
    Frm_SDK.Invalidate()
    ' 04  Fin de partie
    If Wh_Nb_Cell(U).Remplies = 81 Then
      Dim U_Chk(80, 3) As String
      Array.Copy(U, U_Chk, UNbCopy)
      Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
      If U_Check.Check AndAlso Wh_Nb_Cell(U).Initiales < 81 Then
        Plcy_Strg = "   "
        Jrn_Add(, {"La grille est correcte."}, "Red")
        Frm_SDK.B_Info.Text = "La grille est correcte."
        ' Configuration du Timer
        Frm_SDK.AnimationTimer.Interval = 100 ' ms
        Frm_SDK.AnimationTimer.Start()
      End If
    End If
  End Sub
  Sub Cell_Val_Delete(Cellule As Integer, Origine As String)
    'Avant toute modification
    Dim Av_Jeu As String = Act_Jeu()
    Dim Av_AllCdd As String = Act_Candidats()
    Game_Undo_Redo = "Normal"
    Dim Val As String = U(Cellule, 2)        'Val effacée  

    If Plcy_Gnrl = "Edi" Then Exit Sub
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "Obj" Then Exit Sub
    Pbl_Cell_Select = Cellule
    Select Case Plcy_Gnrl
      Case "Nrm"
        U(Cellule, 2) = " "
        U(Cellule, 3) = Cnddts
        U_CddExc(Cellule) = Cnddts_Blancs
        U_nb(0) -= 1                         ' Décompte les cellules saisies
        U_nb(CInt(Val)) -= 1                 ' Décompte les cellules par valeur

        'Remettre le candidat enlevé Val dans les cellules collatérales
        Dim Grp() As Integer = U_20Cell_Coll(Cellule)
        For g As Integer = 0 To UBound(Grp)
          If U(Grp(g), 2) <> " " Then Continue For
          If U(Grp(g), 3).Contains(Val) = False Then
            Dim Candidats As String = U(Grp(g), 3)
            Mid$(Candidats, CInt(Val), 1) = Val
            U(Grp(g), 3) = Candidats
          End If
        Next g
        Grid_Cdd_Remove_Cell_Coll(U)
    End Select
    Build_Bmp_Valeurs()
    Mnu_Mngt_Barre_Outils_Filtres_Enabled()
    Act_Add(Cellule, "Effacer", Val, U(Cellule, 3), Origine, Av_Jeu, Av_AllCdd)
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Game_Nb_Cellules_Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
    Frm_SDK.Invalidate()
  End Sub

  Sub Cell_Cdd_Insert(V As String, Cellule As Integer, Origine As String)
    'Le candidat est enlevé des candidats U(Cellule,3)
    '   ET       est ajouté dans les candidats Exclus U_CddExc(Cellule)
    Try
      'Avant toute modification
      Game_Undo_Redo = "Normal"
      Dim Av_Jeu As String = Act_Jeu()
      Dim Av_AllCdd As String = Act_Candidats()

      Dim Candidats As String = U(Cellule, 3)
      Dim Candidats_Exclus As String = U_CddExc(Cellule)
      Mid$(Candidats, CInt(V), 1) = V
      Mid$(Candidats_Exclus, CInt(V), 1) = " "
      U(Cellule, 3) = Candidats
      U_CddExc(Cellule) = Candidats_Exclus
      ' Evite le message pour les cellules déjà remplies
      If U(Cellule, 1) <> " " Then Act_Add(Cellule, "Replacer" & Origine, V, Candidats_Exclus, Proc_Name_Get(), Av_Jeu, Av_AllCdd)
      Pbl_Cell_Select = Cellule

      Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Game_Nb_Cellules_Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})

      Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
      Frm_SDK.Invalidate()

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Sub Cell_Cdd_Exclude_GRslt()
    If GRslt.CelExcl_hs.Count > 0 Then
      For Each CelExcl_key As Tuple(Of Integer, String) In GRslt.CelExcl_hs
        Dim XCel_Cel As Integer = CelExcl_key.Item1
        Dim XCel_Cdd As String = CelExcl_key.Item2
        Dim Candidats As String = U(XCel_Cel, 3)
        Dim Candidats_Exclus As String = U_CddExc(XCel_Cel)
        Mid$(Candidats, CInt(XCel_Cdd), 1) = " "
        Mid$(Candidats_Exclus, CInt(XCel_Cdd), 1) = XCel_Cdd
        U(XCel_Cel, 3) = Candidats
        U_CddExc(XCel_Cel) = Candidats_Exclus
        Act_Add(XCel_Cel, "Exclure_Cdd", XCel_Cdd, Candidats, Plcy_Strg, Act_Jeu(), Act_Candidats())
      Next
    End If
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Game_Nb_Cellules_Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
    Frm_SDK.Invalidate()
  End Sub
  Sub Cell_Cdd_Exclude(V As String, Cellule As Integer)
    If Plcy_Gnrl = "Edi" Then Exit Sub
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "Obj" Then Exit Sub
    Try
      ' Jrn_Add_Yellow(Proc_Name_Get())
      'If Candidat = XSolution(Cellule) Then Exit Sub
      If Not Cell_Cdd_Controle(V, Cellule, "Exclude") Then Exit Sub
      If U(Cellule, 3).Contains(V) = False Then Exit Sub
      Game_Undo_Redo = "Normal"
      'Avant toute modification
      Dim Av_Jeu As String = Act_Jeu()
      Dim Av_AllCdd As String = Act_Candidats()

      Dim Candidats As String = U(Cellule, 3)
      Dim Candidats_Exclus As String = U_CddExc(Cellule)
      Mid$(Candidats, CInt(V), 1) = " "
      Mid$(Candidats_Exclus, CInt(V), 1) = V
      U(Cellule, 3) = Candidats
      U_CddExc(Cellule) = Candidats_Exclus
      Act_Add(Cellule, "Exclure_Cdd", V, Candidats, Plcy_Strg, Av_Jeu, Av_AllCdd)
      Pbl_Cell_Select = Cellule
      Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Game_Nb_Cellules_Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})

      Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
      Frm_SDK.Invalidate()

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Function Cdd_Remove_Cell_Coll_Opt(ByRef U_temp(,) As String, Cellule As Integer) As Integer
    ' Enlever la valeur placée dans la Cellule des 20 Cellules Collatérales
    ' Retourne le nombre de cellules dans lesquelles un candidat a été enlevé 
    '         -1 en cas d'erreur
    ' Le tableau U_temp des cellules est passé en ByRef, car il sort modifié de la fonction 
    Dim nb As Integer = 0
    Dim Val As String = U_temp(Cellule, 2)
    Dim Grp() As Integer = U_20Cell_Coll(Cellule)
    For Each Cell_Coll As Integer In Grp
      If U_temp(Cell_Coll, 2) <> " " Then Continue For  ' La cellule collatérale a déjà une valeur, on continue
      Dim Candidats As String = U_temp(Cell_Coll, 3)
      'If Candidats.Substring(CInt(Val) - 1, 1) = Val Then
      '  Mid$(Candidats, CInt(Val), 1) = " "          ' Remplacer la valeur par un espace
      '  U_temp(Cell_Coll, 3) = Candidats
      '  nb += 1
      'End If
      If Candidats(CInt(Val) - 1) = Val Then
        Dim sb As New System.Text.StringBuilder(Candidats)
        sb(CInt(Val) - 1) = " "c
        U_temp(Cell_Coll, 3) = sb.ToString()
      End If
    Next Cell_Coll
    Return nb
  End Function

  Public Function Cdd_Remove_Cell_Coll(ByRef U_temp(,) As String, Cellule As Integer) As Integer
    ' Enlever la valeur placée dans la Cellule des 20 Cellules Collatérales
    ' Retourne le nombre de cellules dans lesquelles un candidat a été enlevé 
    '         -1 en cas d'erreur
    ' Le tableau U_temp des cellules est passé en ByRef, car il sort modifié de la fonction 

    Dim nb As Integer = 0
    If Cellule < 0 Or Cellule > 80 Then Return -1

    Dim Valeur As String = U_temp(Cellule, 2)
    If Valeur = " " Then Return 0

    Dim Grp() As Integer = U_20Cell_Coll(Cellule)
    For Each Cell_Coll As Integer In Grp
      If U_temp(Cell_Coll, 2) <> " " Then Continue For  ' La cellule collatérale a déjà une valeur, on continue
      Dim Candidats As String = U_temp(Cell_Coll, 3)
      If Candidats.Substring(CInt(Valeur) - 1, 1) = Valeur Then
        Mid$(Candidats, CInt(Valeur), 1) = " "          ' Remplacer la valeur par un espace
        U_temp(Cell_Coll, 3) = Candidats
        nb += 1
      End If
    Next Cell_Coll

    Return nb
  End Function
  Public Sub Grid_Cdd_Remove_Cell_Coll(ByRef U_temp(,) As String)
    For i As Integer = 0 To 80
      Cdd_Remove_Cell_Coll(U_temp, i)
    Next i
  End Sub

  Public Function Cell_Cdd_Controle(Candidat As String, Cellule As Integer, Type_IE As String) As Boolean
    'Jrn_Add_Yellow(Proc_Name_Get())
    'Le contrôle n'est pas effectué s'il n'y a pas de solution DL.
    If XSolution(Cellule) = "0" Then Return True
    Select Case Type_IE
      Case "Include"
        If XSolution(Cellule) <> Candidat Then
          Cpt_Pénalités += 1
          Jrn_Add(, {Cpt_Pénalités.ToString().PadLeft(3) & " Erreur en " & U_Coord(Cellule) & "! Le candidat " & XSolution(Cellule) & " est attendu à la place de " & Candidat & "."}, "Red")
          Return False
        End If
      Case "Exclude"
        If XSolution(Cellule) = Candidat Then
          Cpt_Pénalités += 1
          Jrn_Add(, {Cpt_Pénalités.ToString().PadLeft(3) & " Erreur en " & U_Coord(Cellule) & "! Le candidat " & Candidat & " est la solution. "}, "Red")
          Return False
        End If
    End Select
    Return True
  End Function

  Public Sub Pénalités(Origine As String)
    Cpt_Pénalités += 1
    Jrn_Add(, {Cpt_Pénalités.ToString().PadLeft(3) & " " & Origine}, "Red")
  End Sub

  Public Function UniqueVideLigne(U(,) As String, row As Integer) As Integer
    Dim idxVide As Integer = -1
    Dim count As Integer = 0

    For c As Integer = 0 To 8
      Dim i As Integer = row * 9 + c
      If U(i, 2) = " " Then
        idxVide = i
        count += 1
        If count > 1 Then Return -1
      End If
    Next

    Return If(count = 1, idxVide, -1)
  End Function

  Public Function UniqueVideCol(U(,) As String, col As Integer) As Integer
    Dim idxVide As Integer = -1
    Dim count As Integer = 0

    For r As Integer = 0 To 8
      Dim i As Integer = r * 9 + col
      If U(i, 2) = " " Then
        idxVide = i
        count += 1
        If count > 1 Then Return -1
      End If
    Next

    Return If(count = 1, idxVide, -1)
  End Function

  Public Function UniqueVideRegion(U(,) As String, reg As Integer) As Integer
    Dim r0 As Integer = (reg \ 3) * 3
    Dim c0 As Integer = (reg Mod 3) * 3

    Dim idxVide As Integer = -1
    Dim count As Integer = 0

    For k As Integer = 0 To 8
      Dim r As Integer = r0 + (k \ 3)
      Dim c As Integer = c0 + (k Mod 3)
      Dim i As Integer = r * 9 + c

      If U(i, 2) = " " Then
        idxVide = i
        count += 1
        If count > 1 Then Return -1
      End If
    Next

    Return If(count = 1, idxVide, -1)
  End Function

  Public Sub CalculDernieresVides()

    Array.Clear(U_dv, 0, U_dv.Length)

    ' Lignes
    For r As Integer = 0 To 8
      Dim i As Integer = UniqueVideLigne(U, r)
      If i >= 0 Then U_dv(i) = True
    Next

    ' Colonnes
    For c As Integer = 0 To 8
      Dim i As Integer = UniqueVideCol(U, c)
      If i >= 0 Then U_dv(i) = True
    Next

    ' Régions
    For reg As Integer = 0 To 8
      Dim i As Integer = UniqueVideRegion(U, reg)
      If i >= 0 Then U_dv(i) = True
    Next

  End Sub

End Module