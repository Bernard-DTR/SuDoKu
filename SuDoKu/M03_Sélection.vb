Friend Module M03_Sélection
  '-------------------------------------------------------------------------------
  ' Traitement de la Sélection 
  '-------------------------------------------------------------------------------
  Sub Cell_Val_Insert(V As String, Cellule As Integer, Origine As String)
    ' Jrn_Add_Yellow(Proc_Name_Get())

    ' 01  Les Conditions d'Insertion
    If Cellule < 0 Or Cellule > 80 Then Exit Sub
    If U(Cellule, 2) <> " " Then Exit Sub
    If V < "1" Or V > "9" Then Exit Sub
    If Plcy_Gnrl <> "Nrm" Then Exit Sub

    If Plcy_Strg <> "Sai" AndAlso Not Cell_Cdd_Controle(V, Cellule, "Include") Then Exit Sub

    Game_Undo_Redo = "Normal"
    Dim Av_Jeu As String = Act_Jeu()
    Dim Av_AllCdd As String = Act_Candidats()
    Dim Candidats_Avant As String = U(Cellule, 3)

    Pbl_Cell_Select = Cellule

    ' 02  L'insertion dans les ressources
    U(Cellule, 2) = V : U(Cellule, 3) = Cnddts_Blancs
    U_CddExc(Cellule) = Cnddts_Blancs
    Cell_Coll_Modifiées_List.Clear()
    Cell_Coll_Modifiées_List = Cdd_Remove_Cell_Coll_List(U, Cellule)
    Act_Add(Cellule, "Ajouter", V, Candidats_Avant, Origine, Av_Jeu, Av_AllCdd)
    Pbl_Valeur_CdS = V

    Mnu_Mngt_Barre_Outils_Filtres()
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()

    ' 03  L'affichage du résultat
    Select Case Stg_Get(Plcy_Strg).Family
      Case 0, 2
        ' Aucune stratégie, stratégie des filtres 
        Event_OnPaint_Origine = Proc_Name_Get() & " " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
        Event_OnPaint = "Cellule"
        Using reg As New Region(Sqr_Pth(Pbl_Cell_Select))
          Frm_SDK.Invalidate(reg, False)
        End Using
        Application.DoEvents()

      Case 1
        ' Stratégie Cdd, les candidats sont affichés
        Event_OnPaint_Origine = Proc_Name_Get() & " " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
        Event_OnPaint = "Cell_Coll"
        If Cell_Coll_Modifiées_List.Count > 0 Then
          Using reg As New Region(Sqr_Pth(Pbl_Cell_Select))
            For Each cell As Integer In Cell_Coll_Modifiées_List
              reg.Union(Sqr_Pth(cell))
            Next
            Frm_SDK.Invalidate(reg, False)
            Application.DoEvents()
          End Using
        Else
          ' Si la liste est vide, on peut invalider uniquement la cellule principale
          Using reg As New Region(Sqr_Pth(Pbl_Cell_Select))
            Frm_SDK.Invalidate(reg, False)
            Application.DoEvents()
          End Using
        End If

      Case 3, 4
        ' Toutes les stratégies, et de nombreuses situations spéciales
        Event_OnPaint = "Global"
        Frm_SDK.Invalidate()
        Application.DoEvents()
      Case Else
    End Select

    ' 04 Fin de partie
    If Wh_Nb_Cell(U).Remplies = 81 Then
      ' 041 L'animation est faite en 2 temps d'abord l'animation 
      Event_OnPaint_Origine = Proc_Name_Get() & "_A " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
      Event_OnPaint = "Animation"
      Frm_SDK.Invalidate()
      Application.DoEvents()
      ' 042 puis un affichage de la grille complète
      Event_OnPaint_Origine = Proc_Name_Get() & "_B " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub
  Sub Cell_Val_Delete(Cellule As Integer, Origine As String)
    'Avant toute modification
    Dim Av_Jeu As String = Act_Jeu()
    Dim Av_AllCdd As String = Act_Candidats()
    Game_Undo_Redo = "Normal"
    Dim VE As String = U(Cellule, 2)        'Valeur effacée  

    If Plcy_Gnrl = "Edi" Then Exit Sub
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "Obj" Then Exit Sub
    Pbl_Cell_Select = Cellule
    Select Case Plcy_Gnrl
      Case "Nrm"
        U(Cellule, 2) = " "
        U(Cellule, 3) = Cnddts
        U_CddExc(Cellule) = Cnddts_Blancs
        'Remettre le candidat enlevé VE dans les cellules collatérales
        Dim Grp() As Integer = U_20Cell_Coll(Cellule)
        For g As Integer = 0 To UBound(Grp)
          If U(Grp(g), 2) <> " " Then Continue For
          If U(Grp(g), 3).Contains(VE) = False Then
            Dim Candidats As String = U(Grp(g), 3)
            Mid$(Candidats, CInt(VE), 1) = VE
            U(Grp(g), 3) = Candidats
          End If
        Next g
        Grid_Cdd_Remove_Cell_Coll(U)
    End Select
    'Act_Add(Cellule, "Effacer", VE, Cnddts_Blancs, Origine, Av_Jeu, Av_AllCdd)
    Act_Add(Cellule, "Effacer", VE, U(Cellule, 3), Origine, Av_Jeu, Av_AllCdd)
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Game_Nb_Cellules_Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})

    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
    Event_OnPaint_Origine = Proc_Name_Get() & " " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()

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
      Event_OnPaint_Origine = Proc_Name_Get() & " " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()

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
    Event_OnPaint_Origine = Proc_Name_Get() & " " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
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
      Event_OnPaint_Origine = Proc_Name_Get() & " " & Plcy_Gnrl & " Plcy_Strg: '" & Plcy_Strg & "'"
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub

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
      If U_temp(Cell_Coll, 2) <> " " Then Continue For  ' Si la cellule collatérale a déjà une valeur, continuer
      Dim Candidats As String = U_temp(Cell_Coll, 3)
      If Candidats.Substring(CInt(Valeur) - 1, 1) = Valeur Then
        Mid$(Candidats, CInt(Valeur), 1) = " "        ' Remplacer la valeur par un espace
        U_temp(Cell_Coll, 3) = Candidats
        nb += 1
      End If
    Next Cell_Coll

    Return nb
  End Function
  Public Function Cdd_Remove_Cell_Coll_List(ByRef U_temp(,) As String, cellule As Integer) As List(Of Integer)
    ' TODO à terme cette fonction devrait remplacer Cdd_Remove_Cell_Coll
    ' Le tableau U_temp des cellules est passé en ByRef, car il sort modifié de la fonction 
    ' Enlever la valeur placée dans la Cellule des 20 Cellules Collatérales
    ' Retourne la liste des cellules dans lesquelles un candidat a été enlevé
    Dim list_Coll As New List(Of Integer)

    ' Valeur placée dans la cellule
    Dim valeur As String = U_temp(cellule, 2)
    If valeur = " " Then Return list_Coll   ' Rien à enlever

    Dim v As Integer = CInt(valeur)
    Dim Grp() As Integer = U_20Cell_Coll(cellule)

    For Each cell_coll As Integer In Grp
      ' Si la cellule collatérale a déjà une valeur, on ignore
      If U_temp(cell_coll, 2) <> " " Then Continue For
      Dim candidats As String = U_temp(cell_coll, 3)

      ' Vérification de sécurité : longueur correcte
      If candidats.Length < 9 Then Continue For

      ' Si le candidat est présent
      If candidats(v - 1) = valeur Then
        Dim sb As New System.Text.StringBuilder(candidats)
        sb(v - 1) = " "c
        U_temp(cell_coll, 3) = sb.ToString()
        list_Coll.Add(cell_coll)
      End If
    Next
    Return list_Coll
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
          Jrn_Add(, {"⛔" & "   Erreur en " & U_Coord(Cellule) & "! Le candidat : " & XSolution(Cellule) & " est attendu à la place de " & Candidat & "."}, "Emoji")
          Return False
        End If
      Case "Exclude"
        If XSolution(Cellule) = Candidat Then
          Jrn_Add(, {"⛔" & "   Erreur en " & U_Coord(Cellule) & "! Le candidat " & Candidat & " est la solution. "}, "Emoji")
          Return False
        End If
    End Select
    Return True
  End Function
End Module