Option Strict On
Option Explicit On
Imports System.Drawing.Drawing2D
Friend Module M03_Sélection
  '-------------------------------------------------------------------------------
  ' Traitement de la Sélection 
  '-------------------------------------------------------------------------------

  Sub Cell_Val_Insert(V As String, Cellule As Integer, Origine As String)
    ' 01  Les Conditions d'Insertion
    If Cellule < 0 Or Cellule > 80 Then Exit Sub
    If U(Cellule, 2) <> " " Then Exit Sub
    If (V < "1") Or (V > "9") Then Exit Sub
    If Plcy_Gnrl = "Edi" Then Exit Sub
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "Obj" Then Exit Sub

    Game_Undo_Redo = "Normal"
    Dim Av_Jeu As String = Act_Jeu()
    Dim Av_AllCdd As String = Act_Candidats()
    Dim Candidats_Avant As String = U(Cellule, 3)
    Pbl_Cell_Select = Cellule

    ' 02  L'insertion dans les ressources
    Select Case Plcy_Gnrl
      Case "Nrm"
        If Plcy_Solution_Existante = True And V = U_Sol(Cellule) _
        Or Plcy_Solution_Existante = False Then
          U(Cellule, 2) = V : U(Cellule, 3) = Cnddts_Blancs
          U_CddExc(Cellule) = Cnddts_Blancs
          Cdd_Remove_Cell_Coll(U, Cellule)
          Act_Add(Cellule, "Ajouter", V, Candidats_Avant, Origine, Av_Jeu, Av_AllCdd)
        End If
        If Plcy_Solution_Existante = True And V <> U_Sol(Cellule) Then
          Insertion_Exclusion_Nb_Erreurs += 1
          Act_Add(Cellule, "? Ajouter", V, Candidats_Avant, Origine, Av_Jeu, Av_AllCdd)
        End If
        Pbl_Cell_Candidat_CdS = V
      Case "Sas"
        U(Cellule, 2) = V : U(Cellule, 3) = Cnddts_Blancs
    End Select

    ' 03  L'affichage du résultat
    '     Traitements communs
    '     L'insertion ne concerne qu'une cellule à la fois
    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    Dim Gril As New Grille_Cls
    Select Case Plcy_Gnrl
      Case "Nrm"
        Select Case Plcy_Strg
          Case "   "
            Event_OnPaint = "Cell_Val_Insert"
            Using reg As New Region(Sqr_Pth(Cellule)) ' le Sqr_Cel est un rectangle, alors que le Sqr_Pth comporte les arrondis
              Frm_SDK.Invalidate(reg, False)
            End Using
            Application.DoEvents()   'Affiche la grille avec solutions
          Case Else
            If Plcy_AideGraphique Then
              Event_OnPaint = "Global"
              Application.DoEvents()
            End If
            If Not Plcy_AideGraphique Then
              ' Todo à voir plus tard
              'sc.Cellule_Refresh_Cell_Coll()
              Event_OnPaint = "Global"
              Application.DoEvents()
            End If
        End Select
      Case "Sas"
        Event_OnPaint = "Cell_Val_Insert"
        Using reg As New Region(Sqr_Pth(Cellule)) ' le Sqr_Cel est un rectangle, alors que le Sqr_Pth comporte les arrondis
          Frm_SDK.Invalidate(reg, False)
        End Using
        Application.DoEvents()
    End Select

    ' Traitements communs
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()

    Insert_Nb_Cell += 1

    ' 05 Fin de partie
    If Wh_Nb_Cell(U).Remplies = 81 Then
      'L'animation est faite en 2 temps d'abord l'animation 
      Event_OnPaint = "Animation"
      Frm_SDK.Invalidate()
      Application.DoEvents()
      'puis un Gril.Grille_Refresh_g(e.Graphics)
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()   'Affiche la grille avec solutions
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
      Case "Sas"
        U(Cellule, 2) = " "
        U(Cellule, 3) = Cnddts_Blancs
    End Select
    Act_Add(Cellule, "Effacer", VE, Cnddts_Blancs, Origine, Av_Jeu, Av_AllCdd)
    Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_00114", {CStr(Game_Nb_Cellules_Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})

    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub

  Sub Cell_Cdd_Insert(V As String, Cellule As Integer, Origine As String)
    'Le candidat est enlevé des candidats U(Cellule,3)
    '   ET       est ajouté dans les candidats Exclus U_CddExc(Cellule)
    If Plcy_Gnrl = "Edi" Then Exit Sub
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "Obj" Then Exit Sub

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
      Dim sc As New Cellule_Cls With {.Numéro = Cellule}
      Dim Gril As New Grille_Cls
      ' TODO à refaire
      'G3_Grille_Paint_Indirecte()
      'Gril.G3_Grille_Paint_Indirecte()
      'sc.Cellule_Refresh()
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

    ''Lors de chaque insertion, si la mode Suggestion est actif, alors Pzzl_Suggest est lancé
    '' Insertion d'une valeur   Cell_Val_Insert
    '' Insertion d'un  candidat Cell_Cdd_Insert
    'If Plcy_Gnrl = "Nrm" And Plcy_Strg = "   " And Swt_Mode_Suggestion = 1 Then
    '  'Affiche du coup dans la zone Info l'explication de la suggestion
    '  Cell_Slv_Interactif("S", "Mode Suggestion")
    'End If
  End Sub


  Sub Cell_Cdd_Exclude(V As String, Cellule As Integer)
    If Plcy_Gnrl = "Edi" Then Exit Sub
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "Obj" Then Exit Sub
    Try
      If U(Cellule, 3).Contains(V) = False Then Exit Sub
      Game_Undo_Redo = "Normal"
      'Avant toute modification
      Dim Av_Jeu As String = Act_Jeu()
      Dim Av_AllCdd As String = Act_Candidats()

      Dim Candidats As String = U(Cellule, 3)
      If Plcy_Solution_Existante = True And V <> U_Sol(Cellule) _
      Or Plcy_Solution_Existante = False Then
        Dim Candidats_Exclus As String = U_CddExc(Cellule)
        Mid$(Candidats, CInt(V), 1) = " "
        Mid$(Candidats_Exclus, CInt(V), 1) = V
        U(Cellule, 3) = Candidats
        U_CddExc(Cellule) = Candidats_Exclus
        Act_Add(Cellule, "Exclure_Cdd", V, Candidats, Plcy_Strg, Av_Jeu, Av_AllCdd)
      End If
      If Plcy_Solution_Existante = True And V = U_Sol(Cellule) Then
        Insertion_Exclusion_Nb_Erreurs += 1
        Act_Add(Cellule, "? Exclure_Cdd", V, Candidats, Plcy_Strg, Av_Jeu, Av_AllCdd)
      End If
      Pbl_Cell_Select = Cellule
      Dim sc As New Cellule_Cls With {.Numéro = Cellule}
      Dim Gril As New Grille_Cls
      'TODO à refaire aussi
      'Gril.G3_Grille_Paint_Indirecte()
      'sc.Cellule_Refresh()
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

  Public Sub Grid_Cdd_Remove_Cell_Coll(ByRef U_temp(,) As String)
    For i As Integer = 0 To 80
      Cdd_Remove_Cell_Coll(U_temp, i)
    Next i
  End Sub
End Module