Imports SuDoKu.DancingLink
Imports SuDoKu.DancingLink.A_Copyright
'08/12/2023  
' Prd Module de Production de Puzzle
'     Crt Création d'une Grille + Slv Solution d'une grille = Prd d'un Puzzle
'     Pzzl_Prd_Interactive | Pzzl_Prd_Batch = Pzzl_Crt + Pzzl_Slv
Friend Module M50_Production
  'Pzzl_Prd dépend uniquement des paramètres entrés et sortis de Prd.Prd_ qqch
  'Production d'une grille de Sudoku
  'La production d'une grille de Sudoku est faite soit en arrière-plan
  '                                               soit interactivement

  Public Prv_U_temp(80, 3) As String             ' Usage exclusif de Jrn_Add_Pzzl_Slv
  Public Prv_Col3 As String                      ' Usage exclusif de Jrn_Add_Pzzl_Slv    
  Public Prv_Col1, Prv_Col10 As Integer
  Public Col1, Col3, Col8, Col10 As String
  Public Act_Col3 As String
  Public Act_Col1, Act_Col10 As Integer

  Public Sub Pzzl_Crt(ByVal Production_Type As String, ByRef Prd As Prd_Struct)
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ' Pzzl_Crt  création correcte d'une grille de Sudoku 
    '                    crée une grille COMPLETE correctement
    ' Traitement réussi pour 81 cellules remplies, avec et sans symétrie
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '   Pzzl_Crt peut créer une grille complète de 81 cellules 
    '   Rappel Jrn_Add(,{"Message...}) MsgId est optionnel et a pour valeur omise SDK_00000
    '   La collection des régions ne comporte plus que des régions éligibles
    '   Prd.Code_Retour = -2  lors du lancement de Pzzl_Crt 
    '                   = -1  si U_Check.Check = False (U_Checking(U_temp))
    '                         s'il existe des cellules sans candidats
    '                         si le nombre d'anamalies excède la limite de 10
    '                         si le nombre de cellules demandées n'est pas atteint
    '                   = 0   Dans tous les autres cas
    '   Si Chat: Sont listés les messages 
    '   A l'issue de Pzzl_Crt, si une grille est créée, il faut vérifier qu'elle soit solvable
    '
    '////////////////////////////////////////////////////////////////////////////////////////////////////////
    '
    '   Pzzl_Crt est utilisé dans Pzzl_Crt_Interactif
    '                             Pzzl_Prd_Interactif
    '            dans ces 2 cas la structure Prd est initialisée des valeurs de Préférences / Création
    '
    ' 13/02/2024 Les anomalies sont numérotées
    '            A chaque stratégie, le nombre total de candidats restant est indiqué.
    ' Principe   : 
    '              La collection des Régions est effectuée dès qu'elle est vide
    '            1 On constitue une collection des régions éligibles
    '               (elle doit avoir au moins une cellule vide)
    '            2 On choisit une région au hasard
    '            3 On constitue une collection des cellules éligibles
    '               (une cellule éligible est une cellule vide)
    '            4 On choisit une cellule au hasard
    '            5 On constitue une collection des candidats
    '            6 On choisit un candidat au hasard
    '   il y a une alternative au point 4 lorsqu'une demande de symétrie est paramétrée

    Dim U_temp(80, 3) As String             ' Définition de U_temp pour pouvoir utiliser Prd isolément 
    Dim Régions_Clct As New Collection      ' Collection des 9 Régions
    Dim Région_Num As Integer               ' Numéro de la Région choisie au hasard
    Dim Cellules_Clct As New Collection     ' Collection des Cellules Vides de la Région choisie au hasard
    Dim Candidats_Clct As New Collection
    Dim Nb_Cellules_Ajoutées As Integer
    Dim Cellule As Integer
    Dim Prv_Cellule As Integer
    Dim Candidats As String
    Dim Candidat As String = ""
    Dim Strategy_Rslt(99, 0) As String
    Dim Strategy_Nb As Integer
    'Une limite de 20 anomalies semble raisonnable, y compris pour une grille de 81 cellules demandées
    Dim Anomalies_Nb_Lim As Integer = 20
    Dim Anomalies_Nb As Integer = 0

    Dim Alter As Integer = 1
    Prd.Prd_Phase = "Crt"
    If Prd.Prd_Chat Then Jrn_Add(, {Proc_Name_Get()})

    ' Il semble que le fait de choisir une cellule au hasard par région (au hasard) disperse mieux
    '  les valeurs initiales plutôt que de choisir une cellule au hasard dans la grille.

Pzzl_Crt_Init:
    ' Etape  1: Initialisations diverses, la grille est vierge, Cnddts = "123456789" 
    For i As Integer = 0 To 80
      U_temp(i, 1) = " "       'Valeurs Initiales 
      U_temp(i, 2) = " "       'Valeur
      U_temp(i, 3) = Cnddts    'Candidats 123456789
    Next i

Pzzl_Choix_Cellule:
    ' Le choix des cellules s'arrête si
    '    Le nombre d'anomalies dépasse la limite fixée
    '    le nombre de cellules choisies atteint le nombre souhaité
    '    le nombre de cellules vides est égal au nombre total de candidats, soit un candidat par cellule
    Alter *= -1

    If Wh_Grid_Nb_Cellules_Vides(U_temp) = Wh_Grid_Nb_Candidats(U_temp) Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10251", {}, "Italique")
    End If

    If Anomalies_Nb >= Anomalies_Nb_Lim _
    Or Nb_Cellules_Ajoutées = CInt(Prd.Prd_Create_Nb_Cel_Demandées) _
    Or Wh_Grid_Nb_Cellules_Vides(U_temp) = Wh_Grid_Nb_Candidats(U_temp) Then GoTo Pzzl_Crt_End

    ' Choix de la Cellule et de la Valeur 
    ' La cellule est choisie au hasard dans une région choisie au hasard

    If (Not Prd.Prd_Cnt_Sym) Or (Prd.Prd_Cnt_Sym And Alter = -1) Then
      ' Lorsque la collection des régions est vide, elle est créée ou à nouveau remplie
      If Régions_Clct.Count = 0 Then
        'La collection Régions_Clct doit comporter uniquement les régions ayant au moins une cellule vide
        For r1 As Integer = 0 To 8
          'Avec And les 3 expressions sont évaluées
          'Avec AndAlso l'expression est évaluée si la première est True
          If Prd.Prd_Cnt_Excl AndAlso Prd.Prd_Cnt_Type = "R" _
                        AndAlso r1 = Prd.Prd_Cnt_Valeur Then Continue For
          Dim Grp_r1() As Integer = U_9CelReg(r1)
          For r2 As Integer = 0 To 8
            If U_temp(Grp_r1(r2), 2) = " " Then
              Clct_Add(Régions_Clct, r1)
              Exit For
            End If
          Next r2
        Next r1
      End If

      'La Collection des régions est remplie, si elle est vide le jeu s'arrête
      If Régions_Clct.Count = 0 Then
        If Prd.Prd_Chat Then Jrn_Add("PRD_10222", {CStr(Anomalies_Nb), CStr(Anomalies_Nb_Lim)}, "Italique")
        Anomalies_Nb += 1
        GoTo Pzzl_Crt_End
      End If

      Région_Num = Clct_Random(Régions_Clct) ' Une région au hasard

      Cellules_Clct.Clear()           ' Remplissage de la collection
      Dim Grp() As Integer = U_9CelReg(Région_Num)
      ' On ne prend que les cellules vides
      For r As Integer = 0 To 8
        If Prd.Prd_Cnt_Excl AndAlso Prd.Prd_Cnt_Type = "L" AndAlso
                   U_Row(Grp(r)) = Prd.Prd_Cnt_Valeur Then Continue For
        If Prd.Prd_Cnt_Excl AndAlso Prd.Prd_Cnt_Type = "C" AndAlso
                   U_Col(Grp(r)) = Prd.Prd_Cnt_Valeur Then Continue For
        If U_temp(Grp(r), 2) = " " Then Clct_Add(Cellules_Clct, Grp(r))
      Next r
      If Cellules_Clct.Count = 0 Then
        'PRD_10220=La région %0 est complètement remplie
        If Prd.Prd_Chat Then Jrn_Add("PRD_10220", {CStr(Anomalies_Nb), CStr(Anomalies_Nb_Lim), CStr(Région_Num)}, "Italique")
        Anomalies_Nb += 1
        GoTo Pzzl_Choix_Cellule
      End If
      ' Une cellule est choisie parmi les cellules vides
      Cellule = Clct_Random(Cellules_Clct)
      Prv_Cellule = Cellule
    End If

    'La symétrie choisie en début de traitement est appliquée
    'NOTE: La recherche de la cellule symétrique n'est pas concernée par la contrainte d'exclusion
    '      car les contraintes Symétrie et Exclusion sont exclusives.
    If (Prd.Prd_Cnt_Sym And Alter = 1) Then
      Select Case Prd.Prd_Cnt_Valeur
        Case 0 : Cellule = Wh_Cellule_ColRow(U_Row(Prv_Cellule), U_Col(Prv_Cellule))
        Case 1 : Cellule = Wh_Cellule_ColRow(8 - U_Row(Prv_Cellule), 8 - U_Col(Prv_Cellule))
        Case 2 : Cellule = Wh_Cellule_ColRow(8 - U_Col(Prv_Cellule), 8 - U_Row(Prv_Cellule))
        Case 3 : Cellule = Wh_Cellule_ColRow(U_Col(Prv_Cellule), 8 - U_Row(Prv_Cellule))
        Case 4 : Cellule = Wh_Cellule_ColRow(8 - U_Col(Prv_Cellule), U_Row(Prv_Cellule))
      End Select
      ' En mode asymétrique, les cellules remplies étant exclues, il ne peut pas avoir de problèmes de symétrie
      ' La cellule symétrique n'est pas vide
      If U_temp(Cellule, 2) <> " " Then
        'PRD_10230=%0 symétrique de %1 est déjà remplie.
        ' Ce qui est vraisemblable avec les cellules placées sur les axes de symétrie !
        If Prd.Prd_Chat Then Jrn_Add("PRD_10230", {CStr(Anomalies_Nb), CStr(Anomalies_Nb_Lim), U_Coord(Cellule), U_Coord(Prv_Cellule)}, "Italique")
        Anomalies_Nb += 1
        GoTo Pzzl_Choix_Cellule
      End If
    End If

Pzzl_Choix_Candidat:
    ' PRD_10241=|                             | La cellule choisie %0 ne contient plus de candidats.
    ' Anomalie à envisager
    If U_temp(Cellule, 3) = Cnddts_Blancs Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10241", {CStr(Anomalies_Nb), CStr(Anomalies_Nb_Lim), U_Coord(Cellule)}, "Italique")
      Anomalies_Nb += 1
      GoTo Pzzl_Crt_End
    End If
    Candidats_Clct.Clear()
    Candidats = U_temp(Cellule, 3)
    For i As Integer = 1 To 9  ' La collection est remplie avec les candidats restants
      Dim Cdd_1 As String = Mid$(Candidats, i, 1)
      If Cdd_1 <> " " Then
        If Prd.Prd_Cnt_Excl And Prd.Prd_Cnt_Type = "V" And
                            Cdd_1 = CStr(Prd.Prd_Cnt_Valeur) Then Continue For
        Candidats_Clct.Add(Cdd_1)
      End If
    Next i
    If Candidats_Clct.Count = 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10242", {CStr(Anomalies_Nb), CStr(Anomalies_Nb_Lim), U_Coord(Cellule), CStr(Prd.Prd_Cnt_Valeur)}, "Italique")
      Anomalies_Nb += 1

      GoTo Pzzl_Choix_Cellule
    End If
    Candidat = CStr(Clct_Random(Candidats_Clct))

    ' La valeur trouvée est ajoutée dans la cellule trouvée 
    ' La saisie d'une valeur dans cellule comporte ces 4 opérations
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Nb_Cellules_Ajoutées += 1
    U_temp(Cellule, 1) = Candidat
    U_temp(Cellule, 2) = Candidat
    U_temp(Cellule, 3) = Cnddts_Blancs
    ' Cette valeur est enlevée des 20 cellules collatérales
    Dim Cell_Coll_Nb As Integer = Cdd_Remove_Cell_Coll(U_temp, Cellule)

    ' La cellule choisie et le candidat sélectionné sont listés
    '   avec le n° de la cellule, le nombre de candidats enlevés des cell_coll,
    '   et le nombre total de candidats restants de la grille 
    If Prd.Prd_Chat Then Jrn_Add("PRD_10005", {CStr(Nb_Cellules_Ajoutées).PadLeft(3), U_Coord(Cellule), Candidats, Candidat,
                          CStr(Cell_Coll_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

#Region "Pzzl_Crt_Analyse"
Pzzl_Crt_Analyse:
    ' 00 Contrôle de la grille
    Dim U_Chk(80, 3) As String
    Array.Copy(U_temp, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    If Not U_Check.Check Then
      If Prd.Prd_Chat Then U_Checking_Display(U_Check, Prd.Prd_Chat)
      GoTo Pzzl_Crt_End
    End If

    ' Les analyses sont faites en fonction de PRD.Prd_Plcy_Strg_Bll(x), sauf pour les CdU et CdO 
    ' en cas de résultat positif, un message informe de la stratégie et du nombre de candidats concernés
    ' et le traitement est refait.

    ' 02 Analyse des Candidats Uniques _ C'est une stratégie obligatoire quelque soit la Ext_Plcy_Strg_Profondeur
    '    Quand on détecte un candidat unique dans une cellule, 
    '     comme ce candidat sera LA valeur quand la cellule sera choisie,
    '     alors il est NECESSAIRE d'enlever cette valeur dans les 20 cellules collatérales
    Strategy_Nb = 0
    Strategy_Rslt = Strategy_CdU(U_temp)

    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      Dim Cellule_CdU As Integer = CInt(Strategy_Rslt(10, i))
      Dim Candidat_CdU As Integer = CInt(Strategy_Rslt(5, i))
      Dim Grp_CdU() As Integer = U_20Cell_Coll(Cellule_CdU)
      For g As Integer = 0 To UBound(Grp_CdU)
        Dim Cell_Coll As Integer = Grp_CdU(g)
        If U_temp(Cell_Coll, 3).Contains(CStr(Candidat_CdU)) Then
          'Le candidat est enlevé des 20 cellules collatérales
          Mid$(U_temp(Cell_Coll, 3), CInt(Candidat_CdU), 1) = " "
          Strategy_Nb += 1
          Prd.Crt_Strg_Nb(0) += 1
        End If
      Next g
    Next i
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"CdU", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      GoTo Pzzl_Crt_Analyse
    End If

    ' Une stratégie est invoquée si la précédente n'est pas efficiente 
    ' 03 Analyse des Candidats Obligatoires _ C'est une stratégie obligatoire quelque soit la Ext_Plcy_Strg_Profondeur
    '    A ce stade, si une cellule comporte une valeur obligatoire, elle doit être la seule conservée.
    '      Ce traitement permet par la suite de ne pas choisir une autre valeur
    '      Elle re-deviendra alors CdU et sera traitée au-dessus quand elle sera choisie dans un prochain passage
    Strategy_Nb = 0
    Strategy_Rslt = Strategy_CdO(U_temp)
    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      Dim Cellule_CdO As Integer = CInt(Strategy_Rslt(10, i))
      Dim Candidats_CdO As String = Cnddts_Blancs
      Mid$(Candidats_CdO, CInt(Strategy_Rslt(5, i)), 1) = Strategy_Rslt(5, i)
      'Cette fois-ci, le CdO reste le seul candidat.
      U_temp(Cellule_CdO, 3) = Candidats_CdO
      Strategy_Nb += 1
      Prd.Crt_Strg_Nb(1) += 1
    Next i
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {Strategy_Rslt(1, 0), CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 2, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"Cbl", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(2) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 3, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"Tpl", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(3) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 4, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"Xwg", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(4) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 5, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"XYw", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(5) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 6, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"Swf", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(6) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 7, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"Jly", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(7) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 8, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"XYZ", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(8) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

    Strategy_Nb = Strategy_Upd_BTXYSJZKQ("Crt", Production_Type, 9, U_temp, Prd)
    If Strategy_Nb > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10100", {"Sky", CStr(Strategy_Nb).PadLeft(3), CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(3), CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)})
      'Prd.Crt_Strg_Nb(9) += Strategy_Nb
      GoTo Pzzl_Crt_Analyse
    End If

#End Region
    GoTo Pzzl_Choix_Cellule 'Jusqu'au nombre de cellules souhaitées

Pzzl_Crt_End:
    'U_Checking retourne False
    If Not U_Check.Check Then
      Prd.Prd_Code_Retour = -1
      GoTo Pzzl_Crt_Exit
    End If
    'Il existe des Cellules Vides sans candidats
    Dim Cell_VSC As Integer = Wh_Nb_Cell(U_temp).Vides_sans_Candidats
    If Cell_VSC > 0 Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10213", {CStr(Cell_VSC)}, "Erreur")
      Prd.Prd_Code_Retour = -1
      GoTo Pzzl_Crt_Exit
    End If
    'Le nombre d'anomalies dépasse la limite
    If Anomalies_Nb >= Anomalies_Nb_Lim Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10212", {CStr(Anomalies_Nb), CStr(Anomalies_Nb_Lim)}, "Erreur")
      Prd.Prd_Code_Retour = -1
      GoTo Pzzl_Crt_Exit
    End If
    'Le nombre de cellules ajoutées n'atteint pas le nombre de cellules souhaitées
    If Nb_Cellules_Ajoutées < CInt(Prd.Prd_Create_Nb_Cel_Demandées) _
        And Wh_Grid_Nb_Cellules_Vides(U_temp) <> Wh_Grid_Nb_Candidats(U_temp) Then
      If Prd.Prd_Chat Then Jrn_Add("PRD_10214", {CStr(Nb_Cellules_Ajoutées), CStr(Prd.Prd_Create_Nb_Cel_Demandées)}, "Erreur")
      Prd.Prd_Code_Retour = -1
      GoTo Pzzl_Crt_Exit
    End If

    'Pour le reste Le Code_retour est 0
    Prd.Prd_Code_Retour = 0

Pzzl_Crt_Exit:
    'Documentation de Prd
    With Prd
      For i As Integer = 0 To 80
        .Prd_Ini(i) = U_temp(i, 1)
        .Prd_Val(i) = U_temp(i, 2)
        .Prd_Candidats(i) = U_temp(i, 3)
      Next i
    End With
  End Sub

  Public Sub Pzzl_Slv(ByVal Production_Type As String, Cellules_Type As String, ByRef Prd As Prd_Struct, ByRef Strategy_Rslt(,) As String)
    'ByVal Cellules_Type As String             Soit *All, *One
    '      Cellules_Type tente de résoudre TOUTES les cellules ou Une seule            
    'ByRef Prd As Prd_Struct paramètres cochées dans Préférences / Stratégies (y compris les stratégies)
    'ByRef Strategy_Rslt(,) As String
    '      Permet de récupérer le résultat de la dernière stratégie si Cellules_Type = "*One"
    '      TOUTES les stratégies utilisent Strategy_Rslt(,)
    'Pzzl_Slv est une procédure et non une fonction
    '      Prd et Strategy_Rslt(,) sont passées en ByRef
    '
    'La fonction tente de résoudre une grille ou une cellule AVEC les stratégies développées dans SDK
    '            et cochées dans Préférences / Stratégies
    '
    ' Pzzl_Crt propose à Pzzl_Slv une grille dont certains candidats ont été enlevés suite aux stratégies

    ' Initialisation des variables
    Dim U_temp(80, 3) As String
    Dim Strategy_CdU_Nb As Integer = 0
    Dim Strategy_CdO_Nb As Integer = 0

    Prd.Prd_Phase = "Slv"

    ' Configuration de U_temp en fonction du type de production
    Select Case Production_Type
      Case "P"
        For i As Integer = 0 To 80
          U_temp(i, 2) = Prd.Prd_Val(i)
          If Prd.Prd_Val(i) = " " Then
            U_temp(i, 1) = " "
            U_temp(i, 3) = Cnddts
          Else
            U_temp(i, 1) = Prd.Prd_Ini(i)
            U_temp(i, 3) = Cnddts_Blancs
          End If
        Next i
        Grid_Cdd_Remove_Cell_Coll(U_temp) ' Mise à jour des candidats éligibles
      Case "S"
        For i As Integer = 0 To 80
          U_temp(i, 1) = Prd.Prd_Ini(i)
          U_temp(i, 2) = Prd.Prd_Val(i)
          U_temp(i, 3) = Prd.Prd_Candidats(i)
        Next i
    End Select

    ' Contrôle de la grille et arrêt si incorrecte
    If Not Pzzl_Slv_ControlGrille(U_temp, Prd) Then Exit Sub

    ' Analyse des candidats
    Do
      Strategy_Rslt = Strategy_CdU(U_temp)
      If Pzzl_Slv_AnalyseStrategie_CdU_CdO("CdU", Production_Type, U_temp, Prd, Strategy_Rslt, Cellules_Type, Strategy_CdU_Nb) Then Continue Do

      Strategy_Rslt = Strategy_CdO(U_temp)
      If Pzzl_Slv_AnalyseStrategie_CdU_CdO("CdO", Production_Type, U_temp, Prd, Strategy_Rslt, Cellules_Type, Strategy_CdO_Nb) Then Continue Do

      If Pzzl_Slv_AnalyseComprehensive("Slv", Production_Type, U_temp, Prd) Then Continue Do

      Exit Do
    Loop

    ' Finalisation
    With Prd
      If .Prd_Code_Retour <> -1 Then
        .Prd_Code_Retour = If(Wh_Nb_Cell(U_temp).Vides = 0, 0, 1)
      End If
      For i As Integer = 0 To 80
        .Prd_Ini(i) = U_temp(i, 1)
        .Prd_Val(i) = U_temp(i, 2)
        .Prd_Candidats(i) = U_temp(i, 3)
      Next i

    End With
  End Sub

  Private Function Pzzl_Slv_ControlGrille(ByRef U_temp(,) As String, ByRef Prd As Prd_Struct) As Boolean
    Dim U_Chk(80, 3) As String
    Array.Copy(U_temp, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    If Not U_Check.Check Then
      If Prd.Prd_Chat Then
        U_Checking_Display(U_Check, Prd.Prd_Chat)
      End If
      Prd.Prd_Code_Retour = -1
      Return False
    End If
    Return True
  End Function

  Private Function Pzzl_Slv_AnalyseStrategie_CdU_CdO(StrgUO As String, Production_Type As String, ByRef U_temp(,) As String, ByRef Prd As Prd_Struct, Strategy_Rslt(,) As String, Cellules_Type As String, ByRef Strategy_Nb As Integer) As Boolean
    Select Case Cellules_Type
      Case "*One"
        If Strategy_Rslt.GetUpperBound(1) <> 0 Then Return False
      Case "*All"
        Strategy_Nb = Strategy_Upd_UO(Production_Type, U_temp, Strategy_Rslt, Prd)
        If StrgUO = "CdU" Then Prd.Slv_Strg_Nb(0) += Strategy_Nb
        If StrgUO = "CdO" Then Prd.Slv_Strg_Nb(1) += Strategy_Nb
    End Select
    Return Strategy_Nb > 0
  End Function

  Private Function Pzzl_Slv_AnalyseComprehensive(Type As String, Production_Type As String, ByRef U_temp(,) As String, ByRef Prd As Prd_Struct) As Boolean
    For i As Integer = 2 To 10
      If i = 10 AndAlso Production_Type <> "S" Then Exit For
      If Strategy_Upd_BTXYSJZKQ(Type, Production_Type, i, U_temp, Prd) > 0 Then Return True
    Next i
    Return False
  End Function

  Public Function Strategy_Upd_UO(Production_Type As String, ByRef U_temp(,) As String, Strategy_Rslt(,) As String, ByRef Prd As Prd_Struct) As Integer
    'Application des Mises à jour des Stratégies CdU_CdO 
    Dim Strategy_Nb As Integer

    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      Dim Cellule_CdU_CdO As Integer = CInt(Strategy_Rslt(10, i))
      Dim Candidat As String = Strategy_Rslt(5, i)
      Dim Candidats_av As String = U_temp(Cellule_CdU_CdO, 3)
      U_temp(Cellule_CdU_CdO, 2) = Candidat
      Dim Valeur As String = U_temp(Cellule_CdU_CdO, 2)
      U_temp(Cellule_CdU_CdO, 3) = Cnddts_Blancs
      Dim Candidats_ap As String = U_temp(Cellule_CdU_CdO, 3)
      Strategy_Nb += 1
      Nsd_i = Cdd_Remove_Cell_Coll(U_temp, Cellule_CdU_CdO)
      Jrn_Add_Pzzl_Slv(Production_Type, U_temp, Strategy_Rslt, i, Prd, Cellule_CdU_CdO, Valeur, Candidat, Candidats_av, Candidats_ap, Nsd_i)
    Next i
    Return Strategy_Nb
  End Function

  Public Function Strategy_Upd_BTXYSJZKQ(Origine As String, Production_Type As String, Strg As Integer, ByRef U_temp(,) As String, ByRef Prd As Prd_Struct) As Integer
    'Utilisation dans Pzzl_Crt ET Pzzl_Slv SAUF pour Unq qui n'est utilisé QUE pour Pzzl_Slv S'il est Production_Type = "S"
    'Exécution et Application des Mises à jour des Stratégies BTXYSJZKQ 
    'Un candidat est enlevé de U_temp(Cellule, 3)
    Dim Strategy_Nb As Integer = 0
    If Prd.Prd_Plcy_Strg_Bll(Strg) Then
      Dim Strategy_Rslt(99, 0) As String
      Select Case Strg
        'ase 0 : Strategy_Rslt = Strategy_CdU(U_temp)   ' U
        'ase 1 : Strategy_Rslt = Strategy_CdO(U_temp)   ' O
        Case 2 : Strategy_Rslt = Strategy_Cbl(U_temp)   ' B
        Case 3 : Strategy_Rslt = Strategy_Tpl(U_temp)   ' T
        Case 4 : Strategy_Rslt = Strategy_Xwg(U_temp)   ' X
        Case 5 : Strategy_Rslt = Strategy_XYw(U_temp)   ' Y
        Case 6 : Strategy_Rslt = Strategy_Swf(U_temp)   ' S
        Case 7 : Strategy_Rslt = Strategy_Jly(U_temp)   ' J
        Case 8 : Strategy_Rslt = Strategy_XYZ(U_temp)   ' Z
        Case 9 : Strategy_Rslt = Strategy_SKy(U_temp)   ' K
        Case 10 : Strategy_Rslt = Strategy_Unq(U_temp)  ' Q
      End Select
      'Pour les stratégies Cbl, Tpl, Xwg, XYw, Swf, Jly, XYZ, SKy
      '                    on a Dim Candidat As String = Strategy_Rslt(5, i)
      'Pour la stratégie   Unq
      '                    on a Dim Candidats As String = Strategy_Rslt(5, i)
      Select Case Strg
        Case 2, 3, 4, 5, 6, 7, 8, 9
          For i As Integer = 1 To UBound(Strategy_Rslt, 2)
            For j As Integer = 55 To 99
              If Strategy_Rslt(j, i) = "__" Then Exit For
              If Strategy_Rslt(j, i) <> "__" Then
                Dim Candidat As String = Strategy_Rslt(5, i)
                Dim Cellule As Integer = CInt(Strategy_Rslt(j, i))
                If U_temp(Cellule, 3).Contains(Candidat) = True Then
                  Dim Candidats_av As String = U_temp(Cellule, 3)
                  Dim Valeur As String = " "
                  Dim Candidats As String = U_temp(Cellule, 3)
                  Mid$(Candidats, CInt(Candidat), 1) = " "
                  U_temp(Cellule, 3) = Candidats
                  Dim Candidats_ap As String = U_temp(Cellule, 3)
                  Jrn_Add_Pzzl_Slv(Production_Type, U_temp, Strategy_Rslt, i, Prd, Cellule, Valeur, Candidat, Candidats_av, Candidats_ap, Nsd_i)
                  Strategy_Nb += 1
                End If
              End If
            Next j
          Next i

        Case 10
          For i As Integer = 1 To UBound(Strategy_Rslt, 2)
            For j As Integer = 55 To 99
              If Strategy_Rslt(j, i) = "__" Then Exit For
              If Strategy_Rslt(j, i) <> "__" Then
                Dim Candidats_Unq As String = Strategy_Rslt(5, i)
                For c As Integer = 1 To 9
                  If Mid$(Candidats_Unq, c, 1) = CStr(c) Then
                    Dim Candidat_Unq As String = CStr(c)
                    Dim Cellule As Integer = CInt(Strategy_Rslt(j, i))
                    If U_temp(Cellule, 3).Contains(Candidat_Unq) = True Then
                      Dim Candidats_av As String = U_temp(Cellule, 3)
                      Dim Valeur As String = " "
                      Dim Candidats As String = U_temp(Cellule, 3)
                      Mid$(Candidats, CInt(Candidat_Unq), 1) = " "
                      U_temp(Cellule, 3) = Candidats
                      Dim Candidats_ap As String = U_temp(Cellule, 3)
                      Jrn_Add_Pzzl_Slv(Production_Type, U_temp, Strategy_Rslt, i, Prd, Cellule, Valeur, Candidat_Unq, Candidats_av, Candidats_ap, Nsd_i)
                      Strategy_Nb += 1
                    End If
                  End If
                Next c
              End If
            Next j
          Next i
      End Select
      ' Pour une stratégie d'éviction, la fonction enlève le candidat dans les cellules concernées
      Select Case Origine
        Case "Crt" : Prd.Crt_Strg_Nb(Strg) += Strategy_Nb
        Case "Slv" : Prd.Slv_Strg_Nb(Strg) += Strategy_Nb
      End Select
    End If
    Return Strategy_Nb
  End Function

  Public Sub Jrn_Add_Pzzl_Slv(ByVal Production_Type As String, U_temp(,) As String, Strategy_Rslt(,) As String, Rslt_Ligne As Integer, Prd As Prd_Struct,
                              Cellule As Integer, Valeur As String,
                              Candidat As String, Candidats_av As String, Candidats_ap As String, Nsd_i As Integer)
    ' Mode d'emploi
    ' Col  1 Le Nombre de Cellules Remplies
    ' Col  2 Les Coordonnées de la cellule traitée 
    ' Col  3 La Stratégie et la Sous-Stratégie
    ' Col  4 La Valeur placée
    ' Col  5 Le Candidat placé ou exclu
    ' Col  6 Les candidats AVANT
    ' Col  7 Les candidats APRES
    ' Col  8 Le nombre de candidats exclus dans les cellules collatérales
    ' Col  9 Le Nombre total de candidats dans la grille
    ' Col 10 Le Nombre de Cellules Vides
    'Pour tester avec plusieurs stratégies .3.8...7..48...1....745..69...74.951..4......7...2........1.....25.7...36.....7..

    'Strategy_Upd_UO & Strategy_Upd_BTXYSJZKQ
    If Production_Type = "S" And Prd.Prd_Chat Then
      Prv_Col1 = Wh_Grid_Nb_Cellules_Remplies(Prv_U_temp)
      Act_Col1 = Wh_Grid_Nb_Cellules_Remplies(U_temp)
      If Act_Col1 <> Prv_Col1 Then
        Col1 = CStr(Wh_Grid_Nb_Cellules_Remplies(U_temp)).PadLeft(3)
      Else
        Col1 = "   "
      End If
      Act_Col3 = Strategy_Rslt(1, Rslt_Ligne) & " " & Strategy_Rslt(2, Rslt_Ligne)
      If Act_Col3 <> Prv_Col3 Then
        Col3 = Strategy_Rslt(1, Rslt_Ligne) & " " & Strategy_Rslt(2, Rslt_Ligne).PadRight(10)
      Else
        Col3 = "              "
      End If
      If Nsd_i = 0 Then
        Col8 = " "
      Else
        Col8 = CStr(Nsd_i)
      End If
      Prv_Col10 = Wh_Grid_Nb_Cellules_Vides(Prv_U_temp)
      Act_Col10 = Wh_Grid_Nb_Cellules_Vides(U_temp)
      If Act_Col10 <> Prv_Col10 Then
        Col10 = CStr(Wh_Grid_Nb_Cellules_Vides(U_temp)).PadLeft(3)
      Else
        Col10 = "   "
      End If

      Jrn_Add(, {Col1 &
            "| " & U_Coord(Cellule) &
            " | " & Col3 &
            " | " & Valeur & " | " & Candidat &
            " | " & Candidats_av & " | " & Candidats_ap &
            " | " & Col8.PadLeft(3) &
            " | " & CStr(Wh_Grid_Nb_Candidats(U_temp)).PadLeft(5) &
            " | " & Col10})
      Array.Copy(U_temp, Prv_U_temp, UNbCopy)
      Prv_Col3 = Act_Col3
    End If
  End Sub

  Public Sub Pzzl_Prd_Interactif(ByVal Production_Type As String)
    'Production_Type = "P"
    'La production interactive génère 1 Grille dans un nombre limité de tentatives
    'Comme il existe un stock de puzzle, la production interactive n'a d'intêret qu'avec la
    'création en mode bavard pour suivre le traitement de création et de résolution

    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Tentative As Integer = 0
    ' Le nombre de tentative dépend de la valeur stockée dans Préférences / Création / Nombre limite de tentatives
    Dim Prc As Integer = 5

    Jrn_Add("SDK_00011", JourDateHeure())
    For i As Integer = 0 To 80
      U(i, 1) = " "
      U(i, 2) = " "
      U(i, 3) = Cnddts
    Next i
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()

    Jrn_Add(, {Proc_Name_Get()})
    Cursor.Current = Cursors.WaitCursor
    Frm_SDK.B_Info.Visible = False

    Dim Prd As Prd_Struct = Nothing

    'Il est probable qu'il faille plusieurs tentatives pour crééer une grille
Pzzl_Prd_Boucle:
    Prd_Init(Prd, U, "I")

    If Tentative = CInt(Create_Nb_Tentatives) Then
      Dim Message As String = "Désolé.. " & vbCrLf & "Il n'a pas été possible de créer un SuDoKu..."
      Jrn_Add(, {Message})
      GoTo Pzzl_Prd_End
    End If
    Tentative += 1
    Prc = CInt(((Tentative * 100) / CInt(Create_Nb_Tentatives)))
    If Prc >= 100 Then Prc = 100
    With Frm_SDK
      .B_Pourcentage.Text = CStr(Prc) & "%"
      .B_Pourcentage.Refresh()
    End With

    If Create_Chat Then
      Jrn_Add("SDK_Space")
      Jrn_Add("SDK_Space")
      Jrn_Add(, {"Tentatives            : " & CStr(Tentative) & "/" & Create_Nb_Tentatives})
      Jrn_Add(, {"Contrainte (absolue)  : " & Prd.Prd_Cnt_Type & CStr(Prd.Prd_Cnt_Valeur)})
    End If

    '   Crt : Construction d'une grille
    Pzzl_Crt(Production_Type, Prd)
    If Prd.Prd_Code_Retour <> 0 Then GoTo Pzzl_Prd_Boucle
    '  Prd.Prd_Code-Retour = 0, alors on tente de résoudre la grille
    Prd.Prd_Code_Retour = -2
    Dim Strategy_Rslt(99, 0) As String
    Pzzl_Slv(Production_Type, "*All", Prd, Strategy_Rslt)
    Jrn_Add(, {""})

    'Prd.Prd_Code_retour = -2 lors du lancement de Pzzl_Slv 
    '                      -1 U_Check.Check = False
    '                       0 Il ne reste plus de cellule vide
    '                       1 Il reste des cellules vides

    If Prd.Prd_Code_Retour = 1 Then
      Pzzl_Crt_Triplet(Prd)
      Pzzl_Crt_XWing(Prd)
      Prd.Prd_Code_Retour = -2
      Pzzl_Slv(Production_Type, "*All", Prd, Strategy_Rslt)
    End If

    If Prd.Prd_Code_Retour = 0 Then
      ' Il ne reste plus de cellules vides, Il FAUDRA que DLCode soit Dlu
      Dim U_temp(,) As String = Prd_To_U_temp(Prd)
      Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
      Prd.Prd_DlCode = DL.DLCode
      Prd.Prd_DlSolution = DL.Solution(0)
    End If

    If Prd.Prd_Code_Retour = 1 Then
      ' Il reste des cellules vides (Si DLCode = Dlu, alors les stratégies sont insuffisantes)
      Dim U_temp(,) As String = Prd_To_U_temp(Prd)
      Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
      Prd.Prd_DlCode = DL.DLCode
      Prd.Prd_DlSolution = DL.Solution(0)
      If DL.DLCode = "Dlu" Then Prd.Prd_Code_Retour = 0
    End If

    If Prd.Prd_Code_Retour <> 0 Then GoTo Pzzl_Prd_Boucle

Pzzl_Prd_End:
    Select Case Prd.Prd_Code_Retour
      Case 0
        Jrn_Add(, {Proc_Name_Get() & " /Fin Création d'un SuDoKu"})
        Dim File As String = Pzzl_Save(Prd)
        Pzzl_Open(File)
      Case Else
        Jrn_Add(, {Proc_Name_Get() & " /Fin sans réussite"})
    End Select

    Jrn_Add(, {"Tentatives            : " & CStr(Tentative) & "/" & Create_Nb_Tentatives})
    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée: " & CStr(Durée).PadLeft(5) & " ms, soit " & String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)})

    Cursor.Current = Cursors.Default
    Frm_SDK.B_Info.Visible = True
  End Sub

  Public Sub Pzzl_Crt_Interactif_81(ByVal Production_Type As String)
    'Production_Type = "P"
    'La création interactive génère 1 Grille dans un nombre limité de tentatives
    'l'intérêt réside dans la tentative d'une grille de 81 cellules
    '                 dans la vérification du mode chat.
    'Cette procédure n'est pas utilisée, elle est codée UNIQUEMENT pour vérifier Pzzl_Crt(Prd)
    '
    '
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Tentative As Integer = 0
    Dim Prc As Integer = 5

    Jrn_Add("SDK_00011", JourDateHeure())
    For i As Integer = 0 To 80
      U(i, 1) = " "
      U(i, 2) = " "
      U(i, 3) = Cnddts
    Next i
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()

    Jrn_Add(, {Proc_Name_Get()})
    Cursor.Current = Cursors.WaitCursor
    Frm_SDK.B_Info.Visible = False

    Dim Prd As Prd_Struct = Nothing

    'Il est probable qu'il faille plusieurs tentatives pour créer une grille
Pzzl_Prd_Boucle:
    If Tentative = CInt(Create_Nb_Tentatives) Then
      Dim Message As String = "Désolé.. " & vbCrLf & "Il n'a pas été possible de créer un SuDoKu..."
      Jrn_Add(, {Message})
      GoTo Pzzl_Prd_End
    End If

    Tentative += 1
    Prc = CInt(((Tentative * 100) / CInt(Create_Nb_Tentatives)))
    If Prc >= 100 Then Prc = 100
    With Frm_SDK
      .B_Pourcentage.Text = CStr(Prc) & "%"
      .B_Pourcentage.Refresh()
    End With

    Jrn_Add("SDK_Space")
    Jrn_Add("SDK_Space")
    Jrn_Add(, {"Tentatives            : " & CStr(Tentative) & "/" & Create_Nb_Tentatives})

    Prd_Init(Prd, U, "I")
    '//// Code ajouté /////////////////////////////////////////////////////////////////////////////
    Prd.Prd_Chat = True
    Prd.Prd_Create_Nb_Cel_Demandées = 81
    '//// Code ajouté /////////////////////////////////////////////////////////////////////////////

    '   Crt : Construction d'une grille
    Pzzl_Crt(Production_Type, Prd)
    Prd_Display(Prd)
    If Prd.Prd_Code_Retour <> 0 Then GoTo Pzzl_Prd_Boucle

    Dim Grille_Ini As String = ""
    Dim Grille_Val As String = ""
    Dim Grille_Sol As String = ""
    Dim Grille_Candidats As String = ""

    For i As Integer = 0 To 80
      Grille_Ini &= Prd.Prd_Ini(i)
      Grille_Val &= Prd.Prd_Val(i)
      Grille_Sol &= " "
      Grille_Candidats &= Prd.Prd_Candidats(i)
    Next i
    Game_Load(Proc_Name_Get(), Grille_Ini, Grille_Val, Grille_Sol, "1")
    For i As Integer = 0 To 80
      If U(i, 2) = " " Then
        U(i, 3) = Grille_Candidats.Substring(i * 9, 9).Replace("0", " ")
      End If
    Next i
    Jrn_Add(, {"Les candidats ne sont pas recalculés."}, "Info")
    Dim Wh_Nb As Wh_Nb_Cell_Struct = Wh_Nb_Cell(U)
    Wh_Nb_Cell_Display(Wh_Nb)
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()

Pzzl_Prd_End:
    Select Case Prd.Prd_Code_Retour
      Case 0
      Case Else
        Jrn_Add(, {Proc_Name_Get() & " /Fin sans réussite"})
    End Select
    Jrn_Add(, {"Tentatives            : " & CStr(Tentative) & "/" & Create_Nb_Tentatives})
    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée: " & CStr(Durée).PadLeft(5) & " ms, soit " & String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)})

    Cursor.Current = Cursors.Default
    Frm_SDK.B_Info.Visible = True

  End Sub

  Public Sub Pzzl_Prd_DL()
    ' 15/10/2024 Nouvelle création d'un Sudoku
    'A à/p d'une grille complète de 81 valeurs initiales,
    'B on enlève une cellule au hasard
    '  si Solve_DL est correct, on continue, sinon on remet la valeur
    '  Tant que le Préférences / Création / Nb de Cellules demandées n'est pas atteint on retourne en B
    '  Ce mode de création peut ne pas être résolvable par Résolution ! 
    Dim Limite As Integer = 0
    Dim Limite_Max As Integer = 100
    Dim Limite_Successive_Dl As Integer = 0
    Dim Limite_Successive_Dl_Max As Integer = My.Settings.Prf_02C_Nb_Max_Dl
    Dim Nb_VI_Origine As Integer
    Dim Cellule As Integer
    Dim Valeur As String
    Dim Sol As String = "" ' La solution est toujours la même

    Dim Prd As Prd_Struct
    Dim Strategy_Rslt(99, 0) As String
    Dim U_temp(80, 3) As String

    ' On travaille désormais sur U_Temp
    Array.Copy(U, U_temp, UNbCopy)
    Nb_VI_Origine = Wh_Grid_Nb_Cellules_Initiales(U_temp)
    'La grille doit être remplie, les 81 valeurs sont considérées comme VI
    If Wh_Grid_Nb_Cellules_Remplies(U_temp) < 81 Then
      Prd = Nothing
      Prd_Init(Prd, U_temp, "I")
      Prd.Prd_Chat = False
      Pzzl_Slv("S", "*All", Prd, Strategy_Rslt)
      Prd.Prd_Chat = Create_Chat

      ' Si la grille incomplète n'est pas solutionnée
      If Prd.Prd_Val.Contains(" ") Then
        Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
        Dim MsgTxt As String = "Le Puzzle n'est pas résolu par SDK ! Abandon "
        Jrn_Add(, {"Résolution                   : " & MsgTxt}, "Orange")
        Nsd_i = MsgBox(MsgTxt,, MsgTit)
        Exit Sub
      End If

      ' On complète U_Temp des valeurs solutionnées
      For i As Integer = 0 To 80
        U_temp(i, 1) = Prd.Prd_Ini(i)
        U_temp(i, 2) = Prd.Prd_Val(i)
        U_temp(i, 3) = Prd.Prd_Candidats(i)
      Next i
    End If

    ' A ce niveau le grille est complète
    ClipBoard_Coller_RTF()
    For i As Integer = 0 To 80
      U_temp(i, 1) = U_temp(i, 2)
      U_temp(i, 3) = Cnddts_Blancs
      Sol &= U_temp(i, 2)      ' Sol comporte l'ensemble des valeurs
    Next i

    Dim Cellules_Clct As New Collection      ' Collection des Cellules à traiter
Phase_A:
    ' La collection comporte d'abord les 81 cellules, puis celles qui ne sont pas initiales
    ' Je n'utilise pas l’algorithme de Fisher-Yates
    For i As Integer = 0 To 80
      If U_temp(i, 1) <> " " Then Clct_Add(Cellules_Clct, i)
    Next i
    If Create_Chat Then Jrn_Add(, {"Remplissage " & CStr(Cellules_Clct.Count)})

Phase_B:
    Limite += 1
    If Limite = 1 Then
      If Create_Chat Then Jrn_Add(, {"N°  Cellule V  Candidats      DL  VI "})
    End If
    Cellule = Clct_Random(Cellules_Clct)
    If Cellule = -1 Then GoTo Phase_A
    Valeur = U_temp(Cellule, 2)

    U_temp(Cellule, 1) = " "
    U_temp(Cellule, 2) = " "
    U_temp(Cellule, 3) = Cnddts
    Grid_Cdd_Remove_Cell_Coll(U_temp)
    Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
    Prd.Prd_DlCode = DL.DLCode
    Prd.Prd_DlSolution = DL.Solution(0)
    If Create_Chat Then Jrn_Add(, {CStr(Limite).PadLeft(3) & "  " & U_Coord(Cellule) & "  " & Valeur & "  " & U_temp(Cellule, 3) & "  " & CStr(DL.Nb_Solution).PadLeft(6) & "  " & Wh_Grid_Nb_Cellules_Initiales(U_temp)})
    If DL.Nb_Solution = 1 Then
      'Limite_Successive_DL = 0
    Else
      U_temp(Cellule, 1) = Valeur
      U_temp(Cellule, 2) = " "
      U_temp(Cellule, 3) = Cnddts
      Grid_Cdd_Remove_Cell_Coll(U_temp)
      Limite_Successive_Dl += 1
      If Create_Chat Then Jrn_Add(, {"   " & "  " & U_Coord(Cellule) & "  " & Valeur & "  " & U_temp(Cellule, 3) & "               Limite_Successive_DL =" & CStr(Limite_Successive_Dl)})
    End If

    Dim Limite_Atteinte As Boolean = False
    Dim Limite_Message As String = ""
    If Limite > Limite_Max Then
      Limite_Atteinte = True : Limite_Message = Msg_Read("PRD_30010")
    End If
    If Limite_Successive_Dl > Limite_Successive_Dl_Max Then
      Limite_Atteinte = True : Limite_Message = Msg_Read("PRD_30020")
    End If
    If Wh_Grid_Nb_Cellules_Initiales(U_temp) = CInt(Create_Nb_Cel_Demandées) Then
      Limite_Atteinte = True : Limite_Message = Msg_Read("PRD_30030")
    End If

    If Limite > Limite_Max _
    Or Limite_Successive_Dl > Limite_Successive_Dl_Max _
    Or Wh_Grid_Nb_Cellules_Initiales(U_temp) = CInt(Create_Nb_Cel_Demandées) Then GoTo Phase_End
    GoTo Phase_B

Phase_End:
    'On travaille désormais sur U 
    Array.Copy(U_temp, U, UNbCopy)
    Dim Prb As String = ""
    Dim Jeu As String = ""
    For i As Integer = 0 To 80
      U(i, 2) = U(i, 1)
      If U(i, 1) = " " Then U(i, 3) = Cnddts
      If U(i, 1) <> " " Then U(i, 3) = Cnddts_Blancs
      Prb &= U(i, 1)
      Jeu &= U(i, 1)
    Next i
    Grid_Cdd_Remove_Cell_Coll(U)

    '   Sol n'a pas changé

    ' La vérification permet de vérifier le niveau de résolution de la grille
    ' Production_Type est positionné "S"

    If Create_Chat Then Jrn_Add("SDK_Space")
    Prd = Nothing
    Prd_Init(Prd, U, "I")
    ReDim Strategy_Rslt(99, 0)

    Pzzl_Slv("S", "*All", Prd, Strategy_Rslt)
    Dim DL_Nom As DL_Solve_Struct = A_Copyright.DL_Solve(U)
    Prd.Prd_DlCode = DL_Nom.DLCode
    Prd.Prd_DlSolution = DL_Nom.Solution(0)

    'Compte-rendu
    If Create_Chat And Limite_Atteinte Then Jrn_Add(, {Limite_Message})
    If Create_Chat Then Jrn_Add(, {"Limite       : " & CStr(Limite).PadLeft(3) & "/" & CStr(Limite_Max)})
    If Create_Chat Then Jrn_Add(, {"Limite_DL    : " & CStr(Limite_Successive_Dl).PadLeft(3) & "/" & CStr(Limite_Successive_Dl_Max)})
    If Create_Chat Then Jrn_Add(, {"Nb VI        : " & CStr(Nb_VI_Origine).PadLeft(3) & "(Origine)/" &
                                                       CStr(Wh_Grid_Nb_Cellules_Initiales(U_temp)).PadLeft(3) & "(Atteint)/" &
                                                       Create_Nb_Cel_Demandées.PadLeft(3) & "(Objectif)"})
    Select Case Prd.Prd_Code_Retour
      Case 0
        Jrn_Add(, {"Résolution   : " & "Le Puzzle est résolu par SDK."})
      Case 1
        Jrn_Add(, {"Résolution   : " & "Le Puzzle n'est pas résolu par SDK ! "}, "Italique")
        ' Dans ce cas la solution ne comporte que les cellules résolues, on aura donc
        ' de nombreuses cellules en erreur non conforme à la solution ... de SDK!
      Case Else
        Jrn_Add(, {"Résolution   : " & "Problème de résolution,  Prd.Prd_Code_Retour=" & Prd.Prd_Code_Retour}, "Erreur")
    End Select

    Prd_Display(Prd)
    Dim File As String = Pzzl_Save(Prd)
    Pzzl_Open(File)
  End Sub



  Public Sub Pzzl_Prd_DL_Save()
    ' 15/10/2024 Nouvelle création d'un Sudoku
    'A à/p d'une grille complète de 81 valeurs initiales,
    'B on enlève une cellule au hasard
    '  si Solve_DL est correct, on continue, sinon on remet la valeur
    '  Tant que le Préférences / Création / Nb de Cellules demandées n'est pas atteint on retourne en B
    '  Ce mode de création peut ne pas être résolvable par Résolution ! 
    Dim Limite As Integer = 0
    Dim Limite_Max As Integer = 100
    Dim Limite_Successive_Dl As Integer = 0
    Dim Limite_Successive_Dl_Max As Integer = My.Settings.Prf_02C_Nb_Max_Dl
    Dim Nb_VI_Origine As Integer
    Dim Cellule As Integer
    Dim Valeur As String
    Dim Sol As String = "" ' La solution est toujours la même

    Dim Prd As Prd_Struct
    Dim Strategy_Rslt(99, 0) As String
    Dim U_temp(80, 3) As String

    ' On travaille désormais sur U_Temp
    Array.Copy(U, U_temp, UNbCopy)
    Nb_VI_Origine = Wh_Grid_Nb_Cellules_Initiales(U_temp)
    'La grille doit être remplie, les 81 valeurs sont considérées comme VI
    If Wh_Grid_Nb_Cellules_Remplies(U_temp) < 81 Then
      Prd = Nothing
      Prd_Init(Prd, U_temp, "I")
      Prd.Prd_Chat = False
      Pzzl_Slv("S", "*All", Prd, Strategy_Rslt)
      Prd.Prd_Chat = Create_Chat

      ' Si la grille incomplète n'est pas solutionnée
      If Prd.Prd_Val.Contains(" ") Then
        Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
        Dim MsgTxt As String = "Le Puzzle n'est pas résolu par SDK ! Abandon "
        Jrn_Add(, {"Résolution                   : " & MsgTxt}, "Orange")
        Nsd_i = MsgBox(MsgTxt,, MsgTit)
        Exit Sub
      End If

      ' On complète U_Temp des valeurs solutionnées
      For i As Integer = 0 To 80
        U_temp(i, 1) = Prd.Prd_Ini(i)
        U_temp(i, 2) = Prd.Prd_Val(i)
        U_temp(i, 3) = Prd.Prd_Candidats(i)
      Next i
    End If

    ' A ce niveau le grille est complète
    ClipBoard_Coller_RTF()
    For i As Integer = 0 To 80
      U_temp(i, 1) = U_temp(i, 2)
      U_temp(i, 3) = Cnddts_Blancs
      Sol &= U_temp(i, 2)      ' Sol comporte l'ensemble des valeurs
    Next i

    Dim Cellules_Clct As New Collection      ' Collection des Cellules à traiter
Phase_A:
    ' La collection comporte d'abord les 81 cellules, puis celles qui ne sont pas initiales
    For i As Integer = 0 To 80
      If U_temp(i, 1) <> " " Then Clct_Add(Cellules_Clct, i)
    Next i
    If Create_Chat Then Jrn_Add(, {"Remplissage " & CStr(Cellules_Clct.Count)})

Phase_B:
    Limite += 1
    If Limite = 1 Then
      If Create_Chat Then Jrn_Add(, {"N°  Cellule V  Candidats      DL  VI "})
    End If
    Cellule = Clct_Random(Cellules_Clct)
    If Cellule = -1 Then GoTo Phase_A
    Valeur = U_temp(Cellule, 2)

    U_temp(Cellule, 1) = " "
    U_temp(Cellule, 2) = " "
    U_temp(Cellule, 3) = Cnddts
    Grid_Cdd_Remove_Cell_Coll(U_temp)
    Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
    If Create_Chat Then Jrn_Add(, {CStr(Limite).PadLeft(3) & "  " & U_Coord(Cellule) & "  " & Valeur & "  " & U_temp(Cellule, 3) & "  " & CStr(DL.Nb_Solution).PadLeft(6) & "  " & Wh_Grid_Nb_Cellules_Initiales(U_temp)})
    If DL.Nb_Solution = 1 Then
      'Limite_Successive_DL = 0
    Else
      U_temp(Cellule, 1) = Valeur
      U_temp(Cellule, 2) = " "
      U_temp(Cellule, 3) = Cnddts
      Grid_Cdd_Remove_Cell_Coll(U_temp)
      Limite_Successive_Dl += 1
      If Create_Chat Then Jrn_Add(, {"   " & "  " & U_Coord(Cellule) & "  " & Valeur & "  " & U_temp(Cellule, 3) & "               Limite_Successive_DL =" & CStr(Limite_Successive_Dl)})
    End If

    Dim Limite_Atteinte As Boolean = False
    Dim Limite_Message As String = ""
    If Limite > Limite_Max Then
      Limite_Atteinte = True : Limite_Message = Msg_Read("PRD_30010")
    End If
    If Limite_Successive_Dl > Limite_Successive_Dl_Max Then
      Limite_Atteinte = True : Limite_Message = Msg_Read("PRD_30020")
    End If
    If Wh_Grid_Nb_Cellules_Initiales(U_temp) = CInt(Create_Nb_Cel_Demandées) Then
      Limite_Atteinte = True : Limite_Message = Msg_Read("PRD_30030")
    End If

    If Limite > Limite_Max _
    Or Limite_Successive_Dl > Limite_Successive_Dl_Max _
    Or Wh_Grid_Nb_Cellules_Initiales(U_temp) = CInt(Create_Nb_Cel_Demandées) Then GoTo Phase_End
    GoTo Phase_B

Phase_End:



    ' On travaille désormais sur U 
    Array.Copy(U_temp, U, UNbCopy)
    Dim Prb As String = ""
    Dim Jeu As String = ""
    For i As Integer = 0 To 80
      U(i, 2) = U(i, 1)
      If U(i, 1) = " " Then U(i, 3) = Cnddts
      If U(i, 1) <> " " Then U(i, 3) = Cnddts_Blancs
      Prb &= U(i, 1)
      Jeu &= U(i, 1)
    Next i
    Grid_Cdd_Remove_Cell_Coll(U)

    '   Sol n'a pas changé
    Dim Frc As String = "0"
    ' A ce niveau, la partie doit avoir un nom
    ' et           la partie doit être résolvable par DL, Pzzl_Slv 
    ' Composition du Nom
    Dim Nom As String = "Prd" & "_"

    Dim Difficulté As String
    Dim Nb_VI As Integer = 0
    For i As Integer = 0 To 80
      If U(i, 1) <> " " Then Nb_VI += 1
    Next i
    Select Case Nb_VI
      Case 0 To 21 : Difficulté = "I"      ' Impossible
      Case 22 To 25 : Difficulté = "D"     ' Difficile
      Case 26 To 29 : Difficulté = "M"     ' Moyenne
      Case 30 To 35 : Difficulté = "F"     ' Facile
      Case Else : Difficulté = "A"         ' Autre
    End Select

    ' La vérification permet de vérifier le niveau de résolution de la grille
    ' Production_Type est positionné "S"

    If Create_Chat Then Jrn_Add("SDK_Space")
    Prd = Nothing
    Prd_Init(Prd, U, "I")
    ReDim Strategy_Rslt(99, 0)
    Pzzl_Slv("S", "*All", Prd, Strategy_Rslt)
    Prd_Display(Prd)
    Dim Difficulté_Strg_Slv As String = "#"
    For i As Integer = 10 To 0 Step -1
      If Prd.Slv_Strg_Nb(i) <> 0 Then
        Difficulté_Strg_Slv = Stg_List_Lettre(i)
        Exit For
      End If
    Next i
    Nom &= "_" & Difficulté & CStr(Nb_VI) & "~" & Difficulté_Strg_Slv
    Dim DL_Nom As DL_Solve_Struct = A_Copyright.DL_Solve(U)
    Nom &= "_" & DL_Nom.DLCode
    Nom &= "_" & Format(Now, "yyyy_MM_dd") & "_" & Format(Now, "HH_mm_ss")

    'Compte-rendu
    If Create_Chat Then Jrn_Add(, {"Nom          : " & Nom})
    If Create_Chat And Limite_Atteinte Then Jrn_Add(, {Limite_Message})
    If Create_Chat Then Jrn_Add(, {"Limite       : " & CStr(Limite).PadLeft(3) & "/" & CStr(Limite_Max)})
    If Create_Chat Then Jrn_Add(, {"Limite_DL    : " & CStr(Limite_Successive_Dl).PadLeft(3) & "/" & CStr(Limite_Successive_Dl_Max)})
    If Create_Chat Then Jrn_Add(, {"Nb VI        : " & CStr(Nb_VI_Origine).PadLeft(3) & "(Origine)/" &
                                                       CStr(Wh_Grid_Nb_Cellules_Initiales(U_temp)).PadLeft(3) & "(Atteint)/" &
                                                       Create_Nb_Cel_Demandées.PadLeft(3) & "(Objectif)"})
    Select Case Prd.Prd_Code_Retour
      Case 0
        Jrn_Add(, {"Résolution   : " & "Le Puzzle est résolu par SDK."})
      Case 1
        Jrn_Add(, {"Résolution   : " & "Le Puzzle n'est pas résolu par SDK ! "}, "Italique")
        ' Dans ce cas la solution ne comporte que les cellules résolues, on aura donc
        ' de nombreuses cellules en erreur non conforme à la solution ... de SDK!
      Case Else
        Jrn_Add(, {"Résolution   : " & "Problème de résolution,  Prd.Prd_Code_Retour=" & Prd.Prd_Code_Retour}, "Erreur")
    End Select

    Game_New_Game(Plcy_Gnrl, "   ", Nom, Prb, Jeu, Sol, StrDup(729, " "), Frc, Proc_Name_Get())
  End Sub




  Public Sub Pzzl_Slv_Interactif(ByVal Production_Type As String)
    Jrn_Add(, {Proc_Name_Get()})
    ' Production_Type est positionné "S"
    Dim Prd As Prd_Struct = Nothing
    Prd_Init(Prd, U, "I")
    Dim Strategy_Rslt(99, 0) As String
    Pzzl_Slv(Production_Type, "*All", Prd, Strategy_Rslt)
    Prd_Display(Prd)

    For i As Integer = 0 To 80
      U(i, 1) = Prd.Prd_Ini(i)
      U(i, 2) = Prd.Prd_Val(i)
      U(i, 3) = Prd.Prd_Candidats(i)
    Next i

    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub

  Public Sub Pzzl_Crt_Triplet(ByRef Prd As Prd_Struct)
    'Le calcul de Triplet prend en compte 3 cellules vides par Ligne, Colonne ou Région
    '                     dont une seule avec 3 candidats, les 2 autres avec 2 seulement
    Dim Triplet As Boolean = False
    Dim CelA, CelB, CelC As Integer
    Dim U_temp(80, 3) As String
    'Documentation de U_temp à/p de Prd
    With Prd
      For i As Integer = 0 To 80
        U_temp(i, 1) = .Prd_Ini(i)
        U_temp(i, 2) = .Prd_Val(i)
        U_temp(i, 3) = .Prd_Candidats(i)
      Next i
      .Prd_Code_Retour = -1
    End With
    Dim Grp(0 To 8) As Integer
    'Analyse des Lignes, Colonnes, Régions
    For LCR As Integer = 0 To 8
      If Triplet Then Exit For
      Grp = U_9CelRow(LCR)
      Pzzl_Crt_Triplet_Exists(U_temp, U_9CelRow(LCR), CelA, CelB, CelC, Triplet)
      If Triplet Then Exit For
      Pzzl_Crt_Triplet_Exists(U_temp, U_9CelCol(LCR), CelA, CelB, CelC, Triplet)
      If Triplet Then Exit For
      Pzzl_Crt_Triplet_Exists(U_temp, U_9CelReg(LCR), CelA, CelB, CelC, Triplet)
      If Triplet Then Exit For
    Next LCR

    If Triplet Then
      'Mise en place d'un candidat au hasard dans une des 3 cellules au hasard
      Dim Cel_Clct As New Collection          ' Collection des 3 cellules Triplet
      'Les 3 cellules sont ajoutées
      Clct_Add(Cel_Clct, CelA)
      Clct_Add(Cel_Clct, CelB)
      Clct_Add(Cel_Clct, CelC)

      'En fonction des contraintes d'exclusion, certaines cellules doivent être enlevées
      If Prd.Prd_Cnt_Excl Then
        Select Case Prd.Prd_Cnt_Type
          Case "L"
            If U_Row(CelA) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelA)
            If U_Row(CelB) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelB)
            If U_Row(CelC) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelC)
          Case "C"
            If U_Col(CelA) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelA)
            If U_Col(CelB) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelB)
            If U_Col(CelC) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelC)
          Case "R"
            If U_Reg(CelA) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelA)
            If U_Reg(CelB) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelB)
            If U_Reg(CelC) = Prd.Prd_Cnt_Valeur Then Clct_Remove(Cel_Clct, CelC)
          Case "V" ' Rien
        End Select
      End If

      If Cel_Clct.Count = 0 Then GoTo Triplet_End
      Dim Cel As Integer = Clct_Random(Cel_Clct)

      Dim Cdds As String = U_temp(Cel, 3)
      Dim Cdd_Clct As New Collection          ' Collection des 2 candidats
      'Les candidats sont ajoutés
      For j As Integer = 1 To 9
        If Mid$(Cdds, j, 1) <> " " Then Clct_Add(Cdd_Clct, j)
      Next j

      If Prd.Prd_Cnt_Excl Then
        If Prd.Prd_Cnt_Type = "V" Then
          Clct_Remove(Cdd_Clct, Prd.Prd_Cnt_Valeur)
        End If
      End If
      If Cdd_Clct.Count = 0 Then GoTo Triplet_End
      Dim Cdd As String = CStr(Clct_Random(Cdd_Clct))

      'La saisie d'une valeur dans cellule comporte ces 4 opérations
      U_temp(Cel, 1) = Cdd
      U_temp(Cel, 2) = Cdd
      U_temp(Cel, 3) = Cnddts_Blancs
      'Dim Cell_Coll_Nb As Integer = Cdd_Remove_Cell_Coll(U_temp, Cel)
      Cdd_Remove_Cell_Coll(U_temp, Cel)
      Prd.Prd_Ext_Triplet_Cellule = Cel
Triplet_End:
    End If

    'Documentation de Prd
    With Prd
      If Triplet Then
        .Prd_Code_Retour = 0
        .Prd_Ext_Triplet = "T"
        For i As Integer = 0 To 80
          .Prd_Ini(i) = U_temp(i, 1)
          .Prd_Val(i) = U_temp(i, 2)
          .Prd_Candidats(i) = U_temp(i, 3)
        Next i
      End If
    End With

  End Sub

  Public Sub Pzzl_Crt_Triplet_Exists(ByRef U_temp(,) As String, Grp() As Integer, ByRef CelA As Integer, ByRef CelB As Integer, ByRef CelC As Integer, ByRef Triplet As Boolean)
    Triplet = False
    Dim n As Integer = 0
    'Il faut d'abord SEULEMENT 3 cellules vides
    For g As Integer = 0 To 8
      Dim Cellule As Integer = Grp(g)
      If U_temp(Cellule, 2) = " " Then n += 1
    Next g
    If n <> 3 Then Exit Sub
    'Il faut ensuite que ces 3 cellules n'aient que 2 candidats
    n = 0
    For g As Integer = 0 To 8
      Dim Cellule As Integer = Grp(g)
      If U_temp(Cellule, 2) = " " And Wh_Cell_Nb_Candidats(U_temp, Cellule) = 2 Then
        n += 1
        'Par anticipation
        If n = 1 Then CelA = Cellule
        If n = 2 Then CelB = Cellule
        If n = 3 Then CelC = Cellule
      End If
    Next g
    If n <> 3 Then Exit Sub
    Triplet = True
  End Sub

  Public Sub Pzzl_Crt_XWing(ByRef Prd As Prd_Struct)
    'Le calcul du Xwing prend en compte un seul Xwing RECTANGULAIRE
    Dim XWing As Boolean = False
    Dim U_temp(80, 3) As String
    'Documentation de U_temp à/p de Prd
    With Prd
      For i As Integer = 0 To 80
        U_temp(i, 1) = .Prd_Ini(i)
        U_temp(i, 2) = .Prd_Val(i)
        U_temp(i, 3) = .Prd_Candidats(i)
      Next i
      .Prd_Code_Retour = -1
    End With

    Dim CelA, CelB, CelC, CelD As Integer
    Dim Cdds As String

    For i As Integer = 0 To 80
      If XWing Then Exit For
      ' Recherche de la première cellule vide avec 2 candidats
      If U_temp(i, 2) = " " And Wh_Cell_Nb_Candidats(U_temp, i) = 2 Then
        CelA = i 'Cellule A en haut à gauche
        Cdds = U_temp(CelA, 3)
        'Le XWing est calculé de Haut-Gauche à Bas-Droit
        'Existe-t'il une cellule B sur la même rangée ayant les mêmes candidats
        For c As Integer = U_Col(CelA) + 1 To 8
          If XWing Then Exit For
          CelB = Wh_Cellule_ColRow(c, U_Row(CelA))
          If U_temp(CelB, 3) = Cdds Then
            'Existe-t'il une cellule C sur la même colonne ayant les mêmes candidats
            For r As Integer = U_Row(CelA) + 1 To 8
              CelC = Wh_Cellule_ColRow(U_Col(CelA), r)
              If U_temp(CelC, 3) = Cdds Then
                'Existe-t'il une cellule D en Xwing ayant les mêmes candidats
                CelD = Wh_Cellule_ColRow(U_Col(CelB), U_Row(CelC))
                If U_temp(CelD, 3) = Cdds Then
                  'Le XWing rectangulaire est trouvé
                  'Mise en place d'un candidat au hasard dans une des 4 cellules au hasard
                  'Ajout systématique des Cellules
                  Dim Cel_Clct As New Collection          ' Collection des 4 cellules Xwing

                  If Prd.Prd_Cnt_None _
                  Or Prd.Prd_Cnt_Sym Then
                    Clct_Add(Cel_Clct, CelA)
                    Clct_Add(Cel_Clct, CelB)
                    Clct_Add(Cel_Clct, CelC)
                    Clct_Add(Cel_Clct, CelD)
                  End If

                  If Prd.Prd_Cnt_Excl Then ' (And Contrainte_Exclusion = True)
                    If Prd.Prd_Cnt_Type = "L" Then
                      If U_Row(CelA) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelA)
                      If U_Row(CelB) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelB)
                      If U_Row(CelC) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelC)
                      If U_Row(CelD) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelD)
                    End If
                    If Prd.Prd_Cnt_Type = "C" Then
                      If U_Col(CelA) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelA)
                      If U_Col(CelB) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelB)
                      If U_Col(CelC) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelC)
                      If U_Col(CelD) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelD)
                    End If
                    If Prd.Prd_Cnt_Type = "R" Then
                      If U_Reg(CelA) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelA)
                      If U_Reg(CelB) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelB)
                      If U_Reg(CelC) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelC)
                      If U_Reg(CelD) <> Prd.Prd_Cnt_Valeur Then Clct_Add(Cel_Clct, CelD)
                    End If
                    If Prd.Prd_Cnt_Type = "V" Then
                      Clct_Add(Cel_Clct, CelA)
                      Clct_Add(Cel_Clct, CelB)
                      Clct_Add(Cel_Clct, CelC)
                      Clct_Add(Cel_Clct, CelD)
                    End If
                  End If

                  If Cel_Clct.Count = 0 Then GoTo Xwing_Boucle_End
                  Dim Cel As Integer = Clct_Random(Cel_Clct)

                  Dim Cdd_Clct As New Collection          ' Collection des 2 candidats

                  If Prd.Prd_Cnt_None _
                  Or Prd.Prd_Cnt_Sym Then
                    For j As Integer = 1 To 9
                      If Mid$(Cdds, j, 1) <> " " Then Clct_Add(Cdd_Clct, j)
                    Next j
                  End If

                  If Prd.Prd_Cnt_Excl Then
                    If Prd.Prd_Cnt_Type = "L" Then
                      For j As Integer = 1 To 9
                        If Mid$(Cdds, j, 1) <> " " Then Clct_Add(Cdd_Clct, j)
                      Next j
                    End If
                    If Prd.Prd_Cnt_Type = "C" Then
                      For j As Integer = 1 To 9
                        If Mid$(Cdds, j, 1) <> " " Then Clct_Add(Cdd_Clct, j)
                      Next j
                    End If
                    If Prd.Prd_Cnt_Type = "R" Then
                      For j As Integer = 1 To 9
                        If Mid$(Cdds, j, 1) <> " " Then Clct_Add(Cdd_Clct, j)
                      Next j
                    End If
                    If Prd.Prd_Cnt_Type = "V" Then
                      For j As Integer = 1 To 9
                        If Mid$(Cdds, j, 1) <> " " And Mid$(Cdds, j, 1) <> CStr(Prd.Prd_Cnt_Valeur) Then Clct_Add(Cdd_Clct, j)
                      Next j
                    End If
                  End If
                  If Cdd_Clct.Count = 0 Then GoTo Xwing_Boucle_End
                  Dim Cdd As String = CStr(Clct_Random(Cdd_Clct))

                  'La saisie d'une valeur dans cellule comporte ces 4 opérations
                  U_temp(Cel, 1) = Cdd
                  U_temp(Cel, 2) = Cdd
                  U_temp(Cel, 3) = Cnddts_Blancs
                  'Dim Cell_Coll_Nb As Integer = Cdd_Remove_Cell_Coll(U_temp, Cel)
                  Cdd_Remove_Cell_Coll(U_temp, Cel)
                  Prd.Prd_Ext_XWing_Cellule = Cel
                  XWing = True
Xwing_Boucle_End:
                End If
                If XWing Then Exit For
              End If
            Next r
          End If
        Next c
      End If
    Next i

    'Documentation de Prd
    With Prd
      If XWing Then
        .Prd_Code_Retour = 0
        .Prd_Ext_XWing = "X"
        For i As Integer = 0 To 80
          .Prd_Ini(i) = U_temp(i, 1)
          .Prd_Val(i) = U_temp(i, 2)
          .Prd_Candidats(i) = U_temp(i, 3)
        Next i
      End If
    End With
  End Sub

  '-------------------------------------------------------------------------------------------
  '05/02/2024
  'Pzzl_Slv est utilisé pour résoudre interactivement un SuDoKu Pzzl_Slv_Interactif
  '                     pour produire interactivement un SuDoKu Pzzl_Prd_Interactif  
  '                     pour créer par lot des SuDoKus          Pzzl_Prd_Batch 
  'Egalement
  'Pzzl_Slv est utilisé pour Résoudre une Cellule
  '                     pour Suggérer une Cellule
  '                     en Mode Suggestion
  '-------------------------------------------------------------------------------------------

  Public Sub Cell_Slv_Interactif(ByVal Production_Type As String, Cellules_Type As String)
    'Cellules_Type peut prendre les valeurs suivantes : Résoudre une Cellule
    '                                                   Suggérer une Cellule
    '                                                   Mode Suggestion
    ' Production_Type est positionné "S"
    Dim Prd As Prd_Struct = Nothing
    Prd_Init(Prd, U, "I")
    Dim Strategy_Rslt(99, 0) As String
    Pzzl_Slv(Production_Type, "*One", Prd, Strategy_Rslt)
    Cell_Slv_Result(Cellules_Type, Strategy_Rslt)
  End Sub

  Public Sub Cell_Slv_Result(Cellules_Type As String, Strategy_Rslt(,) As String)
    'La fonction propose une solution au hasard des solutions de Strategy_Rslt
    Dim Plage As String = ""
    Dim Virgule As Integer = 0
    'For i As Integer = 0 To 80 : U_Suggest(i) = "0" : Next i
    If Strategy_Rslt.GetUpperBound(1) = 0 Then
      Strategy_Rslt_Display(Strategy_Rslt, -1)
      Exit Sub
    End If
    Dim n As Integer = Rd8.Next(1, UBound(Strategy_Rslt, 2) + 1) 'la plage des valeurs de retour inclut minValue mais pas maxValue.

    Dim Stratégie As String = Strategy_Rslt(1, 0)
    Dim Strategy_Index As Integer = Stg_List_Code.IndexOf(Stratégie)
    Dim Stratégie_Explicite As String = Stg_Get(Strategy_Rslt(1, 0)).Texte
    Dim Cellule As Integer = CInt(Strategy_Rslt(10, n))
    Dim Valeur As String = Strategy_Rslt(5, n)
    For k As Integer = 10 To 54
      If Strategy_Rslt(k, n) = "__" Then Exit For
      If Strategy_Rslt(k, n) <> "__" Then
        'U_Suggest(CInt(Strategy_Rslt(k, n))) = Strategy_Rslt(1, 0) 'La valeur "0" est la valeur négative
        Virgule += 1
        If Virgule > 1 Then Plage &= ", "
        Plage &= U_Coord(CInt(Strategy_Rslt(k, n)))
      End If
    Next k
    Select Case Cellules_Type
      Case "Résoudre une Cellule"
        If Strategy_Index < 2 Then 'CdU et CdO
          Cell_Val_Insert(Valeur, Cellule, Msg_Read("SDK_40202"))
          Frm_SDK.B_Info.Text = "Résolution de la Cellule. " & Stratégie_Explicite & " en " & Plage
        Else
          Frm_SDK.B_Info.Text = "Résolution de la Cellule. " & Stratégie_Explicite & " en " & Plage
        End If
        'Case "Suggérer une Cellule"
        '  If Strategy_Index < 2 Then 'CdU et CdO
        '    Frm_SDK.B_Info.Text = "Suggestion de la Cellule. " & Stratégie_Explicite & " en " & Plage
        '  End If
        'Case "Mode Suggestion"
        '  Frm_SDK.B_Info.Text = "Mode Suggestion. " & Stratégie_Explicite & " en " & Plage
    End Select
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub


#Region "Production Batch de SuDoKus"
  '------------------------------------------------------------------------------------------
  'Date de création: Samedi 15/10/2022
  'Ce module regroupe les traitements Batch de création de Sudokus
  'Le terme batch signifie que la création de grilles est effectué dans un thread en arrière-plan
  '                     et que le nombre de grilles créées peut être conséquent.
  '               signifie que les SuDoKus sont créés en lot.
  '
  'Batch_Thread est arrêté dans Frm_SDK_FormClosing s'il est actif 
  'La ressource Journal n'est pas accessible par ce thread.
  'Parce que la création d'une grille de SuDoKu est un process long et non garanti d'un résultat,
  'SDK dispose d'un "stock" de grille. Ce stock est créé en batch en arrière-plan.
  'Un icône est affiché dans le menu Extension pendant le traitement,
  '         est enlevé à la fin du traitement
  'Exécution:
  '  Frm_SDK_Load (dernière instruction)
  '  --> Batch_Initial()
  '      Si les conditions sont remplies avec Plcy_Generate_Batch
  '------------------------------------------------------------------------------------------


  Public Sub Pzzl_Prd_Batch(ByVal Production_Type As String, ByRef Prd As Prd_Struct)
    'Production_Type = "P"
    'Création de grilles de Sudoku en batch
    'La création est faite en arrière-plan, donc en mode silencieux
    Dim Tentative As Integer = 0
    Try

Pzzl_Prd_Batch_Boucle:
      Tentative += 1
      If Tentative > CInt(Create_Nb_Tentatives) Then GoTo Pzzl_Prd_Batch_Exit
      '   Prd est Remis à zéro
      '   Prd garde la même contrainte pour tout le traitement en arrière-plan
      With Prd
        ReDim .Crt_Strg_Nb(10)
        ReDim .Slv_Strg_Nb(10)
        ReDim .Prd_Ini(80)
        ReDim .Prd_Val(80)
        ReDim .Prd_Candidats(80)
        .Prd_Code_Retour = -2
      End With

      '   Crt : Construction d'une grille
      Pzzl_Crt(Production_Type, Prd)
      If Prd.Prd_Code_Retour <> 0 Then GoTo Pzzl_Prd_Batch_Boucle
      '  Prd.Prd_Code-Retour = 0, alors on tente de résoudre la grille
      Prd.Prd_Code_Retour = -2
      Dim Strategy_Rslt(99, 0) As String
      Pzzl_Slv(Production_Type, "*All", Prd, Strategy_Rslt)

      If Prd.Prd_Code_Retour = 1 Then
        Pzzl_Crt_Triplet(Prd)
        Pzzl_Crt_XWing(Prd)
        Prd.Prd_Code_Retour = -2
        Pzzl_Slv(Production_Type, "*All", Prd, Strategy_Rslt)
      End If

      If Prd.Prd_Code_Retour = 0 Then
        ' Il ne reste plus de cellules vides, Il FAUDRA que DLCode soit Dlu
        Dim U_temp(,) As String = Prd_To_U_temp(Prd)
        Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
        Prd.Prd_DlCode = DL.DLCode
        Prd.Prd_DlSolution = DL.Solution(0)
      End If

      If Prd.Prd_Code_Retour = 1 Then
        ' Il reste des cellules vides (Si DLCode = Dlu, alors les stratégies sont insuffisantes)
        Dim U_temp(,) As String = Prd_To_U_temp(Prd)
        Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
        Prd.Prd_DlCode = DL.DLCode
        Prd.Prd_DlSolution = DL.Solution(0)
        If DL.DLCode = "Dlu" Then Prd.Prd_Code_Retour = 0
      End If

      If Prd.Prd_Code_Retour <> 0 Then GoTo Pzzl_Prd_Batch_Boucle

Pzzl_Prd_Batch_End:
      If Prd.Prd_Code_Retour <> 0 Then GoTo Pzzl_Prd_Batch_Exit
      ' Refus d'une grille ayant un CdU
      ' Cette règle a pour conséquence une raréfaction des grilles faciles
      Dim U_tempa(80, 3) As String
      Dim Cdu_Exists As Boolean
      For i As Integer = 0 To 80
        U_tempa(i, 2) = Prd.Prd_Ini(i)
        '  Select Case U_tempa(i, 2)
        If U_tempa(i, 2) = " " Then U_tempa(i, 3) = Cnddts
        If U_tempa(i, 2) <> " " Then U_tempa(i, 3) = Cnddts_Blancs
      Next i
      Grid_Cdd_Remove_Cell_Coll(U_tempa)
      For i As Integer = 0 To 80
        If Trim(U_tempa(i, 3)).Length = 1 Then
          Cdu_Exists = True
          Exit For
        End If
      Next i

      If Create_Grille_CdU Then
      Else
        If Cdu_Exists Then GoTo Pzzl_Prd_Batch_Exit
      End If

      'Enregistrement du Puzzle
      Dim File As String = Pzzl_Save(Prd)
    Catch ex As Exception
      Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      MsgBox(ex.ToString(),, MsgTit)
    End Try
Pzzl_Prd_Batch_Exit:
  End Sub


#End Region

  Public Function Cell_Coll_Val_Check(U_temp(,) As String, Valeur As String, Cellule As Integer) As Integer
    'La fonction retourne -1 si la valeur dans cette Cellule n'existe pas dans 20 Cellules Collatérales
    '            retourne la première cellule en anomalie (valeur 0 à 80)
    Cell_Coll_Val_Check = -1
    Dim Grp() As Integer = U_20Cell_Coll(Cellule)
    For g As Integer = 0 To UBound(Grp)
      Dim Cell_Coll As Integer = Grp(g)
      If U_temp(Cell_Coll, 2) = " " Then Continue For
      If U_temp(Cell_Coll, 2) = Valeur Then
        Cell_Coll_Val_Check = Cell_Coll
        Exit For
      End If
    Next g
    Return Cell_Coll_Val_Check
  End Function

End Module