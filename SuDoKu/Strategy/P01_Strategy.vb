'---------------------------------------------------------------------------------------------------------------------------------
' Strategy_Rslt  comporte les informations des calculs des différentes stratégies
'          il est suffisamment large pour contenir les informations de la stratégie
'                                                  45 cellules concernées    
'                                                  45 cellules exclues
'     Str Stratégie (et string!)
'---------------------------------------------------------------------------------------------------------------------------------
'Strategy_Rslt est un tableau de (99, x) comportant les résultats des stratégies en STRING
'      il est "local" à chaque traitement  
'Le  poste   Strategy_Rslt(0, 0)      comporte le nombre d'enregistrements
'Le  poste   Strategy_Rslt(1, 0)      comporte la stratégie  
'Le  poste   Strategy_Rslt(2, 0)      comporte la stratégie
'Le  poste   Strategy_Rslt(3, 0)         
'Lors de l'initialisation on a:
'            Strategy_Rslt(0, 0)      = "-1"
'            Strategy_Rslt(1, 0)      = la stratégie  
'            Strategy_Rslt(2, 0)      = la stratégie
'            Strategy_Rslt(3, 0)      = "L"

'Les postes  Strategy_Rslt(0, x)      comporte le numéro d'enregistrement, on traite For i = 1 To UBound(Strategy_Rslt, 2)
'Les postes  Strategy_Rslt(1, x)      comporte la stratégie
'Les postes  Strategy_Rslt(2, x)      comporte la sous-stratégie détaillée sur 5 caractères
'Les postes  Strategy_Rslt(3, x)      comporte le Code_LCR, soit "L", "C" ou "R"
'Les postes  Strategy_Rslt(4, x)      comporte le LCR, soit de 0 à 8 le numéro de Ligne, de Colonne ou de Région
'Les postes  Strategy_Rslt(5, x)      comporte Le Candidat concerné par la stratégie
'            Strategy_Rslt(5,         utilisé 27 fois
'                                     NON UTILISéS 6, 7, 8, 9
'Les postes  Strategy_Rslt(6, x)      comporte les Candidats sous la forme 123456789 utilisation Stratégie Unq ?
'Les postes  Strategy_Rslt(7, x)       
'Les postes  Strategy_Rslt(8, x)       
'Les postes  Strategy_Rslt(9, x)       

'Les postes  Strategy_Rslt(10 à 54,x) comporte Les cellules concernées par cette stratégie ou "__"
'Les postes  Strategy_Rslt(55 à 99,x) comporte Les cellules concernées pas l' EXCLUSION du candidat de cette stratégie
' Strategy_Rslt_Add est utilisé 30 fois dans les stratégies

' If Strategy_Rslt.GetUpperBound(1) = 0 signifie que le tableau est créé par Dim ou Public (100, 0)
' et qu'il est vide
'---------------------------------------------------------------------------------------------------------------------------------
'
' Normalement les 45 cellules concernées et en Exclusion sont remplies de gauche à droite sans trou,
'             les boucles de 10 to 54 et 55 to 99 peuvent donc s'arrêter à la première valeur "__"
'
'---------------------------------------------------------------------------------------------------------------------------------

Friend Module P01_Strategy
  '
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  '
  Sub Strategy_Dsp_Standard()
    Plcy_Strg = "   "
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Frm_SDK.Invalidate()
  End Sub

  Sub Strategy_Code(strg_Code As String, Origine As String)
    Plcy_Strg = strg_Code
    Frm_SDK.B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()
    If Plcy_Strg = "   " Then
      Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
      Exit Sub
    End If

    If Plcy_Strg <> "   " Then Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte

    Dim U_temp(80, 3) As String
    Dim Strategy_Rslt(,) As String = Nothing
    Array.Copy(U, U_temp, UNbCopy)
    Select Case Plcy_Strg
      Case "Cdd"
      Case "CdU" : Strategy_Rslt = Strategy_CdU(U_temp)
      Case "CdO" : Strategy_Rslt = Strategy_CdO(U_temp)
      Case "FC1", "FC2", "FC3", "FC4", "FC5", "FC6", "FC7", "FC8", "FC9"
      Case "FV1", "FV2", "FV3", "FV4", "FV5", "FV6", "FV7", "FV8", "FV9"
      Case "Cbl" : Strategy_Rslt = Strategy_Cbl(U_temp)
      Case "Tpl" : Strategy_Rslt = Strategy_Tpl(U_temp)
      Case "Xwg" : Strategy_Rslt = Strategy_Xwg(U_temp)
      Case "XYw" : Strategy_Rslt = Strategy_XYw(U_temp)
      Case "Swf" : Strategy_Rslt = Strategy_Swf(U_temp)
      Case "Jly" : Strategy_Rslt = Strategy_Jly(U_temp)
      Case "XYZ" : Strategy_Rslt = Strategy_XYZ(U_temp)
      Case "SKy" : Strategy_Rslt = Strategy_SKy(U_temp)
      Case "Unq" : Strategy_Rslt = Strategy_Unq(U_temp)
      Case Else
    End Select

    Dim index As Integer = RRslt_Copy_Rnd(Strategy_Rslt)
    Frm_SDK.Invalidate()
  End Sub

  Sub Undo_Redo(Action As String)
    Dim UR_Sym As String = "<>"
    Game_Undo_Redo = Action
    Select Case Action
      Case "Annuler" 'Vert
        UR_Nb_Refaire = 0
        'Dès UR_A_Index est arrivé à Zéro, Annuler s'arrête
        If UR_Nb_Annuler > 0 And UR_A_Index = 0 Then
          Frm_SDK.Mnu02_Annuler.Enabled = False
          UR_Nb_Annuler = 0
          Exit Sub
        End If
        If UR_Nb_Annuler = 0 Then 'Si c'est la première fois
          UR_A_Index = Act_Index
        Else
          UR_A_Index -= 1
        End If
        UR_Nb_Annuler += 1
        'Dès qu'un Effacer a été fait, il faut pouvoir refaire
        If UR_Nb_Annuler > 0 Then
          Frm_SDK.Mnu02_Refaire.Enabled = True
        End If
        For i As Integer = 0 To 80
          U(i, 2) = Act(8, UR_A_Index).Substring(i, 1)
          U(i, 3) = Act(9, UR_A_Index).Substring(i * 9, 9)
        Next i
        UR_Sym = "<" & StrDup(7 - Act(1, UR_A_Index).Length, "-")
      Case "Refaire" 'Bleu
        UR_Nb_Annuler = 0
        'Refaire doit s'arrêter dès que UR_A_Index est arrivé à Act_Index
        If UR_Nb_Refaire > 0 And UR_A_Index > Act_Index Then
          Frm_SDK.Mnu02_Refaire.Enabled = False
          UR_Nb_Refaire = 0
          Exit Sub
        End If
        If UR_Nb_Refaire = 0 Then 'Si c'est la première fois
          UR_A_Index += 0
        Else
          UR_A_Index += 1
        End If
        If UR_A_Index > Act_Index Then Exit Sub
        UR_Nb_Refaire += 1
        For i As Integer = 0 To 80
          U(i, 2) = Act(10, UR_A_Index).Substring(i, 1)
          U(i, 3) = Act(11, UR_A_Index).Substring(i * 9, 9)
        Next i
        UR_Sym = StrDup(7 - Act(1, UR_A_Index).Length, "-") & ">"
    End Select
    Dim S As String = "| " & Act(4, UR_A_Index) &
                      " | " & Act(5, UR_A_Index) &
                      " | " & Act(6, UR_A_Index) &
                      " | " & UR_Sym & " " & Act(1, UR_A_Index) &
                      " | " & Act(2, UR_A_Index).PadRight(15).Substring(0, 14) &
                      " | " & Action
    Jrn_Add(, {S}, Action)

    Frm_SDK.Invalidate()
  End Sub

#Region "RRslt: Les Résultats d'une stratégie 'Classique'"
  Public Sub RRslt_Display()
    With RRslt
      Jrn_Add(, {"Affichage des données RRslt _ Occurence " & .Occurence & " / " & .Nb_Occurences})
      Jrn_Add(, {"Code Stratégie " & .Code_Strg & ", " & Stg_Get(.Code_Strg).Texte})
      Jrn_Add(, {"Sous_Stratégie " & .Code_Sous_Strg})
      Jrn_Add(, {"Code_LCR_LCR   " & .Code_LCR & (.LCR + 1)})
      Jrn_Add(, {"Candidat       " & .Candidat})
      If .Cellule IsNot Nothing Then Jrn_Add(, {"Cellule        " & String.Join(", ", .Cellule.Select(Function(c) U_Coord(c)))})
      If .CelExcl IsNot Nothing Then Jrn_Add(, {"CelExcl        " & String.Join(", ", .CelExcl.Select(Function(c) U_Coord(c)))})
      Jrn_Add(, {"Productivité   " & .Productivité.ToString()})
    End With
  End Sub

  Public Function RRslt_Copy_Rnd(Strategy_Rslt(,) As String) As Integer
    ' La procédure documente RRSlt d'un résultat pris au hasard de Strategy_Rslt
    If Strategy_Rslt Is Nothing OrElse UBound(Strategy_Rslt, 2) <= 0 Then     'Strategy_Rslt.GetLength(1) = 0
      RRslt.Productivité = False
      Return -1
    End If

    Dim rnd As New Random()
    Dim Index As Integer = rnd.Next(1, Strategy_Rslt.GetUpperBound(1) + 1)   ' Tire un nombre entre min inclus et max non inclus
    With RRslt
      .Occurence = Index
      .Nb_Occurences = Strategy_Rslt.GetUpperBound(1)
      .Code_Strg = Strategy_Rslt(1, Index)
      .Code_Sous_Strg = Strategy_Rslt(2, Index)
      .Code_LCR = Strategy_Rslt(3, Index)
      .LCR = CInt(Strategy_Rslt(4, Index))
      .Candidat = Strategy_Rslt(5, Index)

      Dim valeurs As New List(Of Integer)
      For k As Integer = 0 To 44
        Dim s As String = Strategy_Rslt(10 + k, Index)
        If s = "__" Then Exit For
        valeurs.Add(CInt(s))
      Next
      .Cellule = valeurs.ToArray()
      valeurs.Clear()
      For k As Integer = 0 To 44
        Dim s As String = Strategy_Rslt(55 + k, Index)
        If s = "__" Then Exit For
        valeurs.Add(CInt(s))
      Next
      .CelExcl = valeurs.ToArray()

      .Productivité = True
    End With
    'Strategy_Rslt_Display(Strategy_Rslt, -1)
    'RRslt_Display()
    Return Index

  End Function

  Public Sub RRslt_Control_Cdd_Exclure(Candidat As String)
    'Jrn_Add_Red(Proc_Name_Get())
    'Jrn_Add_Red("Liste des cellules du candidat " & Candidat & " à enlever")
    Dim S As String
    For i As Integer = 0 To 80
      If U_Strg_Cdd_Exc(i) <> Cnddts_Blancs Then
        S = (U_Coord(i) & " " & U(i, 3))
        If XSolution(i) = Candidat Then
          Jrn_Add(, {S & "❗ La cellule doit prendre la valeur " & Candidat & " !"}, "Erreur")
        End If
      End If
    Next
  End Sub

#End Region

  Sub Strategy_Rslt_Init(ByRef Strategy_Rslt(,) As String,
                         Stratégie As String,
                         Sous_Stratégie As String)
    ' Initialisation, crée l'enregistrement initial 
    Dim n As Integer = 0
    ReDim Strategy_Rslt(99, 0)                                  ' Effacement du tableau et Ajout de la première ligne  
    Strategy_Rslt(0, n) = "-1"                                  ' Numérotation de 0 à n                 
    Strategy_Rslt(1, n) = Stratégie
    Strategy_Rslt(2, n) = Sous_Stratégie
    Strategy_Rslt(3, n) = "L"
    Strategy_Rslt(4, n) = "_"
    Strategy_Rslt(5, n) = "_"
    Strategy_Rslt(6, n) = "_"
    Strategy_Rslt(7, n) = "_"
    Strategy_Rslt(8, n) = "_"
    Strategy_Rslt(9, n) = "_"
    For k As Integer = 0 To 44 : Strategy_Rslt(10 + k, n) = "__" : Next k  ' Cellules Concernées par la stratégie
    For k As Integer = 0 To 44 : Strategy_Rslt(55 + k, n) = "__" : Next k  ' Cellules Concernées par l’exclusion
  End Sub
  Sub Strategy_Rslt_Add(ByRef Strategy_Rslt(,) As String,
                        Stratégie As String,
                        Sous_Stratégie As String,
                        Code_LCR As String,
                        LCR As String,
                        Candidat As String,
                        Cel45() As String,
                        Exc45() As String)
    Dim n As Integer = UBound(Strategy_Rslt, 2) + 1
    ReDim Preserve Strategy_Rslt(99, n)                                  ' Ajout de la ligne  
    Strategy_Rslt(0, n) = CStr(n)                                        ' Numérotation de 0 à n                 
    Strategy_Rslt(1, n) = Stratégie                                      ' 3 caractères CdU, CdO, Cbl, Tpl, Xwg, XYw, Swf, Jly, XYZ, SKy, Unq
    Strategy_Rslt(2, n) = Sous_Stratégie                                 ' Max 5 caractères
    Strategy_Rslt(3, n) = Code_LCR
    Strategy_Rslt(4, n) = LCR
    Strategy_Rslt(5, n) = Candidat
    Strategy_Rslt(6, n) = "_"
    Strategy_Rslt(7, n) = "_"
    Strategy_Rslt(8, n) = "_"
    Strategy_Rslt(9, n) = "_"
    For k As Integer = 0 To 44 : Strategy_Rslt(10 + k, n) = Cel45(k) : Next k       ' Cellules Concernées par la stratégie de 10 à 54
    For k As Integer = 0 To 44 : Strategy_Rslt(55 + k, n) = Exc45(k) : Next k       ' Cellules Concernées par l’exclusion  de 55 à 100
    ' le Nombre d'enregistrements est stockées dans (0, 0) 
    ' Lors de l'initialisation, cette valeur   vaut -1
    Strategy_Rslt(0, 0) = CStr(n)
  End Sub
  Public Sub Strategy_Rslt_Display(ByRef Strategy_Rslt(,) As String, Ligne As Integer)
    'Cette routine liste une seule ligne de Strategy_Rslt  ou toutes les lignes si ligne = -1
    If Strategy_Rslt Is Nothing Then Exit Sub
    Dim Strategy_Name As String = "Strategy_Rslt"
    Jrn_Add(, {Stg_Get(Strategy_Rslt(1, 0)).Texte})
    Jrn_Add("Prl_00070", {Strategy_Name})
    Jrn_Add("Prl_00000", {"Profondeur         : " & Stg_Profondeur})
    Jrn_Add("Prl_00000", {"Ligne 0            : " & Strategy_Rslt(0, 0) & " Str: " & Strategy_Rslt(1, 0)})
    Jrn_Add("Prl_00000", {"Nombre de postes   : " & Strategy_Rslt(0, 0) & " UBound(Strategy_Rslt, 2): " & CStr(UBound(Strategy_Rslt, 2))})
    Dim Dsp_Ligne As Boolean ' Toujours initialisé à False
    Dim S As String
    Dim p As Integer
    Try
      'le poste 0 est initialisé dans Strategy_Rslt_Init()
      For i As Integer = 1 To UBound(Strategy_Rslt, 2)
        If Ligne = -1 Or i = Ligne Then
          Dim LCR As Integer = CInt(Strategy_Rslt(4, i)) + 1
          Jrn_Add("Prl_00000", {"Ligne " & Strategy_Rslt(0, i).PadRight(5) & " S/Str " & Strategy_Rslt(2, i).PadRight(5) &
                  " LCR " & Strategy_Rslt(3, i) & CStr(LCR) & "  Cdd: " & Strategy_Rslt(5, i)})
          'Lignes des cellules concernées
          For j As Integer = 0 To 4
            Dsp_Ligne = False
            For k As Integer = 0 To 8
              p = 10 + (9 * j) + k
              If Strategy_Rslt(p, i) <> "__" Then
                'Dès qu'une valeur est <> de "__" la ligne sera affichée
                Dsp_Ligne = True : Exit For
              End If
            Next k
            S = ""
            For k As Integer = 0 To 8
              p = 10 + (9 * j) + k
              If Strategy_Rslt(p, i) <> "__" Then
                S &= U_Coord(CInt(Strategy_Rslt(p, i))) & " "
              Else
                S &= "_____" & " "
              End If
            Next k
            If Dsp_Ligne Then Jrn_Add("Prl_00077", {CStr(j), S})
          Next j
          'Lignes des cellules exclues
          For j As Integer = 0 To 4
            Dsp_Ligne = False
            For k As Integer = 0 To 8
              p = 55 + (9 * j) + k
              If Strategy_Rslt(p, i) <> "__" Then
                'Dès qu'une valeur est <> de "__" la ligne sera affichée
                Dsp_Ligne = True : Exit For
              End If
            Next k
            S = ""
            For k As Integer = 0 To 8
              p = 55 + (9 * j) + k
              If Strategy_Rslt(p, i) <> "__" Then
                S &= U_Coord(CInt(Strategy_Rslt(p, i))) & " "
              Else
                S &= "_____" & " "
              End If
            Next k
            If Dsp_Ligne Then Jrn_Add("Prl_00078", {CStr(j), S})
          Next j
        End If
      Next i

    Catch ex As Exception 'Strategy_Rslt peut être vide si la stratégie n'a rien donné 
      ' Il doit y avoir l'enregistrement 0 néanmoins
      Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
      Jrn_Add("ERR_00000", {ex.Message})
      Jrn_Add("ERR_00000", {ex.ToString()})
      Jrn_Add(, {"Anomalie_la stratégie est vide"})
    End Try
  End Sub

  Sub Stg_List_Init()
    'Initialisation de la Liste des Stratégies
    ' 12/10/2025 Toutes les Plcy_Strg sont dans la liste
    '                               Lettre pour la barre d'Outils ou N
    Stg_List.Add(New Stg_Cls("   ", "N", "N", "N", "N", 0, "Aucune Stratégie"))
    Stg_List.Add(New Stg_Cls("DCs", "N", "N", "N", "N", 7, "Dernier Candidat saisi"))
    Stg_List.Add(New Stg_Cls("Sai", "N", "N", "N", "N", 0, "Saisir une grille"))
    Stg_List.Add(New Stg_Cls("Cdd", "C", "O", "N", "N", 1, "Stratégie des Candidats"))
    '                                     O Pour afficher la lettre dans le bouton et le Tooltiptext
    '                                         Pour afficher Insérer ou Exclure avec les candidats des stratégies
    Stg_List.Add(New Stg_Cls("CdU", "U", "O", "I", "O", 3, "Stratégie des Candidats Uniques"))
    Stg_List.Add(New Stg_Cls("CdO", "O", "O", "I", "O", 3, "Stratégie des Candidats Obligatoires"))
    Stg_List.Add(New Stg_Cls("Flt", "N", "N", "N", "N", 3, "Stratégie des Filtres"))
    Stg_List.Add(New Stg_Cls("FV1", "N", "N", "N", "N", 2, "Filtre des Valeurs 1"))
    Stg_List.Add(New Stg_Cls("FV2", "N", "N", "N", "N", 2, "Filtre des Valeurs 2"))
    Stg_List.Add(New Stg_Cls("FV3", "N", "N", "N", "N", 2, "Filtre des Valeurs 3"))
    Stg_List.Add(New Stg_Cls("FV4", "N", "N", "N", "N", 2, "Filtre des Valeurs 4"))
    Stg_List.Add(New Stg_Cls("FV5", "N", "N", "N", "N", 2, "Filtre des Valeurs 5"))
    Stg_List.Add(New Stg_Cls("FV6", "N", "N", "N", "N", 2, "Filtre des Valeurs 6"))
    Stg_List.Add(New Stg_Cls("FV7", "N", "N", "N", "N", 2, "Filtre des Valeurs 7"))
    Stg_List.Add(New Stg_Cls("FV8", "N", "N", "N", "N", 2, "Filtre des Valeurs 8"))
    Stg_List.Add(New Stg_Cls("FV9", "N", "N", "N", "N", 2, "Filtre des Valeurs 9"))
    Stg_List.Add(New Stg_Cls("FC1", "N", "N", "N", "N", 3, "Filtre des Candidats 1"))
    Stg_List.Add(New Stg_Cls("FC2", "N", "N", "N", "N", 3, "Filtre des Candidats 2"))
    Stg_List.Add(New Stg_Cls("FC3", "N", "N", "N", "N", 3, "Filtre des Candidats 3"))
    Stg_List.Add(New Stg_Cls("FC4", "N", "N", "N", "N", 3, "Filtre des Candidats 4"))
    Stg_List.Add(New Stg_Cls("FC5", "N", "N", "N", "N", 3, "Filtre des Candidats 5"))
    Stg_List.Add(New Stg_Cls("FC6", "N", "N", "N", "N", 3, "Filtre des Candidats 6"))
    Stg_List.Add(New Stg_Cls("FC7", "N", "N", "N", "N", 3, "Filtre des Candidats 7"))
    Stg_List.Add(New Stg_Cls("FC8", "N", "N", "N", "N", 3, "Filtre des Candidats 8"))
    Stg_List.Add(New Stg_Cls("FC9", "N", "N", "N", "N", 3, "Filtre des Candidats 9"))
    '                                N pas de lettre
    Stg_List.Add(New Stg_Cls("Obj", "N", "N", "N", "N", 9, "Dessiner sur la Grille"))
    Stg_List.Add(New Stg_Cls("Edi", "N", "N", "N", "N", 9, "Edition de la Grille"))
    Stg_List.Add(New Stg_Cls("Cbl", "B", "O", "E", "O", 3, "Stratégie des Candidats bloqués"))
    Stg_List.Add(New Stg_Cls("Tpl", "T", "O", "E", "O", 3, "Stratégie des Candidats doubles, triples, quadruples"))
    Stg_List.Add(New Stg_Cls("Xwg", "X", "O", "E", "O", 3, "Stratégie X-Wing, Finned, Sashimi"))
    Stg_List.Add(New Stg_Cls("XYw", "Y", "O", "E", "O", 3, "Stratégie XY-Wing"))
    Stg_List.Add(New Stg_Cls("Swf", "S", "O", "E", "O", 3, "Stratégie Swordfish"))
    Stg_List.Add(New Stg_Cls("Jly", "J", "O", "E", "O", 3, "Stratégie Jellyfish"))
    Stg_List.Add(New Stg_Cls("XYZ", "Z", "O", "E", "O", 3, "Stratégie XYZ-Wing"))
    Stg_List.Add(New Stg_Cls("SKy", "K", "O", "E", "O", 3, "Stratégie SKyscraper, Kyte, Empty Rectangle"))
    Stg_List.Add(New Stg_Cls("Unq", "Q", "O", "E", "O", 3, "Stratégie Uniqueness"))

    ' Les codes des nouvelles stratégies des Liens commencent par X ou W
    '                                la lettre est L Link
    Stg_List.Add(New Stg_Cls("GLk", "N", "N", "E", "N", 9, "Affichage des Liens forts"))
    Stg_List.Add(New Stg_Cls("Gbl", "L", "N", "E", "N", 4, "Stratégie bi-locaux (Candidats bloqués)"))
    Stg_List.Add(New Stg_Cls("Gbv", "L", "N", "E", "N", 4, "Stratégie bi-values"))
    Stg_List.Add(New Stg_Cls("GCs", "L", "N", "E", "N", 4, "Stratégie Coloration Simple"))
    Stg_List.Add(New Stg_Cls("GCx", "L", "N", "E", "N", 4, "Stratégie X-Chain"))
    Stg_List.Add(New Stg_Cls("XCy", "L", "N", "E", "N", 5, "Stratégie XY-Chain"))
    Stg_List.Add(New Stg_Cls("XRp", "L", "N", "E", "N", 5, "Stratégie Remote Pairs"))
    Stg_List.Add(New Stg_Cls("XNl", "L", "N", "E", "N", 5, "Stratégie Nice_Loop"))
    Stg_List.Add(New Stg_Cls("WgX", "L", "N", "E", "N", 5, "Stratégie Wing X"))
    Stg_List.Add(New Stg_Cls("WgY", "L", "N", "E", "N", 5, "Stratégie Wing XY"))
    Stg_List.Add(New Stg_Cls("WgZ", "L", "N", "E", "N", 5, "Stratégie Wing XYZ"))
    Stg_List.Add(New Stg_Cls("WgW", "L", "N", "E", "N", 5, "Stratégie Wing W"))

    'Création de Stg_List_Code, Stg_List_Lettre et Stg_List_Link
    For i As Integer = 0 To Stg_List.Count - 1
      If Stg_List.Item(i).Prd = "O" Then
        Stg_List_Code.Add(Stg_List.Item(i).Code)
        Stg_List_Lettre.Add(Stg_List.Item(i).Lettre)
      End If
      If Stg_List.Item(i).Lettre = "L" Then
        Stg_List_Link.Add(Stg_List.Item(i).Code)
      End If
    Next i
  End Sub
  Sub Stg_List_Display()
    'Liste les Stratégies et les Listes des Codes, Lettres et Liens
    Jrn_Add("SDK_Space")
    Jrn_Add(, {"Liste des Stratégies"})
    For i As Integer = 0 To Stg_List.Count - 1
      With Stg_List.Item(i)
        Jrn_Add(, {CStr(i).PadRight(3) & " " & .Code & " " & .Lettre & " " & .Dsp_BO & " " & .Type & " " & .Prd & " " & .Family & " " & .Texte})
      End With
    Next i
    Dim S As String = ""
    For i As Integer = 0 To Stg_List_Code.Count - 1
      S &= Stg_List_Code.Item(i) & " ,"
    Next i
    Jrn_Add(, {"Stg_List_Code    : " & S.Substring(0, S.Length - 2)})

    S = ""
    For i As Integer = 0 To Stg_List_Lettre.Count - 1
      S &= Stg_List_Lettre.Item(i) & " ,"
    Next i
    Jrn_Add(, {"Stg_List_Lettre  : " & S.Substring(0, S.Length - 2)})

    S = ""
    For i As Integer = 0 To Stg_List_Link.Count - 1
      S &= Stg_List_Link.Item(i) & " ,"
    Next i
    Jrn_Add(, {"Stg_List_Link    : " & S.Substring(0, S.Length - 2)})
  End Sub
  Public Function Stg_Get(Code As String) As Stg_Cls
    'Retourne les informations de la stratégie correspondant au Code
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Code = Code Then Return Stg
    Next Stg
    Return New Stg_Cls(Code, "#", "#", "#", "#", -1, "#")
  End Function
End Module