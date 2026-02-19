Option Strict On
Option Explicit On

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
' Stratégie Indicative
'           Placement d'une valeur
'           Eviction d'un Candidat  
' 
'---------------------------------------------------------------------------------------------------------------------------------

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Friend Module P01_Strategy

  Sub Strategy_Compteur_Display(Type As String, Nb As Integer())
    'Compteurs des Stratégies:  (Profondeur : 9)
    '0-CdU 1-CdO 2-Cbl 3-Tpl 4-Xwg 5-XYw 6-Swf 7-Jly 8-XYZ 9-SKy10-Unq 
    '  101    16    10     0     0     0     0     0     0     0 
    Jrn_Add("SDK_Space")
    Jrn_Add(, {"Compteurs des Stratégies"})
    Jrn_Add(, {Type})
    Dim S1 As String = "  "
    Dim S2 As String = "  "
    For i As Integer = 0 To 10
      S1 &= CStr(i).PadLeft(1) & "-" & Stg_List_Code(i) & " "
      S2 &= CStr(Nb(i)).PadLeft(5) & " "
    Next i
    Jrn_Add(, {S1})
    Jrn_Add(, {S2})
  End Sub

  '
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  '
  Sub Strategy_Dsp_Standard()
    'Procédure appelée par: Mnu04n_AnnulerLaDerniereOption_Click

    Strategy_Switch("   ")
    ' Plcy_Strg est positionnée à "   "  
    Dim item_DCd As ToolStripMenuItem = DirectCast(Frm_SDK.Mnu04.DropDown.Items("Mnu04n_DCd"), ToolStripMenuItem)
    item_DCd.Checked = False
    Dim item_CdS As ToolStripMenuItem = DirectCast(Frm_SDK.Mnu04.DropDown.Items("Mnu04n_CdS"), ToolStripMenuItem)
    item_CdS.Checked = False

    Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Event_OnPaint_MAP = Proc_Name_Get() & " Plcy_Strg: '" & Plcy_Strg & "'"
    Event_OnPaint = "Total"
    Frm_SDK.Invalidate()
  End Sub

  ' Procédure générique


  'Sub Dsp_AideGraphique(Dsp As String)
  '  'Afficher / Ne pas afficher l'AideGraphique
  '  'Dsp permet de forcer l'affichage de l'AideGraphique ou non, ou bien de l'alterner
  '  'Les options Aide Graphique et F1 Aide ne sont pas enregistrées dans SDK.ini
  '  Select Case Dsp
  '    Case "Non" : Plcy_AideGraphique = False
  '    Case "Alt" : If Plcy_AideGraphique Then Plcy_AideGraphique = False Else Plcy_AideGraphique = True
  '    Case "Oui" : Plcy_AideGraphique = True
  '    Case "PdC" ' Pas de changement
  '  End Select
  '  If Plcy_AideGraphique Then Frm_SDK.Mnu05_AideSudokuGraphique.Checked = True
  '  If Not Plcy_AideGraphique Then Frm_SDK.Mnu05_AideSudokuGraphique.Checked = False
  'End Sub
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

    Event_OnPaint = "Total"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub
  Sub Strategy_Switch(Strg As String)
    ' Switch la Plcy_Strg_Swt entre 1 et -1
    ' La difficulté réside sur le fait que l'on peut passer d'une stratégie CdU-On à une stratégie CdO-On
    '    sans être passé par la stratégie CdU-Off
    ' Permet d'alterner la stratégie On/Off/On/Off Etc
    ' Strategy_Dsp_Standard() positionne en Off la stratégie précédente

    Dim AllStrategies As New HashSet(Of String)(
    {"   ",
     "Cdd", "CdU", "CdO", "DCd", "CdS", "Cbl", "Tpl", "Xwg", "XYw", "Swf", "Jly", "XYZ", "SKy", "Unq",
     "FV1", "FV2", "FV3", "FV4", "FV5", "FV6", "FV7", "FV8", "FV9",
     "FC1", "FC2", "FC3", "FC4", "FC5", "FC6", "FC7", "FC8", "FC9"})

    If Not AllStrategies.Contains(Strg) Then
      Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
      Jrn_Add("ERR_00140", {"#" & Strg & "#"}, "Erreur")
      Exit Sub
    End If

    Plcy_Strg = Strg
    Plcy_Strg_Swt = If(Plcy_Strg <> Prv_Plcy_Strg, 1, -Plcy_Strg_Swt)
    Prv_Plcy_Strg = Plcy_Strg
  End Sub

  Function Strategy_Click(Cellule As Integer, ByRef Strategy_Rslt(,) As String) As Integer
    'Strategy_Click retourne le numéro de ligne de Strategy_Rslt
    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      For k As Integer = 10 To 54
        If Strategy_Rslt(k, i) = "__" Then Exit For
        If CInt(Strategy_Rslt(k, i)) = Cellule Then Return i
      Next k
    Next i
    Return -1
  End Function

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

  Sub OO_130_Stg_Init()
    'Initialisation de la Liste des Stratégies
    ' 12/10/2025 Toutes les Plcy_Strg sont dans la liste
    '                               Lettre pour la barre d'Outils ou N
    Stg_List.Add(New Stg_Cls("   ", "N", "N", "N", "N", "0", "Aucune Stratégie"))
    Stg_List.Add(New Stg_Cls("Cdd", "C", "O", "N", "N", "1", "Stratégie des Candidats"))
    '                                     O Pour afficher la lettre dans le bouton et le Tooltiptext
    '                                         Pour afficher Insérer ou Exclure avec les candidats des stratégies
    Stg_List.Add(New Stg_Cls("CdU", "U", "O", "I", "O", "3", "Stratégie des Candidats Uniques"))
    Stg_List.Add(New Stg_Cls("CdO", "O", "O", "I", "O", "3", "Stratégie des Candidats Obligatoires"))
    Stg_List.Add(New Stg_Cls("DCd", "N", "N", "I", "N", "4", "Stratégie des Derniers Candidats"))
    Stg_List.Add(New Stg_Cls("CdS", "N", "N", "I", "N", "4", "Stratégie du Candidat Saisi"))
    Stg_List.Add(New Stg_Cls("Flt", "N", "N", "N", "N", "3", "Stratégie des Filtres"))
    Stg_List.Add(New Stg_Cls("FV1", "N", "N", "N", "N", "2", "Filtre des Valeurs 1"))
    Stg_List.Add(New Stg_Cls("FV2", "N", "N", "N", "N", "2", "Filtre des Valeurs 2"))
    Stg_List.Add(New Stg_Cls("FV3", "N", "N", "N", "N", "2", "Filtre des Valeurs 3"))
    Stg_List.Add(New Stg_Cls("FV4", "N", "N", "N", "N", "2", "Filtre des Valeurs 4"))
    Stg_List.Add(New Stg_Cls("FV5", "N", "N", "N", "N", "2", "Filtre des Valeurs 5"))
    Stg_List.Add(New Stg_Cls("FV6", "N", "N", "N", "N", "2", "Filtre des Valeurs 6"))
    Stg_List.Add(New Stg_Cls("FV7", "N", "N", "N", "N", "2", "Filtre des Valeurs 7"))
    Stg_List.Add(New Stg_Cls("FV8", "N", "N", "N", "N", "2", "Filtre des Valeurs 8"))
    Stg_List.Add(New Stg_Cls("FV9", "N", "N", "N", "N", "2", "Filtre des Valeurs 9"))
    Stg_List.Add(New Stg_Cls("FC1", "N", "N", "N", "N", "3", "Filtre des Candidats 1"))
    Stg_List.Add(New Stg_Cls("FC2", "N", "N", "N", "N", "3", "Filtre des Candidats 2"))
    Stg_List.Add(New Stg_Cls("FC3", "N", "N", "N", "N", "3", "Filtre des Candidats 3"))
    Stg_List.Add(New Stg_Cls("FC4", "N", "N", "N", "N", "3", "Filtre des Candidats 4"))
    Stg_List.Add(New Stg_Cls("FC5", "N", "N", "N", "N", "3", "Filtre des Candidats 5"))
    Stg_List.Add(New Stg_Cls("FC6", "N", "N", "N", "N", "3", "Filtre des Candidats 6"))
    Stg_List.Add(New Stg_Cls("FC7", "N", "N", "N", "N", "3", "Filtre des Candidats 7"))
    Stg_List.Add(New Stg_Cls("FC8", "N", "N", "N", "N", "3", "Filtre des Candidats 8"))
    Stg_List.Add(New Stg_Cls("FC9", "N", "N", "N", "N", "3", "Filtre des Candidats 9"))
    '                                N pas de lettre
    Stg_List.Add(New Stg_Cls("Obj", "N", "N", "N", "N", "4", "Dessiner sur la Grille"))
    Stg_List.Add(New Stg_Cls("Edi", "N", "N", "N", "N", "4", "Edition de la Grille"))
    Stg_List.Add(New Stg_Cls("Cbl", "B", "O", "E", "O", "3", "Stratégie des Candidats bloqués"))
    Stg_List.Add(New Stg_Cls("Tpl", "T", "O", "E", "O", "3", "Stratégie des Candidats doubles, triples, quadruples"))
    Stg_List.Add(New Stg_Cls("Xwg", "X", "O", "E", "O", "3", "Stratégie X-Wing, Finned, Sashimi"))
    Stg_List.Add(New Stg_Cls("XYw", "Y", "O", "E", "O", "3", "Stratégie XY-Wing"))
    Stg_List.Add(New Stg_Cls("Swf", "S", "O", "E", "O", "3", "Stratégie Swordfish"))
    Stg_List.Add(New Stg_Cls("Jly", "J", "O", "E", "O", "3", "Stratégie Jellyfish"))
    Stg_List.Add(New Stg_Cls("XYZ", "Z", "O", "E", "O", "3", "Stratégie XYZ-Wing"))
    Stg_List.Add(New Stg_Cls("SKy", "K", "O", "E", "O", "3", "Stratégie SKyscraper, Kyte, Empty Rectangle"))
    Stg_List.Add(New Stg_Cls("Unq", "Q", "O", "E", "O", "3", "Stratégie Uniqueness"))

    ' Les codes des nouvelles stratégies des Liens commencent par X ou W
    '                                la lettre est L Link
    Stg_List.Add(New Stg_Cls("GLk", "N", "N", "E", "N", "4", "Affichage des Liens forts"))
    Stg_List.Add(New Stg_Cls("Gbl", "L", "N", "E", "N", "3", "Stratégie bi-locaux (Candidats bloqués)"))
    Stg_List.Add(New Stg_Cls("Gbv", "L", "N", "E", "N", "3", "Stratégie bi-values"))
    Stg_List.Add(New Stg_Cls("GCs", "L", "N", "E", "N", "3", "Stratégie DFS Coloration Simple"))
    Stg_List.Add(New Stg_Cls("XCx", "L", "N", "E", "N", "3", "Stratégie X-Chain"))
    Stg_List.Add(New Stg_Cls("XCy", "L", "N", "E", "N", "3", "Stratégie XY-Chain"))
    Stg_List.Add(New Stg_Cls("XRp", "L", "N", "E", "N", "3", "Stratégie Remote Pairs"))
    Stg_List.Add(New Stg_Cls("XNl", "L", "N", "E", "N", "3", "Stratégie Nice_Loop"))
    Stg_List.Add(New Stg_Cls("WgX", "L", "N", "E", "N", "3", "Stratégie X-Wing"))
    Stg_List.Add(New Stg_Cls("WgY", "L", "N", "E", "N", "3", "Stratégie XY-Wing"))
    Stg_List.Add(New Stg_Cls("WgZ", "L", "N", "E", "N", "3", "Stratégie XYZ-Wing"))
    Stg_List.Add(New Stg_Cls("WgW", "L", "N", "E", "N", "3", "Stratégie W-Wing"))

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
    Return New Stg_Cls(Code, "#", "#", "#", "#", "#", "#")
  End Function


End Module