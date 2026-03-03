Option Strict On
Option Explicit On

'04/02/2025 

Friend Module M49_Prd_General
#Region "Structure Prd"
  'La structure Prd est nécessaire à la production d'une grille afin de passer des informations
  'entre les différents programmes nécessaires à la production d'un SuDoKu
  'comme entre La Création et les compléments de Création (Triplet, XWing)
  'ou    entre la création et la résolution
  ' 24/09/2024 les contraintes de création (Aucune, Symétrie, Exclusion) font l'objet de fortes modifications
  Public Structure Prd_Struct                    '  Structure d'une Production de Grille
    Public Prd_Chat As Boolean                   '  Définition d'un Mode Bavard/Silencieux
    Public Prd_Phase As String                   '  Phase résolution ou
    Public Prd_Cnt_Type As String                '  N, S, L, C, R, V (Aucune, Symétrie, Exclusion LCR, V)
    Public Prd_Cnt_Valeur As Integer             '  0, 0 à 4, 0 à 8 ou 1 à 9 
    Public Prd_Cnt_None As Boolean               '  Contrainte de Création AUCUNE
    Public Prd_Cnt_Sym As Boolean                '  Contrainte de Symétrie
    Public Prd_Cnt_Excl As Boolean               '  Contrainte d'Exclusion
    Public Prd_Create_Nb_Cel_Demandées As Integer
    Public Prd_Plcy_Strg_Bll() As Boolean        '  UOBTXYSJZKQ_0123456789
    Public Prd_Plcy_Strg_Profondeur As String
    Public Crt_Strg_Nb() As Integer
    Public Slv_Strg_Nb() As Integer
    'Prd_BI est documenté dans Prd_Init et peut comporter Batch ou Interactif, B/I
    Public Prd_BI As String
    Public Prd_Ini() As String                   '  Correspond à U(80, 1)                                     
    Public Prd_Val() As String                   '               U(80, 2)                                         
    Public Prd_Candidats() As String             '               U(80, 3)                                          
    Public Prd_Code_Retour As Integer            '
    Public Prd_Ext_Triplet As String
    Public Prd_Ext_Triplet_Cellule As Integer
    Public Prd_Ext_XWing As String
    Public Prd_Ext_XWing_Cellule As Integer
    Public Prd_DlCode As String                  '  Dl#, Dlz, Dlu ou Dl suivi du nombre de Solutions trouvées dans Dancing Link
    Public Prd_DlSolution As String              '  Dl Solution : la solution Dl
  End Structure

  Public Function Prd_To_U_temp(ByRef Prd As Prd_Struct) As String(,)
    'à/p de la structure de production, retourne un tableau U_temp(80,3) conforme Ini, Val, Cdd
    Dim U_temp(80, 3) As String
    For i As Integer = 0 To 80
      U_temp(i, 1) = Prd.Prd_Ini(i)
      U_temp(i, 2) = Prd.Prd_Val(i)
      U_temp(i, 3) = Cnddts_Blancs
      If U_temp(i, 2) = " " Then U_temp(i, 3) = Cnddts ' Soit 123456789 
    Next i
    'Documentation de U_temp: seuls les candidats éligibles sont gardés
    Grid_Cdd_Remove_Cell_Coll(U_temp)
    Return U_temp
  End Function

  Public Sub Prd_Init(ByRef Prd As Prd_Struct, U_temp(,) As String, BI As String)
    '1  Prd est initialisée
    '   Certaines valeurs inutiles sont gardées pour que Prd soit ENTIEREMENT définie
    With Prd
      .Prd_Chat = Create_Chat
      .Prd_Create_Nb_Cel_Demandées = CInt(Create_Nb_Cel_Demandées)
      ReDim .Prd_Plcy_Strg_Bll(Stg_Bll.Count - 1)
      For i As Integer = 0 To 10
        .Prd_Plcy_Strg_Bll(i) = Stg_Bll(i)
      Next i
      .Prd_Plcy_Strg_Profondeur = Stg_Profondeur
      ReDim .Crt_Strg_Nb(10)
      ReDim .Slv_Strg_Nb(10)
      .Prd_BI = BI
      ReDim .Prd_Ini(80)
      ReDim .Prd_Val(80)
      ReDim .Prd_Candidats(80)
      .Prd_Code_Retour = -2
      .Prd_Cnt_Type = "N"
      .Prd_Cnt_Valeur = 0
      .Prd_Ext_Triplet = "#"
      .Prd_Ext_Triplet_Cellule = -1
      .Prd_Ext_XWing = "#"
      .Prd_Ext_XWing_Cellule = -1
      .Prd_DlCode = ""
      .Prd_DlSolution = ""
    End With

    'Calcul des contraintes de Création comme la Symétrie ou l'Exclusion
    Dim Prd_Cnt As Integer
    '  Create_Contrainte_Originale est issue de = .Prf_02C_Create_Contrainte
    If Create_Contrainte_Originale = 3 Then
      Prd_Cnt = Rdb.Next(0, 3) 'inclu min, mais pas Max 
    Else
      Prd_Cnt = Create_Contrainte_Originale
    End If
    ' à ce niveau create contrainte ne peut être que 1, 2 ou 3
    Prd.Prd_Cnt_None = False
    Prd.Prd_Cnt_Sym = False
    Prd.Prd_Cnt_Excl = False
    Select Case Prd_Cnt
      Case 0
        Prd.Prd_Cnt_None = True
        Prd.Prd_Cnt_Type = "N"
        Prd.Prd_Cnt_Valeur = 0
      Case 1
        Prd.Prd_Cnt_Sym = True
        Dim Rd As Integer = Rd0.Next(0, 5) 'inclu min, mais pas Max 
        Prd.Prd_Cnt_Type = "S"
        Select Case Rd
          Case 0 : Prd.Prd_Cnt_Valeur = 0  ' "Diagonale Haut-Gauche Bas-Droit"
          Case 1 : Prd.Prd_Cnt_Valeur = 1  ' "Diagonale Bas-Gauche Haut-Droit"
          Case 2 : Prd.Prd_Cnt_Valeur = 2  ' "Centre L5-C5"
          Case 3 : Prd.Prd_Cnt_Valeur = 3  ' "Médiane Horizontale"
          Case 4 : Prd.Prd_Cnt_Valeur = 4  ' "Médiane Verticale"
        End Select
      Case 2
        Prd.Prd_Cnt_Excl = True
        Dim Rd As Integer = Rd1.Next(0, 4) 'inclu min, mais pas Max 
        Select Case Rd
          Case 0 : Prd.Prd_Cnt_Type = "L" : Prd.Prd_Cnt_Valeur = Rd2.Next(0, 9)
          Case 1 : Prd.Prd_Cnt_Type = "C" : Prd.Prd_Cnt_Valeur = Rd2.Next(0, 9)
          Case 2 : Prd.Prd_Cnt_Type = "R" : Prd.Prd_Cnt_Valeur = Rd2.Next(0, 9)
          Case 3 : Prd.Prd_Cnt_Type = "V" : Prd.Prd_Cnt_Valeur = Rd2.Next(1, 10)
        End Select
    End Select
    ' Prd_Cnt_Type/valeur peut prendre les valeurs :
    ' N0, S0 à S4, L0 à L4, C0 à C4, R0 à R4, V0 à V4 

    'Documentation à/p des valeurs actuelles et en cours de U
    For i As Integer = 0 To 80
      Prd.Prd_Ini(i) = U_temp(i, 1)
      Prd.Prd_Val(i) = U_temp(i, 2)
      Prd.Prd_Candidats(i) = U_temp(i, 3)
    Next i
  End Sub
  Public Sub Prd_Display(Prd As Prd_Struct)
    Jrn_Add(, {Proc_Name_Get()})
    With Prd
      Jrn_Add(, {"Mode Bavard                  : " & .Prd_Chat})
      Jrn_Add(, {"Phase                        : " & .Prd_Phase})
      Jrn_Add(, {"Plcy_Strg_Profondeur         : " & .Prd_Plcy_Strg_Profondeur})
      Jrn_Add(, {"Nb_Cellules_Demandées        : " & CStr(.Prd_Create_Nb_Cel_Demandées)})
      Jrn_Add(, {"Nombre Limite de Tentatives  : " & Create_Nb_Tentatives})
      If .Prd_Phase = "Crt" Then
        Jrn_Add(, {"Contrainte (absolue)         : " & .Prd_Cnt_Type & CStr(.Prd_Cnt_Valeur)})
      End If
      'Compteurs des Stratégies:  (Profondeur : 9)
      '0-CdU 1-CdO 2-Cbl 3-Tpl 4-Xwg 5-XYw 6-Swf 7-Jly 8-XYZ 9-SKy10-Unq 
      '  101    16    10     0     0     0     0     0     0     0 
      Jrn_Add("SDK_Space")
      Jrn_Add(, {"Compteurs des Stratégies       " & .Prd_Plcy_Strg_Profondeur})

      Dim S1 As String = "  "
      Dim S2 As String = "  "
      Dim S3 As String = "  "
      Dim S4 As String = "  "
      For i As Integer = 0 To 10
        S1 &= CStr(i).PadLeft(1) & "-" & Stg_List_Code(i) & " "
        S3 &= CStr(.Crt_Strg_Nb(i)).PadLeft(5) & " "
        S4 &= CStr(.Slv_Strg_Nb(i)).PadLeft(5) & " "
      Next i
      Jrn_Add(, {S1})
      Jrn_Add(, {S3})
      Jrn_Add(, {S4})
      Jrn_Add("SDK_Space")
      S1 = ""
      S2 = ""
      For i As Integer = 0 To 80
        S1 &= .Prd_Ini(i)
        S2 &= .Prd_Val(i)
      Next i

      Jrn_Add(, {"Ini " & S1.Replace(" ", ".")})
      Jrn_Add(, {"Val " & S2.Replace(" ", ".")})

      If .Prd_Ext_Triplet <> "#" Then Jrn_Add(, {"Triplet                      : " & .Prd_Ext_Triplet & " " & U_Coord(.Prd_Ext_Triplet_Cellule)})
      If .Prd_Ext_XWing <> "#" Then Jrn_Add(, {"Xwing                        : " & .Prd_Ext_XWing & " " & U_Coord(.Prd_Ext_XWing_Cellule)})
      If .Prd_DlCode <> "" Then Jrn_Add(, {"Prd_DlCode                   : " & .Prd_DlCode})
      If .Prd_DlCode = "Dlu" Then Jrn_Add(, {"Prd_DlSolution               : " & .Prd_DlSolution})
      Jrn_Add(, {"Origine Batch_Interactive    : " & .Prd_BI})
      Jrn_Add(, {"Prd_Code_Retour              : " & .Prd_Code_Retour})
    End With
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
#End Region

#Region "Structure_U_Check"
  Public Structure U_Check_Struct              '  Structure du Résultat de U_Checking
    Public Check As Boolean
    Public Nb_Initiales As Integer
    Public Nb_Vides As Integer
    Public Nb_Remplies As Integer
    Public Nb_Vides_Sans_Candidat As Integer
    Public Code_LCR As String
    Public LCR As Integer
    Public Valeur As Integer
    Public Ini_before As String
    Public Val_before As String
    Public Cdd_before As String
    Public Ini_after As String
    Public Val_after As String
    Public Cdd_after As String
  End Structure

  Function U_Checking(U_Chk(,) As String) As U_Check_Struct
    'U_checking permet de vérifier si une grille complète est correcte
    ' et si une grille incomplète est correcte au moment où elle est contrôlée
    'Une grille de Puzzle est complète ou incomplète
    'Un puzzle est un SuDoKu quand il n'a qu'une seule solution
    'U_Checking peut indiquer True une grille incomplète
    With U_Checking
      .Check = True
      .Nb_Initiales = 0
      .Nb_Vides = 0
      .Nb_Remplies = 0
      .Nb_Vides_Sans_Candidat = 0
      .Code_LCR = "_"
      .LCR = -1
      .Valeur = 0
      .Ini_before = ""
      .Val_before = ""
      .Cdd_before = ""
      .Ini_after = ""
      .Val_after = ""
      .Cdd_after = ""
    End With

    With U_Checking
      For i As Integer = 0 To 80
        .Ini_before &= U_Chk(i, 1)
        .Val_before &= U_Chk(i, 2)
        .Cdd_before &= U_Chk(i, 3)

        If U_Chk(i, 1) <> " " Then .Nb_Initiales += 1
        If U_Chk(i, 2) = " " Then .Nb_Vides += 1
        If U_Chk(i, 2) <> " " Then .Nb_Remplies += 1
      Next i
    End With
    ' Au moment du checking, il y a pour chaque cellule soit une valeur et aucun candidat
    '                                                   soit aucune valeur et un ou plusieurs candidats
    ' Or s'il y a un CdU, il faut considérer qu'il sera LA valeur et il faut supprimer cette valeur des 20 cel-col
    '    s'il y a un CdO, il faut considérer qu'il sera LA valeur et il faut supprimer cette valeur des 20 cel-col

    '1 Traitement temporaire des CdU_Ce traitement est fait UNE SEULE FOIS
    Dim Strategy_Rslt(99, 0) As String
    Strategy_Rslt = Strategy_CdU(U_Chk)
    Dim Cellule_CdU As Integer
    Dim Candidat_CdU As String
    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      Cellule_CdU = CInt(Strategy_Rslt(10, i))
      Candidat_CdU = Strategy_Rslt(5, i)
      U_Chk(Cellule_CdU, 2) = Candidat_CdU
      U_Chk(Cellule_CdU, 3) = Cnddts_Blancs
      Cdd_Remove_Cell_Coll_List(U_Chk, Cellule_CdU)
    Next i

    '2 Traitement temporaire des CdO_Ce traitement est fait UNE SEULE FOIS
    Strategy_Rslt = Strategy_CdO(U_Chk)
    Dim Cellule_CdO As Integer
    Dim Candidat_CdO As String
    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      Cellule_CdO = CInt(Strategy_Rslt(10, i))
      Candidat_CdO = Strategy_Rslt(5, i)
      U_Chk(Cellule_CdO, 2) = Candidat_CdO
      U_Chk(Cellule_CdO, 3) = Cnddts_Blancs
      Cdd_Remove_Cell_Coll_List(U_Chk, Cellule_CdO)
    Next i

    With U_Checking
      For i As Integer = 0 To 80
        .Ini_after &= U_Chk(i, 1)
        .Val_after &= U_Chk(i, 2)
        .Cdd_after &= U_Chk(i, 3)
      Next i
    End With

    '1 Existe-t'il des cellules vides sans candidats
    With U_Checking
      For i As Integer = 0 To 80
        If U_Chk(i, 2) = " " And U_Chk(i, 3) = "         " Then .Nb_Vides_Sans_Candidat += 1
      Next i
      If .Nb_Vides_Sans_Candidat <> 0 Then
        .Check = False
        GoTo U_Checking_End
      End If
    End With

    Dim Unité As String() = {"L", "C", "R"} 'Ligne, Colonne, Région
    For i As Integer = 0 To 2 'L, C, R 
      If Not U_Checking.Check Then Exit For

      Dim Code_LCR As String = Unité(i)
      If Not U_Checking.Check Then Exit For

      For LCR As Integer = 0 To 8
        If Not U_Checking.Check Then Exit For

        Dim Grp(0 To 8) As Integer
        Select Case Code_LCR
          Case "L" : Grp = U_9CelRow(LCR) 'Comporte les 9 cellules de la ligne
          Case "C" : Grp = U_9CelCol(LCR) '                              colonne
          Case "R" : Grp = U_9CelReg(LCR) '                              région
        End Select
        Dim Chk_Val(9), Chk_Cdd(9) As Integer 'Le poste 0 n'est pas utilisé
        Dim Cellule, Valeur As Integer
        For j As Integer = 0 To 8
          Cellule = Grp(j)
          If U_Chk(Cellule, 2) <> " " Then
            ' Remplissage des valeurs dans Chk_Val
            Valeur = CInt(U_Chk(Cellule, 2))
            Chk_Val(Valeur) += 1
          End If
          Dim Candidats, Candidat As String
          If U_Chk(Cellule, 2) = " " And U_Chk(Cellule, 3) <> Cnddts_Blancs Then
            Candidats = U_Chk(Cellule, 3)
            For cdd As Integer = 1 To 9
              Candidat = Mid$(Candidats, cdd, 1)
              If Candidat <> " " Then
                If CInt(Candidat) = cdd Then
                  ' Remplissage des candidats dans Chk_Cdd
                  Chk_Cdd(cdd) += 1
                End If
              End If
            Next cdd
          End If
        Next j

        ' Contrôle Existe-t'il des valeurs en double ou manque-t'il des candidats ?
        For k As Integer = 1 To 9
          If Chk_Val(k) > 1 _
          Or Chk_Val(k) = 0 And Chk_Cdd(k) = 0 Then
            With U_Checking
              .Check = False
              .Code_LCR = Code_LCR
              .LCR = LCR
              .Valeur = k
            End With
            Exit For
          End If
        Next k
      Next LCR
    Next i

U_Checking_End:
    Return U_Checking
  End Function

  Sub U_Checking_Display(U_Check As U_Check_Struct, Chat As Boolean)
    'Ajout d'un Mode Chat Bavard / Silencieux    True / False
    'Indication de la grille remplie, vide ou remplie à x%
    'Si le contrôle détecte une anomalie, celle-ci est signalée en erreur rouge
    Dim S As String = "#"
    Select Case Chat
      Case True
        Jrn_Add(, {"Checking_List (après calcul des CdU & CdO)"}, "Italique")
        With U_Check
          If .Nb_Remplies = 81 Then S = "Grille remplie"
          If .Nb_Vides = 81 - .Nb_Initiales Then S = "Grille vide"
          If .Nb_Remplies <> 81 And .Nb_Vides <> 81 - .Nb_Initiales Then
            S = "Grille remplie à " & Wh_Pourcentage()
          End If
          Select Case .Check
            Case True
              Jrn_Add(, {"Check               : " & " Correct. " & S})
            Case False
              Jrn_Add(, {"Check               : " & " ERREUR !"}, "Erreur")
          End Select
          Jrn_Add(, {"Initiales/Vides     : " & CStr(.Nb_Initiales) & "/" & CStr(.Nb_Vides)})
          Jrn_Add(, {"Remplies            : " & CStr(.Nb_Remplies)})
          Jrn_Add(, {"Vides_Sans_Candidat : " & CStr(.Nb_Vides_Sans_Candidat)})
          Jrn_Add(, {"Code_LCR            : " & .Code_LCR & CStr(.LCR + 1)})
          Jrn_Add(, {"Valeur              : " & CStr(.Valeur)})
          Jrn_Add(, {"Val before/after    : "})
          Jrn_Add(, {"bef " & .Val_before.Replace(" ", ".")})
          Jrn_Add(, {"aft " & .Val_after.Replace(" ", ".")})
        End With
        Jrn_Add(, {"/Checking_List "}, "Italique")

      Case False
    End Select
  End Sub
#End Region
End Module
