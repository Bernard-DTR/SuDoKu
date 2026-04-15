Imports SuDoKu.DancingLink

Friend Module M20_Game
  Public Sub Game_New_Game(Gnrl As String,
                           Strg As String,
                           Nom As String,
                           Prb As String,
                           Jeu As String,
                           Sol As String,
                           Cdd729 As String,
                           Frc As String,
                           Origine As String)
    Jrn_Add("SDK_Space")
    Jrn_Add("SDK_00000", {Proc_Name_Get() & " Name: " & Nom & " Jeu: " & Prb.Substring(0, 9)})
    Plcy_Gnrl = Gnrl

    'Début commun à toute nouvelle partie, quoique seul Nrm utilise Annuler/Refaire
    ReDim Act(11, 100) : Act_Index = -1 ' Ré-initialise les mouvements
    UR_A_Index = -1

    Select Case Plcy_Gnrl
      Case "Nrm" '---------------------------------------------------------------------------------------
        Plcy_Strg = Strg
        Game_Load(Nom, Prb, Jeu, Sol, Frc)
        Grid_Cdd_Remove_Cell_Coll(U)

        If Cdd729.Length = 729 And Cdd729 <> StrDup(729, " ") Then
          For i As Integer = 0 To 80
            If U(i, 2) = " " Then
              U(i, 3) = Cdd729.Substring(i * 9, 9)
            End If
          Next i
        End If
    End Select

    Game_Nb_Cellules_Initiales = Wh_Nb_Cell(U).Initiales

    Paint_Partie_Terminée_Nb = 0
    Pbl_Cell_Select = 0
    Prv_Pbl_Cell_Select = Pbl_Cell_Select
    Mnu_Mngt_Barre_Outils_Filtres()
    Game_New_Game_DL_Solution(U)
    Frm_SDK.B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()
    Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Frm_SDK.Text = LP_Nom
    Cpt_Pénalités = 0

    Event_OnPaint_Origine = Proc_Name_Get() & " Origine : " & Origine
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub

  Public Function Game_New_Game_DL_Solution(U(,) As String) As Integer
    ' La fonction calcule XSolution qui la solution Dancing_Link
    ' soit 000000000000000000000000000000000000000000000000000000000000000000000000000000000
    ' soit la solution unique
    ' et retourne le nombre de solutions,
    '                -2 non calculé,
    '                -1,0 en erreur,
    '                1 est attendu de préférence
    '                x nombre de solution
    XSolution = StrDup(81, "0")
    If Plcy_Dancing_Link Then
      Dim MsgTit As String = Application.ProductName & " " & SDK_Version
      Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U)
      Select Case DL.Nb_Solution
        Case -1, 0
          Nsd_i = MsgBox("Dancing Link : " & DL.DLCode & "  Solution Impossible à calculer pour DL !",, MsgTit)
        Case 1
          XSolution = DL.Solution(0)
        Case Else
          Nsd_i = MsgBox("Dancing Link : " & DL.DLCode & "  Solutions multiples.",, MsgTit)
      End Select
      Return DL.Nb_Solution
    End If
    Return -2
  End Function

  Public Sub Game_Load(Nom As String,
                       Prb As String,
                       Jeu As String,
                       Sol As String,
                       Frc As String)
    LP_Nom = "#" : LP_Frc = "#"
    LP_Nom = Nom
    LP_Frc = Frc
    LP_Prb = Prb : LP_Prb = LP_Prb.Replace("0", " ") : LP_Prb = LP_Prb.Replace(".", " ")
    LP_Jeu = Jeu : LP_Jeu = LP_Jeu.Replace("0", " ") : LP_Jeu = LP_Jeu.Replace(".", " ")
    LP_Sol = Sol : LP_Sol = LP_Sol.Replace("0", " ") : LP_Sol = LP_Sol.Replace(".", " ")

    ' Un Sudoku comporte un problème ET une solution
    '    (la règle n'oblige pas à ce que la solution soit LA solution)
    '08/02/2025
    Plcy_Solution_Existante = False
    If LP_Sol.Contains(" ") = False Then Plcy_Solution_Existante = True
    LP_Frc = Frc

    ' Initialisation des postes U 1 à 3 de U sont initialisés et ensuite chargés
    For i As Integer = 0 To 80
      U(i, 1) = " "                         ' Valeur Initiale
      U(i, 2) = " "                         ' Valeur
      U(i, 3) = Cnddts
      U_CddExc(i) = Cnddts_Blancs
      U_Sol(i) = " "                        ' Solution

      ' 08/02/2025 Copilot LP_Prb(i) au lieu de LP_Prb.substring(i,1)  
      ' Met à jour les valeurs et les solutions si les conditions sont remplies
      If LP_Jeu(i) >= "1" AndAlso LP_Jeu(i) <= "9" Then
        U(i, 1) = LP_Prb(i)                         ' Valeur Initiale
        U(i, 2) = LP_Jeu(i)                         ' Valeur dans la Grille et U
        U(i, 3) = Cnddts_Blancs                     ' Candidats
        U_CddExc(i) = Cnddts_Blancs                 ' Candidats exclus
      End If
      If LP_Sol(i) >= "1" AndAlso LP_Sol(i) <= "9" Then
        U_Sol(i) = LP_Sol(i)                        ' Solution
      End If
      '#536
      If U(i, 1) = " " Then
        U_Clr_Cell_Fond(i) = Color_Fond_Typ_RV
        U_Clr_Cell_Val(i) = Color_VCdd
      Else
        U_Clr_Cell_Fond(i) = Color_Fond_Typ_I
        U_Clr_Cell_Val(i) = Color_VI
      End If
    Next i

  End Sub

End Module