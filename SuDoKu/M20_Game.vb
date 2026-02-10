Option Strict On
Option Explicit On
Imports SuDoKu.DancingLink

Friend Module M20_Game

  Public Sub Game_New_Game(Gnrl As String,
                           Nom As String,
                           Prb As String,
                           Jeu As String,
                           Sol As String,
                           Cdd729 As String,
                           Frc As String)
    Jrn_Add("SDK_Space")
    Jrn_Add("SDK_00000", {Proc_Name_Get() & " Name: " & Nom & " Jeu: " & Prb.Substring(0, 9)})
    Plcy_Gnrl = Gnrl

    'Début commun à toute nouvelle partie, quoique seul Nrm utilise Annuler/Refaire
    ReDim Act(11, 100) : Act_Index = -1 ' Ré-initialise les mouvements
    UR_A_Index = -1
    Insert_Nb_Cell = 0
    Insertion_Exclusion_Nb_Erreurs = 0
    'En mode Sas, il ne reste que le menu Fichier
    '   TOUTES les options disposant d'un raccourci sont .Enabled = False

    Select Case Plcy_Gnrl
      Case "Nrm" '---------------------------------------------------------------------------------------
        Strategy_Switch("   ")
        Game_Load(Nom, Prb, Jeu, Sol, Frc)
        Grid_Cdd_Remove_Cell_Coll(U)

        If Cdd729.Length = 729 And Cdd729 <> StrDup(729, " ") Then
          For i As Integer = 0 To 80
            If U(i, 2) = " " Then
              U(i, 3) = Cdd729.Substring(i * 9, 9)
            End If
          Next i
        End If
        Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})

        If Plcy_Solution_Existante = True Then
          Frm_SDK.B_Solution.Text = "S"
          Frm_SDK.Mnu02_InsérerLaSolution.Enabled = True
          Frm_SDK.Mnu03_AfficherLaSolution.Enabled = True
          Frm_SDK.Mnu08_InsérerTouteLaSolution.Enabled = True
        Else
          Frm_SDK.B_Solution.Text = " "
          Frm_SDK.Mnu02_InsérerLaSolution.Enabled = False
          Frm_SDK.Mnu03_AfficherLaSolution.Enabled = False
          Frm_SDK.Mnu08_InsérerTouteLaSolution.Enabled = False
        End If

      Case "Sas" '---------------------------------------------------------------------------------------
        Strategy_Switch("   ")
        Game_Load(Nom, Prb, Jeu, Sol, Frc)
        '   Si le jeu a été fermé en mode Sas, alors il est ouvert en mode Sas et les Candidats sont replacés
        '   Dans une partie "Sas", SDK ignore et ne calcule pas les candidats 
        For i As Integer = 0 To 80
          If U(i, 2) = " " Then
            U(i, 3) = Cdd729.Substring(i * 9, 9).Replace("0", " ")
          End If
        Next i
        Frm_SDK.B_Solution.Text = "*"
        Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_00300")
    End Select

    'OC_Présentation()

    'Fin commune à toute nouvelle partie
    Game_Nb_Cellules_Initiales = Wh_Nb_Cell(U).Initiales

    Dim ToolTipText As String = Nothing

    Select Case Frm_SDK.B_Solution.Text
      Case " " : ToolTipText = "Puzzle sans solution"
      Case "*" : ToolTipText = "Mode Sans Assistance"
      Case "S" : ToolTipText = "Le Puzzle a une solution"
    End Select
    Frm_SDK.ToolTip_B.SetToolTip(Frm_SDK.B_Solution, ToolTipText)

    Paint_Partie_Terminée_Nb = 0
    'Lors d'une nouvelle partie, une cellule vide est sélectionnée au hasard comme Cellule Sélectionnée 
    Pbl_Cell_Select = Wh_RandomCelluleVide()
    Prv_Pbl_Cell_Select = Pbl_Cell_Select
    Frm_SDK.B_Pourcentage.Text = Wh_Pourcentage()

    Event_OnPaint = "Total"
    Event_OnPaint_MAP = Proc_Name_Get()
    Frm_SDK.Invalidate()


  End Sub

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
      If Plcy_Gnrl = "Sas" Then
        U(i, 3) = Cnddts_Blancs
        U_CddExc(i) = Cnddts
      Else
        U(i, 3) = Cnddts
        U_CddExc(i) = Cnddts_Blancs
      End If
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
    If Plcy_Dancing_Link Then
      ' Vérification d'une solution unique
      '29/07/2025 Dans le cadre des tests XChains, il peut être possible de vérifier la solution! 
      XSolution = StrDup(81, "0")
      Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve_IA(U)
      Select Case DL.Nb_Solution
        Case -1, 0
        Case 1
          XSolution = DL.Solution(0)
        Case Else
          Dim MsgTit As String = Application.ProductName & " " & SDK_Version
          Nsd_i = MsgBox("Dancing Link : " & DL.DLCode & "  Solutions multiples.",, MsgTit)
      End Select
    End If

    Frm_SDK.Text = LP_Nom
  End Sub

End Module