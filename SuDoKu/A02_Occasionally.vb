Option Strict On
Option Explicit On

Imports System.Drawing.Drawing2D
Imports System.Text

'------------------------------------------------------------------------------------------
' Date de création: 28/07/2024
' Ce Module regroupe les traitements effectués lors des changements de présentation
'   Couleur, WH, ...etc
'            Color_Frm_BackColor = Color.FromArgb(255, n'est plus transparent
'------------------------------------------------------------------------------------------

Friend Module A02_Occasionally
  Public Sub OC_Thèmes_Couleurs(Thème As Integer)
    Obj_Symbol = My.Settings.Obj_Symbol
    Obj_Forme = My.Settings.Obj_Forme

    Select Case Thème
      Case 0
        Color_Frm_BackColor = Color.FromArgb(255, 216, 245, 216)    ' Couleur Fond du formulaire et du Grid
        Color_Trait = Color.Green                                   ' Couleur des traits du Grid
        Color_Fond_Typ_I = Color.FromArgb(255, 192, 255, 192)       ' Couleur Fond Valeurs Initiales
        Color_Fond_Typ_RV = Color.FromArgb(255, 129, 224, 129)      ' Couleur Fond Cellule Remplie/Vide 
        Color_Stratégique = Color.FromArgb(255, 15, 196, 101)       ' Couleur Couche Stratégique
        Color_VI = Color.Green                                      ' Couleur des valeurs initiales
        Color_VCdd = Color.Blue                                     ' Couleur des valeurs et des candidats
      Case 1
        Color_Frm_BackColor = Color.FromArgb(255, 216, 245, 242)
        Color_Trait = Color.Blue
        Color_Fond_Typ_I = Color.FromArgb(255, 192, 255, 250)
        Color_Fond_Typ_RV = Color.FromArgb(255, 128, 193, 225)
        Color_Stratégique = Color.FromArgb(255, 9, 89, 149)
        Color_VI = Color.Fuchsia
        Color_VCdd = Color.Purple
      Case 2
        Color_Frm_BackColor = Color.Beige
        Color_Trait = Color.FromArgb(255, 255, 204, 0)
        Color_Fond_Typ_I = Color.FromArgb(255, 217, 179, 179) '    Color.FromArgb(255, 184, 180, 131)
        Color_Fond_Typ_RV = Color.FromArgb(255, 189, 185, 138)
        Color_Stratégique = Color.FromArgb(255, 162, 100, 94)
        Color_VI = Color.Navy
        Color_VCdd = Color.Red
      Case Else 'Identique à Standard
        Color_Frm_BackColor = Color.FromArgb(255, 216, 245, 216)
        Color_Trait = Color.Green
        Color_Fond_Typ_I = Color.FromArgb(255, 192, 255, 192)
        Color_Fond_Typ_RV = Color.FromArgb(255, 129, 224, 129)
        Color_Stratégique = Color.FromArgb(255, 15, 196, 101)
        Color_VI = Color.Green
        Color_VCdd = Color.Blue
    End Select
  End Sub

  Sub OC_Plcy_Stg_UOBTXYSJZKQ()
    ' 1 à/p de Str_List, il faut déterminer le nombre de stratégies concernées par la Production-Résolution
    Dim Stg_Nb As Integer = 0
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Prd = "O" Then Stg_Nb += 1
    Next stg
    ' 2  Dimensionnement 
    ReDim Stg_Bll(0 To Stg_Nb - 1)
    Stg_Profondeur = StrDup(Stg_Nb, "_")
    Dim sb As New StringBuilder(Stg_Profondeur)
    Dim Index As Integer = 0
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Prd = "O" Then
        Dim LettreT As String = Stg.Lettre
        If Mid(Plcy_Stg_Clb, Index + 1, 1) = "O" Then
          sb(Index) = Char.ToUpper(LettreT(0)) ' Utilisation de Char.ToUpper pour convertir en majuscule
          Stg_Bll(Index) = True
        Else
          sb(Index) = Char.ToLower(LettreT(0)) ' Utilisation de Char.ToLower pour convertir en minuscule
          Stg_Bll(Index) = False
        End If
        Index += 1
      End If
    Next Stg
    Stg_Profondeur = sb.ToString()

    ' 3  Gestion des boutons des stratégies Enabled/Unelabled de la barre d'outils
    'Stg_List_Code: CdU , CdO, Cbl, Tpl, Xwg, XYw, Swf, Jly, XYZ, SKy, Unq
    For i As Integer = 0 To Stg_Bll.Count - 1
      Dim Btn_Name As String = "Btn_" & Stg_List_Code.Item(i)
      If Frm_SDK.BarreOutils.Items.ContainsKey(Btn_Name) Then
        If Stg_Bll(i) = True Then Frm_SDK.BarreOutils.Items(Btn_Name).Enabled = True
        If Stg_Bll(i) = False Then Frm_SDK.BarreOutils.Items(Btn_Name).Enabled = False
      End If
    Next

  End Sub

  Public Sub OC_Présentation()
    OC_Grid_Compute()
    OC_U_Pt20_Init()
    OC_Grid_Compute_Font_Size_IA()
    OC_Grid_Cutting_Image()
    OC_Présentation_SDK()
    OC_Présentation_Menu()
  End Sub
  Public Sub OC_Présentation_SDK()
    Dim Int_Seize As Integer = 16
    'Les infos B_*
    Dim B_Width As Integer = Bld_Marge_LT
    Dim B_Top As Integer
    Dim B_Height As Integer = 20 'Hauteur des zones B_*

    If Plcy_Gnrl = "Nrm" Then Frm_SDK.BarreOutils.Visible = True
    If Plcy_Gnrl = "Sas" Then Frm_SDK.BarreOutils.Visible = False
    If Plcy_Gbl_Etendue Then Frm_SDK.Journal.Visible = True Else Frm_SDK.Journal.Visible = False


    Frm_SDK.BackColor = Color_Frm_BackColor

    ' Dimensions du Formulaire Frm_SDK
    '   La largeur de Frm_SDK comprend le journal
    '   La hauteur            comprend la Barre d'Outils
    With Frm_SDK
      .Width = Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT + Bld_Journal_Affiché_Width + Bld_Marge_LT + Int_Seize
      .Height = SI_CaptionHeight + Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT + Bld_WH_Grid _
              + Bld_Marge_LT + B_Height + Bld_Marge_LT + Int_Seize
    End With
    If Not Plcy_Gbl_Etendue And (Plcy_Gnrl = "Nrm" Or Plcy_Gnrl = "Sas") Then
      Frm_SDK.Width = Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT + Int_Seize
    End If

    ' L'emplacement du journal dépend de l'affichage ou non de la barre d'outils
    With Frm_SDK.Journal
      .Location = New Point(Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT, Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT)
      .Size = New Drawing.Size(Bld_Journal_Width, Bld_WH_Grid + Bld_Marge_LT + B_Height)
      .BackColor = Color_Fond_Typ_I
    End With

    ' les infos B_*
    B_Top = Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT
    With Frm_SDK
      .B_Solution.Location = New Point(B_Width, B_Top)
      .B_Solution.Size = New Size(20, B_Height)

      B_Width = Bld_Marge_LT + .B_Solution.Width
      .B_Position.Location = New Point(B_Width, B_Top)
      .B_Position.Size = New Size(60, B_Height)

      B_Width = Bld_Marge_LT + .B_Solution.Width + .B_Position.Width
      .B_Pourcentage.Location = New Point(B_Width, B_Top)
      .B_Pourcentage.Size = New Size(40, B_Height)

      B_Width = Bld_Marge_LT + .B_Solution.Width + .B_Position.Width + .B_Pourcentage.Width
      .B_Info.Location = New Point(B_Width, B_Top)
      .B_Info.Size = New Size(Bld_Marge_LT + Bld_WH_Grid - B_Width, B_Height)
      .B_Info.Visible = True
      ' La progress Bar se substitue à la zone B_Info dans les traitements longues durées 
      .B_ProgressBar.Location = New Point(B_Width, B_Top)
      .B_ProgressBar.Size = New Size(Bld_Marge_LT + Bld_WH_Grid - B_Width, B_Height)
      .B_ProgressBar.Visible = False
    End With

  End Sub

  Public Sub OC_Présentation_Menu()

    ' Colorisation de 4 options de menu
    With Frm_SDK
      .Mnu05_AideSudokuGraphique.BackColor = Color_Stratégique
      .Mnu_EDI_Saisir_Valeur.BackColor = Color_Fond_Typ_RV
      .Mnu_EDI_Val_Normale.BackColor = Color_Fond_Typ_RV
      .Mnu_EDI_Val_Initiale.BackColor = Color_Fond_Typ_I
    End With

    ' Libellé Jen Normal / Sans Assistance
    With Frm_SDK
      Select Case Plcy_Gnrl
        Case "Nrm" : .Mnu01_JeuSansAssistance.Text = "Jeu Sans Assistance"
        Case "Sas" : .Mnu01_JeuSansAssistance.Text = "Jeu Normal"
      End Select
    End With

    ' Options disponibles ou non
    If Plcy_Gbl_Etendue And Plcy_Gnrl = "Nrm" Then
      With Frm_SDK
        'Mnu01 Fichier
        .Mnu01_Ouvrir.Visible = True
        .Mnu01_RejouerLaPartie.Visible = True
        .Mnu01_Sep01.Visible = True
        .Mnu01_Saisir.Visible = True
        .Mnu01_Commencer.Visible = True
        .Mnu01_Sep03.Visible = True
        .Mnu01_EnregistrerUnePartieTest.Visible = True
        .Mnu01_ChargerUnePartieTest.Visible = True
        .Mnu01_OuvrirLaBibliothèqueTestDeHodoku.Visible = True
        .Mnu01_Sep04.Visible = True
        .Mnu01_OuvrirLeRépertoire.Visible = True
        .Mnu01_Sep05.Visible = True
        .Mnu01_CopierLeJournalEnModeRTF.Visible = True
        .Mnu01_Sep06.Visible = True
        .Mnu01_JeuSansAssistance.Visible = True
        .Mnu01_Sep07.Visible = True
        .Mnu01_Quitter.Visible = True
        'Mnu02 Edition
        .Mnu02.Visible = True
        .Mnu02_Annuler.Visible = True
        .Mnu02_Refaire.Visible = True
        .Mnu02_Sep01.Visible = True
        .Mnu02_Effacer.Visible = True
        .Mnu02_InsérerLaSolution.Visible = True
        .Mnu02_Sep03.Visible = True
        .Mnu02_Copier.Visible = True
        .Mnu02_Copier2.Visible = True
        .Mnu02_Coller.Visible = True
        .Mnu02_CopierlaGrilleDansLeJournal.Visible = True
        'Mnu03 Affichage
        .Mnu03.Visible = True
        .Mnu03_EffacerLeJournal.Visible = True
        .Mnu03_Sep01.Visible = True
        .Mnu03_Transformation.Visible = True
        .Mnu03_AfficherLaSolution.Visible = True
        .Mnu03_Rafraîchir.Visible = True
        'Mnu04_Stratégies
        .Mnu04.Visible = True
        'Mnu08 Extension
        .Mnu08.Visible = True
        .Mnu08_Création.Visible = True
        .Mnu08_Résolution.Visible = True
        .Mnu08_RésoudreEnForceBrute.Visible = True
        .Mnu08_RésoudreDancingLink.Visible = True
        .Mnu08_Sep01.Visible = True
        .Mnu08_EditionDuProblème.Visible = True
        .Mnu08_DessinerSurLaGrille.Visible = True
        .Mnu08_Sep03.Visible = True
        .Mnu08_InsérerTouteLaSolution.Visible = True
        .Mnu08_Sep04.Visible = True
        .Mnu08_TestA.Visible = True
        .Mnu08_TestB.Visible = True
        .Mnu08_TestC.Visible = True
        .Mnu08_TestD.Visible = True
        .Mnu08_TestE.Visible = True
        .Mnu08_TestF.Visible = True
        .Mnu08_TestG.Visible = True
        .Mnu08_TestH.Visible = True
        .Mnu08_TestI.Visible = True
        .Mnu08_TestJ.Visible = True
        .Mnu08_TestA.Text = Msg_Read_IA("MNU_0800A")
        .Mnu08_TestB.Text = Msg_Read_IA("MNU_0800B")
        .Mnu08_TestC.Text = Msg_Read_IA("MNU_0800C")
        .Mnu08_TestD.Text = Msg_Read_IA("MNU_0800D")
        .Mnu08_TestE.Text = Msg_Read_IA("MNU_0800E")
        .Mnu08_TestF.Text = Msg_Read_IA("MNU_0800F")
        .Mnu08_TestG.Text = Msg_Read_IA("MNU_0800G")
        .Mnu08_TestH.Text = Msg_Read_IA("MNU_0800H")
        .Mnu08_TestI.Text = Msg_Read_IA("MNU_0800I")
        .Mnu08_TestJ.Text = Msg_Read_IA("MNU_0800J")
        'Mnu05 Aide
        .Mnu05.Visible = True
        .Mnu05_AideSudokuGraphique.Visible = True
        .Mnu05_Préférences.Visible = True
        .Mnu05_FichierDesMessages.Visible = True
        .Mnu05_APropos.Visible = True
        .Mnu05_Documentation.Visible = True
        .Mnu05_Maintenance.Visible = True
        .Mnu05_Dictionnaire.Visible = True
        'Mnu06 Divers
        .Mnu06.Visible = True
        'Mnu07 Outils
        .Mnu07.Visible = True
      End With
    End If
    If Plcy_Gbl_Etendue And Plcy_Gnrl = "Sas" Then
      With Frm_SDK
        'Mnu01 Fichier
        .Mnu01_Ouvrir.Visible = True
        .Mnu01_RejouerLaPartie.Visible = True
        .Mnu01_Sep01.Visible = True
        .Mnu01_Saisir.Visible = True
        .Mnu01_Commencer.Visible = True
        .Mnu01_Sep03.Visible = False
        .Mnu01_EnregistrerUnePartieTest.Visible = False
        .Mnu01_ChargerUnePartieTest.Visible = False
        .Mnu01_OuvrirLaBibliothèqueTestDeHodoku.Visible = False
        .Mnu01_Sep04.Visible = False
        .Mnu01_OuvrirLeRépertoire.Visible = False
        .Mnu01_Sep05.Visible = False
        .Mnu01_CopierLeJournalEnModeRTF.Visible = False
        .Mnu01_Sep06.Visible = True
        .Mnu01_JeuSansAssistance.Visible = True
        .Mnu01_Sep07.Visible = True
        .Mnu01_Quitter.Visible = True
        'Mnu02 Edition
        .Mnu02.Visible = False
        'Mnu03 Affichage
        .Mnu03.Visible = False
        'Mnu04_Stratégies
        .Mnu04.Visible = False
        'Mnu08 Extension
        .Mnu08.Visible = False
        'Mnu05 Aide
        .Mnu05.Visible = False
        'Mnu06 Divers
        .Mnu06.Visible = False
        'Mnu07 Outils
        .Mnu07.Visible = False
      End With
    End If
    If Not Plcy_Gbl_Etendue And Plcy_Gnrl = "Nrm" Then
      With Frm_SDK
        'Mnu01 Fichier
        .Mnu01_Ouvrir.Visible = True
        .Mnu01_RejouerLaPartie.Visible = True
        .Mnu01_Sep01.Visible = True
        .Mnu01_Saisir.Visible = True
        .Mnu01_Commencer.Visible = True
        .Mnu01_Sep03.Visible = False
        .Mnu01_EnregistrerUnePartieTest.Visible = False
        .Mnu01_ChargerUnePartieTest.Visible = False
        .Mnu01_OuvrirLaBibliothèqueTestDeHodoku.Visible = False
        .Mnu01_Sep04.Visible = False
        .Mnu01_OuvrirLeRépertoire.Visible = False
        .Mnu01_Sep05.Visible = False
        .Mnu01_CopierLeJournalEnModeRTF.Visible = False
        .Mnu01_Sep06.Visible = True
        .Mnu01_JeuSansAssistance.Visible = True
        .Mnu01_Sep07.Visible = True
        .Mnu01_Quitter.Visible = True
        'Mnu02 Edition
        .Mnu02.Visible = True
        .Mnu02_Annuler.Visible = True
        .Mnu02_Refaire.Visible = True
        .Mnu02_Sep01.Visible = True
        .Mnu02_Effacer.Visible = True
        .Mnu02_InsérerLaSolution.Visible = True
        .Mnu02_Sep03.Visible = False
        .Mnu02_Copier.Visible = False
        .Mnu02_Copier2.Visible = False
        .Mnu02_Coller.Visible = False
        .Mnu02_CopierlaGrilleDansLeJournal.Visible = False
        'Mnu03 Affichage
        .Mnu03.Visible = True
        .Mnu03_EffacerLeJournal.Visible = False
        .Mnu03_Sep01.Visible = True
        .Mnu03_Transformation.Visible = True
        .Mnu03_AfficherLaSolution.Visible = True
        .Mnu03_Rafraîchir.Visible = True
        'Mnu04 Stratégies
        .Mnu04.Visible = True
        'Mnu08 Extension
        .Mnu08.Visible = True
        .Mnu08_Création.Visible = True
        .Mnu08_Résolution.Visible = True
        .Mnu08_RésoudreEnForceBrute.Visible = True
        .Mnu08_RésoudreDancingLink.Visible = True
        .Mnu08_Sep01.Visible = False
        .Mnu08_EditionDuProblème.Visible = False
        .Mnu08_DessinerSurLaGrille.Visible = False
        .Mnu08_Sep03.Visible = True
        .Mnu08_InsérerTouteLaSolution.Visible = True
        .Mnu08_Sep04.Visible = False
        .Mnu08_TestA.Visible = False
        .Mnu08_TestB.Visible = False
        .Mnu08_TestC.Visible = False
        .Mnu08_TestD.Visible = False
        .Mnu08_TestE.Visible = False
        .Mnu08_TestF.Visible = False
        .Mnu08_TestG.Visible = False
        .Mnu08_TestH.Visible = False
        .Mnu08_TestI.Visible = False
        .Mnu08_TestJ.Visible = False
        'Mnu05 Aide
        .Mnu05.Visible = True
        .Mnu05_AideSudokuGraphique.Visible = True
        .Mnu05_Préférences.Visible = True
        .Mnu05_FichierDesMessages.Visible = False
        .Mnu05_APropos.Visible = True
        .Mnu05_Documentation.Visible = False
        .Mnu05_Maintenance.Visible = False
        .Mnu05_Dictionnaire.Visible = False
        'Mnu06 Divers
        .Mnu06.Visible = False
        'Mnu07 Outils
        .Mnu07.Visible = False
      End With
    End If
    If Not Plcy_Gbl_Etendue And Plcy_Gnrl = "Sas" Then
      With Frm_SDK
        'Mnu01 Fichier
        .Mnu01_Ouvrir.Visible = True
        .Mnu01_RejouerLaPartie.Visible = True
        .Mnu01_Sep01.Visible = True
        .Mnu01_Saisir.Visible = True
        .Mnu01_Commencer.Visible = True
        .Mnu01_Sep03.Visible = False
        .Mnu01_EnregistrerUnePartieTest.Visible = False
        .Mnu01_ChargerUnePartieTest.Visible = False
        .Mnu01_OuvrirLaBibliothèqueTestDeHodoku.Visible = False
        .Mnu01_Sep04.Visible = False
        .Mnu01_OuvrirLeRépertoire.Visible = False
        .Mnu01_Sep05.Visible = False
        .Mnu01_CopierLeJournalEnModeRTF.Visible = False
        .Mnu01_Sep06.Visible = True
        .Mnu01_JeuSansAssistance.Visible = True
        .Mnu01_Sep07.Visible = True
        .Mnu01_Quitter.Visible = True
        'Mnu02 Edition
        .Mnu02.Visible = False
        'Mnu03 Affichage
        .Mnu03.Visible = False
        'Mnu04_Stratégies
        .Mnu04.Visible = False
        'Mnu08 Extension
        .Mnu08.Visible = False
        'Mnu05 Aide
        .Mnu05.Visible = False
        'Mnu06 Divers
        .Mnu06.Visible = False
        'Mnu07 Outils
        .Mnu07.Visible = False
      End With
    End If

    ' Libellé des Menus en fonction de la taille de WH
    Dim l As Integer
    With Frm_SDK
      Select Case WH
        Case 40 To 49 : l = 3
        Case 50 To 100 : l = 10
      End Select
      'PadRight  permet d'avoir un libellé de menu inférieur à 10 caractères
      'Substring permet d'avoir un menu court et un menu long
      .Mnu01.Text = Msg_Read_IA("MNU_01000").PadRight(l).Substring(0, l)
      .Mnu02.Text = Msg_Read_IA("MNU_02000").PadRight(l).Substring(0, l)
      .Mnu03.Text = Msg_Read_IA("MNU_03000").PadRight(l).Substring(0, l)
      .Mnu04.Text = Msg_Read_IA("MNU_04000").PadRight(l).Substring(0, l)
      .Mnu08.Text = Msg_Read_IA("MNU_08000").PadRight(l).Substring(0, l)
      .Mnu05.Text = Msg_Read_IA("MNU_05000").PadRight(l).Substring(0, l)
      .Mnu06.Text = Msg_Read_IA("MNU_06000").PadRight(l).Substring(0, l)
      .Mnu07.Text = Msg_Read_IA("MNU_07000").PadRight(l).Substring(0, l)
    End With

    ' Changement de la police dans Mnu de Frm_Sdk et Mnu04
    '                        dans Mnu_Cel
    '                        dans Mnu_EDI
    '                        dans Mnu_Journal
    Dim Item As ToolStripItem
    'Les menus, options et sous-options sont passées dans la font souhaitée
    'La font par défaut des menus est New System.Drawing.Font("Tahoma", 8.25!)
    'Barre de Menu de Frm_SDK
    For Each Item In Frm_SDK.Mnu.Items : Item.Font = Font_Mnu : Next Item
    For Each Item In Frm_SDK.Mnu04.DropDownItems : Item.Font = Font_Mnu : Next Item
    'Le Menu EDITION Contractuel garde la police standard proportionnelle de Frm_SDK_Mnu
    For Each Item In Frm_SDK.Mnu_EDI.Items : Item.Font = Font_Mnu : Next Item
    'Le Menu Journal Contractuel garde la police standard proportionnelle de Frm_SDK_Mnu
    For Each Item In Frm_SDK.Mnu_Journal.Items : Item.Font = Font_Mnu : Next Item
    'Menu contextuel d'une cellule de la grille
    For Each Item In Frm_SDK.Mnu_Cel.Items : Item.Font = Font_Mnu_Cel : Next Item

  End Sub

#Region " OC_Compute Sqr_Cel Font Fond d'Image et Sqr_Cdd"
  Public Sub OC_Grid_Compute()
    ' Généralités (Dépendance de WH)
    WHhalf = (WH \ 2)
    WHthird = (WH \ 3)
    WHquart = (WH \ 4)
    Bld_WH_Grid = (WH * 9) + 3 + 1 + 1 + 3 + 1 + 1 + 3 + 1 + 1 + 3

    ' La hauteur de la Barre_Outils est soit standard (25) et affichée, soit 0 et dans ce cas non affichée
    If Plcy_Gnrl = "Nrm" Then Barre_Outils_Hauteur = Barre_Outils_Standard
    If Plcy_Gnrl = "Sas" Then Barre_Outils_Hauteur = 0
    If Plcy_Gbl_Etendue Then Bld_Journal_Affiché_Width = Bld_Journal_Width Else Bld_Journal_Affiché_Width = 0

    ' Calcul de Gz_Pt_TopLeft As Point   
    ' Définition du Point Top-Left, Représente une paire ordonnée de coordonnées x et y entiers
    Gz_Pt_TopLeft = New Point(Bld_Marge_LT,
                              Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT)

    ' Définition des traits Gz_Trait_Pos_xy()
    Dim Ep_2 As Integer = 2 ' = 2   car le trait a une épaisseur de 3
    Dim Ep_1 As Integer = 1 ' = 1   car le trait a une épaisseur de 1

    Gz_Trait_Pos_xy(0) = Ep_1
    Gz_Trait_Pos_xy(1) = Ep_2 + WH + Gz_Trait_Pos_xy(0)
    Gz_Trait_Pos_xy(2) = Ep_1 + WH + Gz_Trait_Pos_xy(1)
    Gz_Trait_Pos_xy(3) = Ep_2 + WH + Gz_Trait_Pos_xy(2)
    Gz_Trait_Pos_xy(4) = Ep_2 + WH + Gz_Trait_Pos_xy(3)
    Gz_Trait_Pos_xy(5) = Ep_1 + WH + Gz_Trait_Pos_xy(4)
    Gz_Trait_Pos_xy(6) = Ep_2 + WH + Gz_Trait_Pos_xy(5)
    Gz_Trait_Pos_xy(7) = Ep_2 + WH + Gz_Trait_Pos_xy(6)
    Gz_Trait_Pos_xy(8) = Ep_1 + WH + Gz_Trait_Pos_xy(7)
    Gz_Trait_Pos_xy(9) = Ep_2 + WH + Gz_Trait_Pos_xy(8)

    ' Calcul des 81 Squares (9x9, ils sont TOUS à Angles Droits)
    '   Les Sqr_Cel semblent corrects, surface maximale et aucun recouvrement des traits.
    Dim v, h, k As Integer
    For j As Integer = 0 To 8      ' Axe des y, vertical
      Select Case j
        Case 0, 3, 6, 9
          v = 2
        Case 1, 2, 4, 5, 7, 8
          v = 1
      End Select
      For i As Integer = 0 To 8  ' Axe des x, horizontal
        Select Case i
          Case 0, 3, 6, 9
            h = 2
          Case 1, 2, 4, 5, 7, 8
            h = 1
        End Select
        Sqr_Cel(k) = New Rectangle(x:=Gz_Pt_TopLeft.X + Gz_Trait_Pos_xy(i) + h,
                                   y:=Gz_Pt_TopLeft.Y + Gz_Trait_Pos_xy(j) + v, width:=WH, height:=WH)
        k += 1
      Next i
    Next j

    ' Calcul des 9 Squares de chaque candidat   1 2 3  ou  7 8 9
    '                                           4 5 6      4 5 6
    '                                           7 8 9      1 2 3
    ' Identification du square de chaque candidat : Cellule * 10  + Candidat 
    ' Position des candidats : de gauche à droite et de bas en haut
    k = 0
    For j As Integer = 0 To 8      ' Axe des y, vertical
      For i As Integer = 0 To 8    ' Axe des x, horizontal
        Sqr_Cdd((k * 10) + 1) = New Rectangle(x:=Sqr_Cel(k).X + (0 * WHthird), y:=Sqr_Cel(k).Y + (0 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 1
        Sqr_Cdd((k * 10) + 2) = New Rectangle(x:=Sqr_Cel(k).X + (1 * WHthird), y:=Sqr_Cel(k).Y + (0 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 2
        Sqr_Cdd((k * 10) + 3) = New Rectangle(x:=Sqr_Cel(k).X + (2 * WHthird), y:=Sqr_Cel(k).Y + (0 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 3
        Sqr_Cdd((k * 10) + 4) = New Rectangle(x:=Sqr_Cel(k).X + (0 * WHthird), y:=Sqr_Cel(k).Y + (1 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 4
        Sqr_Cdd((k * 10) + 5) = New Rectangle(x:=Sqr_Cel(k).X + (1 * WHthird), y:=Sqr_Cel(k).Y + (1 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 5
        Sqr_Cdd((k * 10) + 6) = New Rectangle(x:=Sqr_Cel(k).X + (2 * WHthird), y:=Sqr_Cel(k).Y + (1 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 6
        Sqr_Cdd((k * 10) + 7) = New Rectangle(x:=Sqr_Cel(k).X + (0 * WHthird), y:=Sqr_Cel(k).Y + (2 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 7
        Sqr_Cdd((k * 10) + 8) = New Rectangle(x:=Sqr_Cel(k).X + (1 * WHthird), y:=Sqr_Cel(k).Y + (2 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 8
        Sqr_Cdd((k * 10) + 9) = New Rectangle(x:=Sqr_Cel(k).X + (2 * WHthird), y:=Sqr_Cel(k).Y + (2 * WHthird), width:=WHthird, height:=WHthird)     ' Cdd 9
        k += 1
      Next i
    Next j

    ' Calcul des Squares Path à Coins Arrondis et carrés
    k = 0
    For j As Integer = 0 To 8      ' Axe des y, vertical
      For i As Integer = 0 To 8    ' Axe des x, horizontal
        Dim x, y As Integer
        x = Sqr_Cel(k).X
        y = Sqr_Cel(k).Y
        Sqr_Pth(k) = New GraphicsPath
        Select Case k
          Case 0, 3, 6, 27, 30, 33, 54, 57, 60       ' Coin HG arrondi
            With Sqr_Pth(k)
              .StartFigure()
              Select Case Plcy_Format_DAB
                Case 1, 3, 5
                  .AddArc(New Rectangle(x, y, WHhalf, WHhalf), 180, 90)
                  .AddLine(x + WHhalf, y, x + WH, y)
                  .AddLine(x + WH, y, x + WH, y + WH)
                  .AddLine(x + WH, y + WH, x, y + WH)
                  .AddLine(x, y + WH, x, y + WHhalf)
                Case 2, 4, 6
                  .AddLine(x, y + WHquart, x + WHquart, y)
                  .AddLine(x + WHquart, y, x + WH, y)
                  .AddLine(x + WH, y, x + WH, y + WH)
                  .AddLine(x + WH, y + WH, x, y + WH)
                  .AddLine(x, y + WH, x, y + WHquart)
              End Select
              .CloseFigure() 'La figure est déjà fermée
            End With
          Case 2, 5, 8, 29, 32, 35, 56, 59, 62       ' Coin HD arrondi
            With Sqr_Pth(k)
              .StartFigure()
              Select Case Plcy_Format_DAB
                Case 1, 3, 5
                  .AddLine(x, y, x + WH - WHhalf, y)
                  .AddArc(New Rectangle(x + WH - WHhalf, y, WHhalf, WHhalf), 270, 90)
                  .AddLine(x + WH, y + WHhalf, x + WH, y + WH)
                  .AddLine(x + WH, y + WH, x, y + WH)
                  .AddLine(x, y + WH, x, y)
                Case 2, 4, 6
                  .AddLine(x, y, x + WH - WHquart, y)
                  .AddLine(x + WH - WHquart, y, x + WH, y + WHquart)
                  .AddLine(x + WH, y + WHhalf, x + WH, y + WH)
                  .AddLine(x + WH, y + WH, x, y + WH)
                  .AddLine(x, y + WH, x, y)
              End Select
              .CloseFigure()
            End With
          Case 18, 21, 24, 45, 48, 51, 72, 75, 78    ' Coin BG arrondi
            With Sqr_Pth(k)
              .StartFigure()
              Select Case Plcy_Format_DAB
                Case 1, 3, 5
                  .AddLine(x, y, x + WH, y)
                  .AddLine(x + WH, y, x + WH, y + WH)
                  .AddLine(x + WH, y + WH, x + WHhalf, y + WH)
                  .AddArc(New Rectangle(x, y + WH - WHhalf, WHhalf, WHhalf), 90, 90)
                  .AddLine(x, y + WH - WHhalf, x, y)
                Case 2, 4, 6
                  .AddLine(x, y, x + WH, y)
                  .AddLine(x + WH, y, x + WH, y + WH)
                  .AddLine(x + WH, y + WH, x + WHquart, y + WH)
                  .AddLine(x + WHquart, y + WH, x, y + WH - WHquart)
                  .AddLine(x, y + WH - WHquart, x, y)
              End Select
              .CloseFigure()
            End With
          Case 20, 23, 26, 47, 50, 53, 74, 77, 80    ' Coin BD arrondi
            With Sqr_Pth(k)
              .StartFigure()
              Select Case Plcy_Format_DAB
                Case 1, 3, 5
                  .AddLine(x, y, x + WH, y)
                  .AddLine(x + WH, y, x + WH, y + WH - WHhalf)
                  .AddArc(New Rectangle(x + WH - WHhalf, y + WH - WHhalf, WHhalf, WHhalf), 0, 90)
                  .AddLine(x + WH - WHhalf, y + WH, x, y + WH)
                  .AddLine(x, y + WH, x, y)
                Case 2, 4, 6
                  .AddLine(x, y, x + WH, y)
                  .AddLine(x + WH, y, x + WH, y + WH - WHquart)
                  .AddLine(x + WH, y + WH - WHquart, x + WH - WHquart, y + WH)
                  .AddLine(x + WH - WHquart, y + WH, x, y + WH)
                  .AddLine(x, y + WH, x, y)
              End Select
              .CloseFigure()
            End With
          Case Else                   ' Carré
            With Sqr_Pth(k)
              .StartFigure()
              .AddLine(x, y, x + WH, y)
              .AddLine(x + WH, y, x + WH, y + WH)
              .AddLine(x + WH, y + WH, x, y + WH)
              .AddLine(x, y + WH, x, y)
              .CloseFigure() 'La figure est déjà fermée
            End With
        End Select
        k += 1
      Next i
    Next j

    ' Construction des 10 images correspondantes aux polices substituées
    ' Ils sont calculés quelque soit Plcy_Fantasy et indépendamment de WH
    Using font As New Font(Font_Name_ValCdd, 12),
          brsh As New SolidBrush(Color.Black)          'La couleur Color.Black pour les menus contextuels et les filtres de la BO
      For i As Integer = 0 To 9
        Dim Bm_MnuBO As New Bitmap(22, 22)
        Using g As Graphics = Graphics.FromImage(image:=bm_MnuBO)
          g.DrawString(Subst_Police(CStr(i)), font, brsh, 10, 10, Format_Center)
          Sqr_Fantasy(i) = Bm_MnuBO
        End Using
      Next i
    End Using
    Obj_Colors_Load()
  End Sub
  Public Sub OC_U_Pt20_Init()
    '                                                                       0        4   8  12      1
    '                                                                           A0  A1  A2  A3  B0   
    'Retourne un tableau de 20 points A0, A1, A2, A3  pour le côté Haut
    '                                                                      15   D3  E0      F0  B1  5
    '                                 B0, B1, B2, B3  pour le côté droit
    '                                                                      11   D2              B2  9
    '                                 C0, C1, C2, C3  pour le bas
    '                                                                       7   D1  H0      G0  B3  13
    '                              et D0, D1, D2, D3  pour le côté gauche
    '                                                                           D0  C3  C2  C1  C0   
    '                                                                       3       14  10   6      2
    '                             
    'Le tableau est calculé dans OC_Grid_Compute_Squares
    '           est utilisé dans G0_Cell_Figure
    For Cellule As Integer = 0 To 80
      Dim Pt_Origine As New PointF(Sqr_Cel(Cellule).X, Sqr_Cel(Cellule).Y)
      Dim Pt_Left As Single = Pt_Origine.X + 1
      Dim Pt_Top As Single = Pt_Origine.Y + 1
      Dim Pt_Right As Single = Sqr_Cel(Cellule).Right - 3
      Dim Pt_Bottom As Single = Sqr_Cel(Cellule).Bottom - 3
      Dim PA0, PB0, PC0, PD0, PA1, PB1, PC1, PD1, PA2, PB2, PC2, PD2, PA3, PB3, PC3, PD3 As PointF
      Dim PE0, PF0, PG0, PH0 As PointF

      Dim D1_4 As Single = ((1 * WH) \ 4)
      Dim D2_4 As Single = 2 * D1_4
      Dim D3_4 As Single = 3 * D1_4

      PA0.X = Pt_Left
      PA0.Y = Pt_Top
      PB0.X = Pt_Right
      PB0.Y = Pt_Top
      PC0.X = Pt_Right
      PC0.Y = Pt_Bottom
      PD0.X = Pt_Left
      PD0.Y = Pt_Bottom

      PA1.X = PA0.X + D1_4 : PA1.Y = PA0.Y
      PA2.X = PA0.X + D2_4 : PA2.Y = PA0.Y
      PA3.X = PA0.X + D3_4 : PA3.Y = PA0.Y

      PB1.X = PB0.X : PB1.Y = PB0.Y + D1_4
      PB2.X = PB0.X : PB2.Y = PB0.Y + D2_4
      PB3.X = PB0.X : PB3.Y = PB0.Y + D3_4

      PC1.X = PD0.X + D3_4 : PC1.Y = PC0.Y
      PC2.X = PD0.X + D2_4 : PC2.Y = PC0.Y
      PC3.X = PD0.X + D1_4 : PC3.Y = PC0.Y

      PD1.X = PD0.X : PD1.Y = PA0.Y + D3_4
      PD2.X = PD0.X : PD2.Y = PA0.Y + D2_4
      PD3.X = PD0.X : PD3.Y = PA0.Y + D1_4

      PE0.X = PA0.X + D1_4
      PE0.Y = PA0.Y + D1_4
      PF0.X = PB0.X - D1_4
      PF0.Y = PB0.Y + D1_4
      PG0.X = PC0.X - D1_4
      PG0.Y = PC0.Y - D1_4
      PH0.X = PD0.X + D1_4
      PH0.Y = PD0.Y - D1_4

      '                 0    1    2    3    4    5    6    7    8    9    10   11   12   13   14   15
      'G0_Cell_Pt_16 = {PA0, PB0, PC0, PD0, PA1, PB1, PC1, PD1, PA2, PB2, PC2, PD2, PA3, PB3, PC3, PD3}
      '                 16   17   18   19 
      'G0_Cell_Pt_16 = {PE0, PF0, PG0, PH0}
      U_Pt20(Cellule, 0) = PA0
      U_Pt20(Cellule, 1) = PB0
      U_Pt20(Cellule, 2) = PC0
      U_Pt20(Cellule, 3) = PD0
      U_Pt20(Cellule, 4) = PA1
      U_Pt20(Cellule, 5) = PB1
      U_Pt20(Cellule, 6) = PC1
      U_Pt20(Cellule, 7) = PD1
      U_Pt20(Cellule, 8) = PA2
      U_Pt20(Cellule, 9) = PB2
      U_Pt20(Cellule, 10) = PC2
      U_Pt20(Cellule, 11) = PD2
      U_Pt20(Cellule, 12) = PA3
      U_Pt20(Cellule, 13) = PB3
      U_Pt20(Cellule, 14) = PC3
      U_Pt20(Cellule, 15) = PD3
      U_Pt20(Cellule, 16) = PE0
      U_Pt20(Cellule, 17) = PF0
      U_Pt20(Cellule, 18) = PG0
      U_Pt20(Cellule, 19) = PH0
    Next Cellule
  End Sub

  Public Sub OC_Grid_Compute_Font_Size_IA()
    'Calcul de la taille des polices des Valeurs et des Candidats
    Font_Val_Size = CalculateFontSize_IA(Font_Name_ValCdd, FontStyle.Regular, WH, "1")
    Font_Cdd_Size_Zoom = (4 * Font_Val_Size) / 6
    Font_Cdd_Size = CalculateFontSize_IA(Font_Name_ValCdd, FontStyle.Italic, CSng(WH / 3), "1")
  End Sub
  Private Function CalculateFontSize_IA(fontName As String,
                                      fontStyle As FontStyle,
                                      maxHeight As Single,
                                      text As String) As Single
    Dim size As Single
    Using g As Graphics = Graphics.FromHwnd(IntPtr.Zero)
      For p As Integer = 5 To 100
        Using font As New Font(fontName, p, fontStyle, GraphicsUnit.Pixel)
          If g.MeasureString(text, font).Height < maxHeight Then
            size = CSng(p / 2)
          Else
            Exit For
          End If
        End Using
      Next
    End Using
    Return size
  End Function

  Public Sub OC_Grid_Cutting_Image()
    If Plcy_Fond_Grille = 0 Then Exit Sub

    Dim répertoire As String = Path_SDK & "S10_Icônes\Fonds\"
    Dim fond_Files As IEnumerable(Of String) =
        From file In IO.Directory.GetFiles(répertoire)
        Where file.EndsWith("jpg", StringComparison.OrdinalIgnoreCase)
        Order By file Descending

    ' --- 1) Charger l'image source ---
    Using fond_BM As New Bitmap(fond_Files(Plcy_Fond_Grille - 1))

      ' --- 2) Créer une version carrée ---
      Dim côté As Integer = Math.Min(fond_BM.Height, fond_BM.Width)
      Using fond_BM_Carré As New Bitmap(côté, côté)
        Using g_Carré As Graphics = Graphics.FromImage(fond_BM_Carré)
          Dim rct_Carré As New Rectangle(0, 0, côté, côté)
          g_Carré.DrawImage(fond_BM, rct_Carré, 0, 0, côté, côté, GraphicsUnit.Pixel)
        End Using

        ' --- 3) Redimensionner à la taille de la grille ---
        Using fond_BM_WH_Grid As New Bitmap(Bld_WH_Grid, Bld_WH_Grid)
          Using gI As Graphics = Graphics.FromImage(fond_BM_WH_Grid)
            gI.DrawImage(fond_BM_Carré, 0, 0, Bld_WH_Grid, Bld_WH_Grid)
          End Using

          ' --- 4) Appliquer la transparence pixel par pixel ---
          'For x As Integer = 0 To fond_BM_WH_Grid.Width - 1
          '  For y As Integer = 0 To fond_BM_WH_Grid.Height - 1
          '    Dim px As Color = fond_BM_WH_Grid.GetPixel(x, y)
          '    Dim pxA As Color = Color.FromArgb(128, px.R, px.G, px.B)
          '    fond_BM_WH_Grid.SetPixel(x, y, pxA)
          '  Next
          'Next
          ApplyTransparencyFast(fond_BM_WH_Grid, 128)

          ' --- 5) Découper en 81 morceaux ---
          For i As Integer = 0 To 80
            Dim bm_wh As New Bitmap(WH, WH) ' NE PAS mettre dans un Using
            Using gI_wh As Graphics = Graphics.FromImage(bm_wh)
              Dim rct_81 As New Rectangle(0, 0, WH, WH)
              gI_wh.DrawImage(fond_BM_WH_Grid,
                              rct_81,
                              Sqr_Cel(i).X - Sqr_Cel(0).X,
                              Sqr_Cel(i).Y - Sqr_Cel(0).Y,
                              WH, WH,
                              GraphicsUnit.Pixel)
            End Using

            ' On stocke l'image → elle sera disposée plus tard
            Sqr_Img(i) = bm_wh
          Next

        End Using ' fond_BM_WH_Grid
      End Using ' fond_BM_Carré
    End Using ' fond_BM
  End Sub

  Private Sub ApplyTransparencyFast(bmp As Bitmap, alpha As Byte)

    Dim rect As New Rectangle(0, 0, bmp.Width, bmp.Height)
    Dim data As Imaging.BitmapData =
        bmp.LockBits(rect, Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format32bppArgb)

    Dim ptr As IntPtr = data.Scan0
    Dim bytes As Integer = Math.Abs(data.Stride) * bmp.Height
    Dim rgbValues(bytes - 1) As Byte

    ' Copier la mémoire dans le tableau
    Runtime.InteropServices.Marshal.Copy(ptr, rgbValues, 0, bytes)

    ' Boucle rapide : 4 octets par pixel (B, G, R, A)
    For i As Integer = 0 To rgbValues.Length - 1 Step 4
      rgbValues(i + 3) = alpha   ' canal A
    Next

    ' Recopier dans le bitmap
    Runtime.InteropServices.Marshal.Copy(rgbValues, 0, ptr, bytes)

    bmp.UnlockBits(data)

  End Sub

  'Public Sub OC_Grid_Cutting_Image()
  '  'Découpage de l'image en 81 morceaux 
  '  If Plcy_Fond_Grille = 0 Then Exit Sub 'Il n'y a pas d'affichage de fond

  '  'Traitement dupliqué dans Frm/Préférences/Préférences_Load
  '  Dim Répertoire As String = Path_SDK & "S10_Icônes\Fonds\"
  '  Dim Fond_Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(Répertoire)
  '                                             Where File.Contains("JPG") Or File.Contains("jpg")
  '                                             Order By File Descending

  '  'Fond_BM est rectangulaire (Portrait ou paysage) ou carrée
  '  Dim Fond_BM As New Bitmap(filename:=Fond_Files(Plcy_Fond_Grille - 1))
  '  'ou         As Image = Image.FromFile(filename:=Fond_Files(Plcy_Fond_Grille - 1))

  '  'La photo est mise au format carré, à/p du Top-Left. Elle peut donc être mal cadrée.
  '  'Il n'est pas tenu compte du format paysage/portrait. Toujours en paysage
  '  Dim Côté As Integer = Math.Min(Fond_BM.Height, Fond_BM.Width)
  '  Dim Fond_BM_Carré As New Bitmap(Côté, Côté)
  '  Using g_Carré As Graphics = Graphics.FromImage(image:=Fond_BM_Carré)
  '    Dim Rct_Carré As New Rectangle(0, 0, Côté, Côté)
  '    g_Carré.DrawImage(image:=Fond_BM,
  '                      destRect:=Rct_Carré,
  '                      srcX:=0, srcY:=0, srcWidth:=Côté, srcHeight:=Côté, srcUnit:=GraphicsUnit.Pixel)
  '  End Using
  '  'Création d'une image dimensionnée à la grille
  '  Dim Fond_BM_WH_Grid As New Bitmap(Bld_WH_Grid, Bld_WH_Grid)
  '  Using gI As Graphics = Graphics.FromImage(image:=Fond_BM_WH_Grid)
  '    gI.DrawImage(image:=Fond_BM_Carré,
  '                 x:=0, y:=0, width:=Bld_WH_Grid, height:=Bld_WH_Grid)
  '  End Using

  '  ' Boucle de transparence sur le pixel A
  '  ' Comme la photo est transparente, il est plus facile de voir les chiffres et les stratégies
  '  '                                  AVANT le paint, le fond d'effacement est paint
  '  Dim Pxl_before, Pxl_after As Color
  '  For x As Integer = 0 To Fond_BM_WH_Grid.Width - 1
  '    For y As Integer = 0 To Fond_BM_WH_Grid.Height - 1
  '      Pxl_before = Fond_BM_WH_Grid.GetPixel(x, y)
  '      Pxl_after = Color.FromArgb(128, Pxl_before.R, Pxl_before.G, Pxl_before.B)
  '      Fond_BM_WH_Grid.SetPixel(x, y, Pxl_after)
  '    Next y
  '  Next x

  '  For i As Integer = 0 To 80
  '    'Découpage de l'image dimensionnée en 81 images BM_wh
  '    Dim BM_wh As New Bitmap(WH, WH)
  '    Using gI_wh As Graphics = Graphics.FromImage(image:=BM_wh)
  '      Dim Rct_81 As New Rectangle(0, 0, WH, WH)
  '      gI_wh.DrawImage(image:=Fond_BM_WH_Grid,
  '                      destRect:=Rct_81,
  '                      srcX:=Sqr_Cel(i).X - Sqr_Cel(0).X, srcY:=Sqr_Cel(i).Y - Sqr_Cel(0).Y,
  '                      srcWidth:=WH, srcHeight:=WH,
  '                      srcUnit:=GraphicsUnit.Pixel)
  '      Sqr_Img(i) = BM_wh
  '    End Using
  '  Next i
  'End Sub
#End Region
End Module