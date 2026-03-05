Option Strict On
Option Explicit On

Imports System.Drawing.Drawing2D

'RADICAL: Gx_

Friend Module M03_Paint
  '-------------------------------------------------------------------------------
  'Rappel: L'axe HORIZONTAL est l'axe des x
  '        L'axe VERTICAL   est l'axe des y
  '        Le point O est situé en Top-Left
  'Nommage:
  ' Paint
  '     N° de Couche
  '         Grid ou Cellule
  '            Texte plus explicite
  ' 20/09/2022 L'ensemble des dessins sont faits à l'intérieur de Sqr_Cel
  '            avec Top-Left + 1 et Width-Height - 3
  '
  '-------------------------------------------------------------------------------   

  Public Sub U_Strg_Effacer_g(g As Graphics)
    ' TODO Vraisemblablement, tous ces traitements seront supprimés
    'Toutes les cellules concernées par une stratégie ont un rafraîchissement du fond et de la Valeur/Candidat
    Dim sc As New Cellule_Cls
    For i As Integer = 0 To 80
      If U_Strg(i) Then
        sc.Numéro = i
        sc.G2_Cellule_Paint_Fond_g(g)
        sc.G5_Cellule_Paint_Valeur_g(g)
        sc.G6_Cellule_Paint_Candidats_Conditions_Nrm_Cdd_g(g) ' dans le cas où seul un candidat a été exclu
      End If
      U_Strg(i) = False
    Next i
  End Sub

#Region "G4 Couche Stratégie"
  '   La couche G4 stratégies n'est appelée que dans G4_Grid_Stratégie_All,
  Public Sub G4_Grid_Stratégie_All_g(g As Graphics)
    If Plcy_Gnrl = "Nrm" And Plcy_Strg <> "   " Then

      For i As Integer = 0 To 80
        U_Strg_Val_Ins(i) = ""
        U_Strg_Cdd_Exc(i) = Cnddts_Blancs
      Next i
      G4_Grid_Stratégie_Cdd_g(g)
      G4_Grid_Stratégie_CdU_g(g)
      G4_Grid_Stratégie_CdO_g(g)
      G4_Grid_Stratégie_Flt_g(g)
      G4_Grid_Stratégie_CdS_g(g)
      G4_Grid_Stratégie_DCd_g(g)
      G4_Grid_Stratégie_Cbl_g(g)
      G4_Grid_Stratégie_Tpl_g(g)
      G4_Grid_Stratégie_Xwg_g(g)
      G4_Grid_Stratégie_XYw_g(g)
      G4_Grid_Stratégie_Swf_g(g)
      G4_Grid_Stratégie_Jly_g(g)
      G4_Grid_Stratégie_XYZ_g(g)
      G4_Grid_Stratégie_SKy_g(g)
      G4_Grid_Stratégie_Unq_g(g)
      G4_Grid_Stratégie_Obj_g(g)
      G4_Grid_Stratégie_XCx_XCy_XNl_g(g)
      G4_Grid_Stratégie_XRp_g(g)
      G4_Grid_Stratégie_WgX_WgY_WgZ_WgW_g(g)
      G4_Grid_Stratégie_GLk_g(g)
      G4_Grid_Stratégie_Gbl_g(g)
      G4_Grid_Stratégie_Gbv_g(g)
      G4_Grid_Stratégie_GCs_g(g)
    End If
  End Sub

  '-------------------------------------------------------------------------------
  'Les stratégies
  '    Les stratégies sont présentées en deux niveaux
  '        1 un "Double carré" est affiché sur les cellules concernées par la stratégie
  '        2 Aide Graphique cochée
  '          - Signalisation des axes
  '          - Les Candidats sont affichés
  '          - Signalisation en jaune du candidat à placer  + option jaune du menu contextuel
  '          -               en rouge du candidat à enlever + option rouge du menu contextuel
  'Les cellules en stratégie sont notées dans U_Strg(80) as Boolean
  '-------------------------------------------------------------------------------   
  Public Sub G4_Grid_Stratégie_Cdd_g(g As Graphics)
    If Not Plcy_Strg = "Cdd" Then Exit Sub
    For i As Integer = 0 To 80
      If U(i, 2) = " " Then
        Dim sc_a As New Cellule_Cls With {.Numéro = i}
        sc_a.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
      End If
    Next i
  End Sub
  Public Sub G4_Grid_Stratégie_CdU_g(g As Graphics)
    ' La stratégie CdU calcule TOUS les CdU, UN SEUL CdU au hasard est présenté 
    Dim Cellule As Integer
    Dim Candidat As String
    If Not Plcy_Strg = "CdU" Then Exit Sub

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Cellule = RRslt.Cellule(0)
    Candidat = RRslt.Candidat
    U_Strg_Val_Ins(Cellule) = Candidat
    G0_Cell_Figure_g(g, Cellule, "Double_Carré", Color_Stratégique)

    U_MdC_Init()
    G4_MdC_Row_Col_Box("Row", U_Row(Cellule))
    G4_MdC_Row_Col_Box("Col", U_Col(Cellule))
    G4_MdC_Row_Col_Box("Box", U_Reg(Cellule))
    G4_MdC_Paint_g(g) ' Les figures sont dessinées et les candidats affichés
    'Re-dessine le candidat à placer dans un cercle plein Jaune
    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    sc.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Insérer)
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & Candidat & " jaune à placer."
  End Sub
  Public Sub G4_Grid_Stratégie_CdO_g(g As Graphics)
    Dim Cellule As Integer
    Dim Candidat As String
    If Not Plcy_Strg = "CdO" Then Exit Sub
    Try
      If RRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If
      Cellule = RRslt.Cellule(0)
      Candidat = RRslt.Candidat
      U_Strg_Val_Ins(Cellule) = Candidat
      G0_Cell_Figure_g(g, Cellule, "Double_Carré", Color_Stratégique)

      Dim Code_LCR As String = RRslt.Code_LCR
      Dim LCR As Integer = RRslt.LCR
      U_MdC_Init()
        Select Case Code_LCR
          Case "L" : G4_MdC_Row_Col_Box("Row", LCR)
          Case "C" : G4_MdC_Row_Col_Box("Col", LCR)
          Case "R" : G4_MdC_Row_Col_Box("Box", LCR)
        End Select
      G4_MdC_Paint_g(g) ' Les figures sont dessinées et les candidats affichés
      'Re-dessine le candidat à placer dans un cercle plein Jaune
      Dim sc As New Cellule_Cls With {.Numéro = Cellule}
      sc.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Insérer)
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & Candidat & " jaune à placer."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub G4_Grid_Stratégie_Flt_g(g As Graphics)
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte
    ' Affichage des Valeurs Filtrés
    If Mid$(Plcy_Strg, 1, 2) = "FV" Then
      Dim Valeur_Filtrée As String = Mid$(Plcy_Strg, 3, 1)
      For i As Integer = 0 To 80
        If U(i, 2) = Valeur_Filtrée Then
          G0_Cell_Figure_g(g, i, "Double_Carré", Color_Stratégique)
        End If
      Next i
      MW_Prv_Val = CInt(Valeur_Filtrée)
    End If

    ' Affichage des Candidats Filtrés
    If Mid$(Plcy_Strg, 1, 2) = "FC" Then
      Dim Candidat As String = Mid$(Plcy_Strg, 3, 1)
      Dim Color As Color = Color_BySymbol(Obj_Symbol)
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(Candidat) Then sc.G6_Cellule_Paint_Candidat_g(g, Candidat, Color)
      Next i
    End If
  End Sub
  Public Sub G4_Grid_Stratégie_CdS_g(g As Graphics)
    ' La stratégie du Candidat Saisi CdS est activée dans Cell_Val_Insert
    '              qui documente Pbl_Valeur_CdS = V
    '              qui exécute un Invalidate, donc un OnPaint
    If Not Plcy_Strg = "CdS" Then Exit Sub
    If Pbl_Valeur_CdS = "" Then Exit Sub

    ' 1 Aide Simple uniquement
    For i As Integer = 0 To 80
      If U(i, 2) = Pbl_Valeur_CdS Then
        G0_Cell_Figure_g(g, i, "Double_Carré", Color_Stratégique)
      End If
    Next i
  End Sub
  Public Sub G4_Grid_Stratégie_DCd_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Cellule As Integer
    If Not Plcy_Strg = "DCd" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_DCd(U_temp)

      If DCdd_List.Count = 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      'Public U_Strg_Val_Ins(80) As String comporte pour chaque poste U la valeur à insérer 
      ' 1 Aide Simple
      For Each DCdd As DCdd_Cls In DCdd_List
        Cellule = DCdd.Cellule
        U_Strg_Val_Ins(Cellule) = DCdd.Candidat
        G0_Cell_Figure_g(g, Cellule, "Double_Carré", Color_Stratégique)
        U_Strg(Cellule) = True
      Next DCdd
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      If DCdd_List_Exists(Pbl_Cell_Select) Then
        Dim DCdd2 As DCdd_Cls = DCdd_Get(Pbl_Cell_Select)
        U_MdC_Init()
        Select Case DCdd2.Sous_Stratégie
          Case "Bh0" : G4_MdC_Trait_ou_Rectangle(0, 26)
          Case "Bh1" : G4_MdC_Trait_ou_Rectangle(27, 53)
          Case "Bh2" : G4_MdC_Trait_ou_Rectangle(54, 80)
          Case "Bv0" : G4_MdC_Trait_ou_Rectangle(0, 74)
          Case "Bv1" : G4_MdC_Trait_ou_Rectangle(3, 77)
          Case "Bv2" : G4_MdC_Trait_ou_Rectangle(6, 80)
          Case "CdU"
            G4_MdC_Row_Col_Box("Row", U_Row(Pbl_Cell_Select))
            G4_MdC_Row_Col_Box("Col", U_Col(Pbl_Cell_Select))
            G4_MdC_Row_Col_Box("Box", U_Reg(Pbl_Cell_Select))

          Case "CdO_L" : G4_MdC_Row_Col_Box("Row", U_Row(Pbl_Cell_Select))
          Case "CdO_C" : G4_MdC_Row_Col_Box("Col", U_Col(Pbl_Cell_Select))
          Case "CdO_R" : G4_MdC_Row_Col_Box("Box", U_Reg(Pbl_Cell_Select))

          Case Else
            Jrn_Add(, {"DCd Sous_Stratégie inconnue : " & DCdd2.Sous_Stratégie})
        End Select
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g) ' Les figures sont dessinées et les candidats affichés

        '  'Re-dessine le candidat à placer dans un cercle plein Jaune
        Dim sc As New Cellule_Cls With {.Numéro = Pbl_Cell_Select}
        sc.G6_Cellule_Paint_Candidat_g(g, DCdd2.Candidat, Color_Cdd_Insérer)
        U_Strg(Cellule) = True
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & DCdd2.Sous_Stratégie & " " & DCdd2.Candidat & " jaune à placer."
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub G4_Grid_Stratégie_DCd_g_save(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Cellule As Integer
    If Not Plcy_Strg = "DCd" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_DCd(U_temp)

      If DCdd_List.Count = 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      'Public U_Strg_Val_Ins(80) As String comporte pour chaque poste U la valeur à insérer 
      ' 1 Aide Simple
      For Each DCdd As DCdd_Cls In DCdd_List
        Cellule = DCdd.Cellule
        U_Strg_Val_Ins(Cellule) = DCdd.Candidat
        G0_Cell_Figure_g(g, Cellule, "Double_Carré", Color_Stratégique)
        U_Strg(Cellule) = True
      Next DCdd
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      If DCdd_List_Exists(Pbl_Cell_Select) Then
        Dim DCdd2 As DCdd_Cls = DCdd_Get(Pbl_Cell_Select)
        U_MdC_Init()
        Select Case DCdd2.Sous_Stratégie
          Case "Bh0" : G4_MdC_Trait_ou_Rectangle(0, 26)
          Case "Bh1" : G4_MdC_Trait_ou_Rectangle(27, 53)
          Case "Bh2" : G4_MdC_Trait_ou_Rectangle(54, 80)
          Case "Bv0" : G4_MdC_Trait_ou_Rectangle(0, 74)
          Case "Bv1" : G4_MdC_Trait_ou_Rectangle(3, 77)
          Case "Bv2" : G4_MdC_Trait_ou_Rectangle(6, 80)
          Case "CdU"
            G4_MdC_Row_Col_Box("Row", U_Row(Pbl_Cell_Select))
            G4_MdC_Row_Col_Box("Col", U_Col(Pbl_Cell_Select))
            G4_MdC_Row_Col_Box("Box", U_Reg(Pbl_Cell_Select))

          Case "CdO_L" : G4_MdC_Row_Col_Box("Row", U_Row(Pbl_Cell_Select))
          Case "CdO_C" : G4_MdC_Row_Col_Box("Col", U_Col(Pbl_Cell_Select))
          Case "CdO_R" : G4_MdC_Row_Col_Box("Box", U_Reg(Pbl_Cell_Select))

          Case Else
            Jrn_Add(, {"DCd Sous_Stratégie inconnue : " & DCdd2.Sous_Stratégie})
        End Select
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g) ' Les figures sont dessinées et les candidats affichés

        '  'Re-dessine le candidat à placer dans un cercle plein Jaune
        Dim sc As New Cellule_Cls With {.Numéro = Pbl_Cell_Select}
        sc.G6_Cellule_Paint_Candidat_g(g, DCdd2.Candidat, Color_Cdd_Insérer)
        U_Strg(Cellule) = True
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & DCdd2.Sous_Stratégie & " " & DCdd2.Candidat & " jaune à placer."
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub Strategy_BTXYSJZKQ_Aide_Simple_g(g As Graphics, Strategy_Rslt(,) As String, U_Strg() As Boolean)
    Dim Cellule As Integer
    Dim Candidat As String
    Dim Candidats As String
    ' 1 Aide Simple
    'U_Strg_Effacer()

    For i As Integer = 1 To UBound(Strategy_Rslt, 2)
      For k As Integer = 10 To 54
        If Strategy_Rslt(k, i) = "__" Then Exit For
        Cellule = CInt(Strategy_Rslt(k, i))
        U_Strg(Cellule) = True
      Next k
      For k As Integer = 55 To 99
        If Strategy_Rslt(k, i) = "__" Then Exit For
        Cellule = CInt(Strategy_Rslt(k, i))
        Candidat = Strategy_Rslt(5, i)
        Candidats = U_Strg_Cdd_Exc(Cellule)
        Select Case Strategy_Rslt(1, i)
          Case "Cbl", "Tpl", "Xwg", "XYw", "Swf", "Jly", "XYZ", "SKy"
            Mid$(Candidats, CInt(Candidat), 1) = Candidat ' le candidat est ajouté
          Case "Unq"
            ' Unq est la seule stratégie qui propose plusieurs candidats à exclure
            'On ajoute à U_Strg_Cdd_Exc(Cellule) le ou les candidats de la stratégie Unq
            For c As Integer = 1 To 9
              If Mid$(Candidat, c, 1) = CStr(c) Then
                Mid$(Candidats, c, 1) = CStr(c) ' le candidat est ajouté
              End If
            Next c
        End Select
        U_Strg_Cdd_Exc(Cellule) = Candidats
      Next k
    Next i '/For i = 1 To UBound(Strategy_Rslt, 2)

    For i As Integer = 0 To 80
      If U_Strg(i) Then G0_Cell_Figure_g(g, i, "Double_Carré", Color_Stratégique)
    Next i

  End Sub
  Public Sub G4_Grid_Stratégie_Cbl_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String

    If Not Plcy_Strg = "Cbl" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_Cbl(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        Select Case Strategy_Rslt(2, Ligne).Substring(3, 1)
          Case "C" : G4_MdC_Row_Col_Box("Col", U_Col(Pbl_Cell_Select))
          Case "L" : G4_MdC_Row_Col_Box("Row", U_Row(Pbl_Cell_Select))
        End Select
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés

        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ":  " & Candidat & " rouge à enlever."
      End If

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub
  Public Sub G4_Grid_Stratégie_Tpl_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String

    If Not Plcy_Strg = "Tpl" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_Tpl(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        If Strategy_Rslt(3, Ligne) = "C" Then G4_MdC_Row_Col_Box("Col", U_Col(Pbl_Cell_Select))
        If Strategy_Rslt(3, Ligne) = "L" Then G4_MdC_Row_Col_Box("Row", U_Row(Pbl_Cell_Select))
        If Strategy_Rslt(3, Ligne) = "R" Then G4_MdC_Row_Col_Box("Box", U_Reg(Pbl_Cell_Select))
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés

        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ": " & Candidat & " rouge à enlever."
      End If

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub
  Public Sub G4_Grid_Stratégie_Xwg_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String
    ' G4_Grid_Strategy_Xwg diffère des autres G4_Grid_Strategy_abc sur les points suivants:
    '   Elle gère les sous-stratégies Xwg, Fnd
    '   Aide simple identique aux autres stratégies
    '   Aide Graphique
    '     Affichage de toutes les cellules avec le SEUL candidat concerné en disque
    '     Utilisation de Cel(0 To 44) As Integer pour simplifier l'affichage des cellules
    '     sous-stratégie Xwg_C Xwg_L affichage de trait et de rectangle
    '     sous-stratégie Fnd_C Fnd_L affichage de trait et d'une diagonale 
    '                    les candidats concernés sont en carré

    If Not Plcy_Strg = "Xwg" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_Xwg(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        ' Affichage exceptionnel d'un disque dans les cellules avec le SEUL candidat concerné
        For i As Integer = 0 To 80
          If U_temp(i, 3).Contains(Candidat) Then
            U_Strg(i) = True
            Dim sc_a As New Cellule_Cls With {.Numéro = i}
            sc_a.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Stratégique)
          End If
        Next i

        Dim Sous_Stratégie As String = Strategy_Rslt(2, Ligne)
        Dim Cel(0 To 44) As Integer : For k As Integer = 0 To 44 : Cel(k) = -1 : Next k
        For k As Integer = 10 To 54
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cel(k - 10) = CInt(Strategy_Rslt(k, Ligne))
        Next k
        Select Case Sous_Stratégie
           'U_Strg est documenté dans G4_MdC_Row_Col_Box
           '                          G0_Cell_Diagonale
          Case "Xwg_C"
            G4_MdC_Row_Col_Box("Row", U_Row(Cel(0)))
            G4_MdC_Row_Col_Box("Row", U_Row(Cel(2)))
            G4_MdC_Trait_ou_Rectangle(Cel(0), Cel(3))
          Case "Xwg_L"
            G4_MdC_Row_Col_Box("Col", U_Col(Cel(0)))
            G4_MdC_Row_Col_Box("Col", U_Col(Cel(2)))
            G4_MdC_Trait_ou_Rectangle(Cel(0), Cel(3))
          Case "Fnd_La", "Fnd_LbL", "Fnd_LbR", "Fnd_Lc", "Fnd_Ca", "Fnd_CbT", "Fnd_CbB", "Fnd_Cc"
            G4_MdC_Trait_ou_Rectangle(Cel(0), Cel(1))
            G4_MdC_Trait_ou_Rectangle(Cel(2), Cel(4))
            ' Un Disque est déjà affiché, il est complété pour les 2 branches par un Carré
            G0_Cdd_Figure_g(g, Cel(0), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(1), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(2), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(3), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(4), CInt(Candidat), "Carré", Color_Stratégique)
            'L'aileron est présenté par une diagonale de Bresenham
            G0_Cell_Diagonale_g(g, Cel(5), Cel(6))
          Case "Shm_L", "Shm_C"
            G4_MdC_Trait_ou_Rectangle(Cel(0), Cel(1))
            G4_MdC_Trait_ou_Rectangle(Cel(2), Cel(4))
            G0_Cdd_Figure_g(g, Cel(0), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(1), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(2), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(3), CInt(Candidat), "Carré", Color_Stratégique)
            G0_Cdd_Figure_g(g, Cel(4), CInt(Candidat), "Carré", Color_Stratégique)
        End Select
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés
        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Select Case Sous_Stratégie
          Case "Xwg_C", "Xwg_L"
            Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ": " & Candidat & " rouge à enlever."
          Case "Fnd_La", "Fnd_LbL", "Fnd_LbR", "Fnd_Lc", "Fnd_Ca", "Fnd_CbT", "Fnd_CbB", "Fnd_Cc"
            Frm_SDK.B_Info.Text = "Stratégie Finned_X-Wing " & ": " & Candidat & " rouge à enlever."
          Case "Shm_L", "Shm_C"
            Frm_SDK.B_Info.Text = "Stratégie Sashimi Double " & ": " & Candidat & " rouge à enlever."
        End Select
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub
  Public Sub G4_Grid_Stratégie_XYw_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String

    If Not Plcy_Strg = "XYw" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_XYw(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        Select Case Mid$(Strategy_Rslt(2, Ligne), 1, 1)
                'U_Strg est documenté dans G4_MdC_Row_Col_Box 
          Case "C"
            G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
            G4_MdC_Row_Col_Box("Row", U_Row(CInt(Strategy_Rslt(10, Ligne))))
          Case "F"
            G4_MdC_Row_Col_Box("Col", U_Col(CInt(Strategy_Rslt(10, Ligne))))
            G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
          Case Else
            G4_MdC_Row_Col_Box("Row", U_Row(CInt(Strategy_Rslt(55, Ligne))))
            G4_MdC_Row_Col_Box("Col", U_Col(CInt(Strategy_Rslt(55, Ligne))))
        End Select
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés
        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ": " & Candidat & " rouge à enlever."
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub G4_Grid_Stratégie_Swf_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String

    If Not Plcy_Strg = "Swf" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_Swf(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        For k As Integer = 10 To 54
          Dim Cell_s As String = Strategy_Rslt(k, Ligne)
          If Cell_s = "__" Then Exit For
          Dim Cell_i As Integer = CInt(Strategy_Rslt(k, Ligne))
          G4_MdC_Row_Col_Box("Row", U_Row(Cell_i))
          G4_MdC_Row_Col_Box("Col", U_Col(Cell_i))
        Next k
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés
        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ": " & Candidat & " rouge à enlever."
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub G4_Grid_Stratégie_Jly_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String

    If Not Plcy_Strg = "Jly" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_Jly(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        For k As Integer = 10 To 54
          Dim Cell_s As String = Strategy_Rslt(k, Ligne)
          If Cell_s = "__" Then Exit For
          Dim Cell_i As Integer = CInt(Strategy_Rslt(k, Ligne))
          G4_MdC_Row_Col_Box("Row", U_Row(Cell_i))
          G4_MdC_Row_Col_Box("Col", U_Col(Cell_i))
        Next k
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés
        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ": " & Candidat & " rouge à enlever."
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub G4_Grid_Stratégie_XYZ_g(g As Graphics)
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidat As String

    If Not Plcy_Strg = "XYZ" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_XYZ(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      ' 2 Aide Graphique
      '   Une cellule est cliquée (Pbl_Cell_Select), à quelle stratégie correspont-elle ?
      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Candidat = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
        'U_Strg est documenté dans G4_MdC_Row_Col_Box 
        G4_MdC_Paint_g(g)  ' Les figures sont dessinées et les candidats affichés
        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & ": " & Candidat & " rouge à enlever."
      End If
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub

  Public Sub G4_Grid_Stratégie_SKy_g(g As Graphics)
    'La stratégie SKy regroupe 2 Sous-Stratégies:
    '   s/Stratégie   SKY Skyscraper Gratte-Ciel
    '   s/Stratégie   Kyt Kyte       Cerf-Volant
    '   s/Stratégie   EyR Empty Rct  Rectangle Vide
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer

    Dim Candidat As String
    If Not Plcy_Strg = "SKy" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_SKy(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Dim Sous_stratégie As String = Strategy_Rslt(2, Ligne)
        Candidat = Strategy_Rslt(5, Ligne)
        Select Case Mid$(Sous_stratégie, 1, 3)
          Case "SKy"
            U_MdC_Init()
            G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(10, Ligne)), CInt(Strategy_Rslt(11, Ligne)))
            G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(10, Ligne)), CInt(Strategy_Rslt(12, Ligne)))
            G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(11, Ligne)), CInt(Strategy_Rslt(13, Ligne)))
            G4_MdC_Paint_g(g)      ' Les figures sont dessinées et les candidats affichés
            For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
              If Strategy_Rslt(k, Ligne) = "__" Then Exit For
              Cellule = CInt(Strategy_Rslt(k, Ligne))
              Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
              sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
              U_Strg(Cellule) = True
            Next k
            Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " " & Sous_stratégie & ": " & Candidat & "  rouge à exclure."

          Case "Kyt"
            U_MdC_Init()
            G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(10, Ligne)), CInt(Strategy_Rslt(11, Ligne)))
            G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(12, Ligne)), CInt(Strategy_Rslt(13, Ligne)))
            Select Case Sous_stratégie
              Case "KytH1V1" : G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(12, Ligne))))
              Case "KytH1V2" : G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
              Case "KytH2V1" : G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(12, Ligne))))
              Case "KytH2V2" : G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(11, Ligne))))
            End Select
            G4_MdC_Paint_g(g)      ' Les figures sont dessinées et les candidats affichés
            For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
              If Strategy_Rslt(k, Ligne) = "__" Then Exit For
              Cellule = CInt(Strategy_Rslt(k, Ligne))
              Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
              sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
              U_Strg(Cellule) = True
            Next k
            Frm_SDK.B_Info.Text = " Stratégie Sky " & Sous_stratégie & ": " & Candidat & "  rouge à exclure."

          Case "EyR"
            U_MdC_Init()
            G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
            G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(11, Ligne)), CInt(Strategy_Rslt(12, Ligne)))
            G4_MdC_Paint_g(g)      ' Les figures sont dessinées et les candidats affichés
            For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
              If Strategy_Rslt(k, Ligne) = "__" Then Exit For
              Cellule = CInt(Strategy_Rslt(k, Ligne))
              Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
              sc_Cdd.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Exclure)
              U_Strg(Cellule) = True
            Next k
            Frm_SDK.B_Info.Text = " Stratégie Sky " & Sous_stratégie & ": " & Candidat & "  rouge à exclure."

        End Select
      End If

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_Unq_g(g As Graphics)
    'Unq est traité pour l'aide simple comme les autres stratégies avec Strategy_BTXYSJZKQ_Aide_Simple
    '               Pour l'Aide Graphique, le traitement diffère : 
    ' Strategy_BTXYSJZK_ :  Dim Candidat as String  = Strategy_Rslt(5, Ligne)
    '                       comporte LE candidat à éliminer
    ' Strategy_________Q :  Dim Candidats As String = Strategy_Rslt(5, Ligne)
    '                       comporte LES CANDIDATS
    Dim U_temp(80, 3) As String
    Dim Ligne As Integer
    Dim Strategy_Rslt(,) As String
    Dim Cellule As Integer
    Dim Candidats_Txt As String = ""
    Dim Candidats As String
    If Not Plcy_Strg = "Unq" Then Exit Sub
    Try
      U_Strg_Effacer_g(g)
      Array.Copy(U, U_temp, UNbCopy)
      Strategy_Rslt = Strategy_Unq(U_temp)
      If UBound(Strategy_Rslt, 2) <= 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Aide Simple
      Strategy_BTXYSJZKQ_Aide_Simple_g(g, Strategy_Rslt, U_Strg)
      Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte

      Ligne = Strategy_Click(Pbl_Cell_Select, Strategy_Rslt)
      If Ligne <> -1 Then
        Dim Sous_stratégie As String = Strategy_Rslt(2, Ligne)
        'Strategy_Unq_Rectangle_1 candidats comporte 2 candidats (placés dans un string 9)
        '                         la typologie est R11, R12, R13, R14 suivant le sens du rectangle
        '                         Les candidats sont correctement placés dans la zone Candidats
        'Strategy_Unq_Rectangle_2 candidats comporte 1 candidat  
        '                         la typologie est R2HD, R2HG, R2VH, R2VB
        '                         Le candidat est correctement placé dans la zone candidats (string 9)
        '                         
        Candidats = Strategy_Rslt(5, Ligne)
        U_MdC_Init()
        G4_MdC_Trait_ou_Rectangle(CInt(Strategy_Rslt(10, Ligne)), CInt(Strategy_Rslt(12, Ligne)))
        G4_MdC_Paint_g(g)      ' Les figures sont dessinées et les candidats affichés
        For k As Integer = 55 To 99    ' Re-dessine le candidat à enlever dans un cercle plein rouge
          If Strategy_Rslt(k, Ligne) = "__" Then Exit For
          Cellule = CInt(Strategy_Rslt(k, Ligne))
          Dim sc_Cdd As New Cellule_Cls With {.Numéro = Cellule}
          Candidats_Txt = ""
          For c As Integer = 1 To 9
            If Mid$(Candidats, c, 1) = CStr(c) Then
              sc_Cdd.G6_Cellule_Paint_Candidat_g(g, CStr(c), Color_Cdd_Exclure)
              Candidats_Txt &= CStr(c) & ","
            End If
          Next c
          'Pour perdre la dernière virgule
          Candidats_Txt = Mid$(Candidats_Txt, 1, Candidats_Txt.Length - 1)
          U_Strg(Cellule) = True
        Next k
        Frm_SDK.B_Info.Text = Stg_Get(Strategy_Rslt(1, 0)).Texte & " " & Sous_stratégie & ": " & Candidats_Txt & "  rouge à exclure."
      End If

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_XCx_XCy_XNl_g(g As Graphics)
    If Not (Plcy_Strg = "XCx" Or Plcy_Strg = "XCy" Or Plcy_Strg = "XNl") Then Exit Sub
    Try
      'Dsp_AideGraphique("Oui") 'Nécessaire pour afficher et colorier le menu contextuel
      Dim sc As New Cellule_Cls
      If XRslt.Productivité = False Then
        'Dsp_AideGraphique("Non")
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(XRslt.Candidat(0)) Then
          G0_Cdd_Figure_g(g, i, CInt(XRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        With Link
          Select Case Plcy_Strg
            Case "XCx" 'Alternance de liens forts et faibles, sans alternance de couleurs des candidats 
              G0_Cdd_Bézier_g(g, .Cel(0), CInt(.Cdd(0)), .Cel(1), CInt(.Cdd(2)), .Type, Nb)
            Case "XCy"
              G0_Cdd_Bézier_g(g, .Cel(0), CInt(.Cdd(4)), .Cel(1), CInt(.Cdd(4)), .Type, Nb)
            Case "XNl"
              G0_Cdd_Bézier_g(g, .Cel(0), CInt(.Cdd(4)), .Cel(1), CInt(.Cdd(4)), .Type, Nb)
            Case Else
          End Select

          ' 3 Affichage des Extrémités des liens  
          Dim PremierLien As XLink_Cls = XRslt.RoadRight.First()
          Dim DernierLien As XLink_Cls = XRslt.RoadRight.Last()
          G0_Cell_Icône_g(g, PremierLien.Cel(0), "Start")
          G0_Cell_Icône_g(g, DernierLien.Cel(1), "End")
          Select Case Plcy_Strg
            Case "XCy"
              G0_Cdd_Figure_g(g, PremierLien.Cel(0), CInt(PremierLien.Cdd(0)), "Disque", Color_Link_S)
              G0_Cdd_Figure_g(g, DernierLien.Cel(1), CInt(DernierLien.Cdd(1)), "Disque", Color_Link_S)
          End Select

        End With
      Next Link

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As XCel_Excl_Cls In XRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next XCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & XRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_XRp_g(g As Graphics)
    If Not Plcy_Strg = "XRp" Then Exit Sub
    Try
      'Dsp_AideGraphique("Oui") 'Nécessaire pour afficher et colorier le menu contextuel
      Dim sc As New Cellule_Cls
      If XRslt.Productivité = False Then
        ' Dsp_AideGraphique("Non")
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(XRslt.Candidat(0)) And U(i, 3).Contains(XRslt.Candidat(1)) Then
          G0_Cdd_Figure_g(g, i, CInt(XRslt.Candidat(0)), "Cercle", Color.White)
          G0_Cdd_Figure_g(g, i, CInt(XRslt.Candidat(1)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        With Link
          Select Case Nb Mod 2 ' Pour alterner les liens entre les séries de candidats
            Case 0 : G0_Cdd_Bézier_g(g, .Cel(0), CInt(.Cdd(0)), .Cel(1), CInt(.Cdd(2)), .Type, Nb)
            Case 1 : G0_Cdd_Bézier_g(g, .Cel(0), CInt(.Cdd(1)), .Cel(1), CInt(.Cdd(3)), .Type, Nb)
          End Select
        End With
      Next Link

      ' 3 Affichage des Extrémités des liens  
      Dim PremierLien As XLink_Cls = XRslt.RoadRight.First()
      G0_Cell_Icône_g(g, PremierLien.Cel(0), "Start")
      G0_Cdd_Figure_g(g, PremierLien.Cel(0), CInt(XRslt.Candidat(0)), "Disque", Color_Link_W)
      G0_Cdd_Figure_g(g, PremierLien.Cel(0), CInt(XRslt.Candidat(1)), "Disque", Color_Link_W)

      Dim DernierLien As XLink_Cls = XRslt.RoadRight.Last()
      G0_Cell_Icône_g(g, DernierLien.Cel(1), "End")
      G0_Cdd_Figure_g(g, DernierLien.Cel(1), CInt(XRslt.Candidat(0)), "Disque", Color_Link_W)
      G0_Cdd_Figure_g(g, DernierLien.Cel(1), CInt(XRslt.Candidat(1)), "Disque", Color_Link_W)

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As XCel_Excl_Cls In XRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd           ' le candidat est ajouté
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next XCelExcl

      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & XRslt.Candidat(0) & " " & XRslt.Candidat(1) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub

  Public Sub G4_Grid_Stratégie_WgX_WgY_WgZ_WgW_g(g As Graphics)
    If Not (Plcy_Strg = "WgX" Or Plcy_Strg = "WgY" Or Plcy_Strg = "WgZ" Or Plcy_Strg = "WgW") Then Exit Sub

    Try
      'Dsp_AideGraphique("Oui") 'Nécessaire pour afficher et colorier le menu contextuel
      Dim sc As New Cellule_Cls
      If XRslt.Productivité = False Then
        'Dsp_AideGraphique("Non")
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(XRslt.Candidat(0)) Then
          G0_Cdd_Figure_g(g, i, CInt(XRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        Select Case Plcy_Strg
          Case "WgX"
            G0_Cdd_Bézier_g(g, Link.Cel(0), CInt(Link.Cdd(0)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
          Case "WgY"
            G0_Cdd_Bézier_g(g, Link.Cel(0), CInt(Link.Cdd(2)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
            G0_Cdd_Figure_g(g, Link.Cel(1), CInt(Link.Cdd(3)), "Disque", Color_Link_S) 'Mise en exergue du candidat Z
          Case "WgZ"
            G0_Cdd_Bézier_g(g, Link.Cel(0), CInt(Link.Cdd(3)), Link.Cel(1), CInt(Link.Cdd(3)), Link.Type, Nb)
            G0_Cdd_Bézier_g(g, Link.Cel(0), CInt(Link.Cdd(4)), Link.Cel(1), CInt(Link.Cdd(4)), Link.Type, Nb)
            G0_Cdd_Figure_g(g, Link.Cel(0), CInt(Link.Cdd(5)), "Disque", Color_Link_S) 'Mise en exergue du candidat du pivot
          Case "WgW"
            G0_Cdd_Bézier_g(g, Link.Cel(0), CInt(Link.Cdd(0)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
            G0_Cdd_Figure_g(g, XRslt.Cellule(0), CInt(XRslt.Candidat(1)), "Disque", Color_Link_S) 'Mise en exergue du candidat  
            G0_Cdd_Figure_g(g, XRslt.Cellule(1), CInt(XRslt.Candidat(1)), "Disque", Color_Link_S) 'Mise en exergue du candidat  
        End Select
      Next Link

      ' 3 Affichage des Extrémités des liens  
      Select Case Plcy_Strg
        Case "WgX"
          G0_Cell_Icône_g(g, XRslt.RoadRight.Item(0).Cel(0), "Start")
          G0_Cell_Icône_g(g, XRslt.RoadRight.Item(1).Cel(0), "Start")
        Case "WgY"
          G0_Cell_Icône_g(g, XRslt.RoadRight.Item(0).Cel(0), "Start")
        Case "WgZ"
          G0_Cell_Icône_g(g, XRslt.RoadRight.Item(0).Cel(0), "Start")
        Case "WgW"
          G0_Cell_Icône_g(g, XRslt.Cellule(0), "Start")
          G0_Cell_Icône_g(g, XRslt.Cellule(1), "Start")
      End Select

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As XCel_Excl_Cls In XRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next XCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & XRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_GLk_g(g As Graphics)
    If Not Plcy_Strg = "GLk" Then Exit Sub
    Try
      If GLinks.Count = 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats dans la grille
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure_g(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Liens Forts de la list GLinks
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GLinks
        Nb += 1
        G0_Cdd_Bézier_g(g, gLink.Cel(0), CInt(gLink.Cdd(0)), gLink.Cel(1), CInt(gLink.Cdd(2)), gLink.Type, Nb)
        Dim sc_gLink As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(0)) Then
          sc_gLink.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(0), Color_Link_S)
        End If
        sc_gLink.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(2)) Then
          sc_gLink.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(2), Color_Link_S)
        End If
      Next gLink
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " " & GRslt.Nb_Liens & " Liens Forts."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_Gbl_g(g As Graphics)
    If Not Plcy_Strg = "Gbl" Then Exit Sub

    Try
      'Dsp_AideGraphique("Oui") 'Nécessaire pour afficher et colorier le menu contextuel
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        'Dsp_AideGraphique("Non")
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure_g(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage de toutes les Courbes de Bézier
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        Nb += 1
        G0_Cdd_Bézier_g(g, gLink.Cel(0), CInt(gLink.Cdd(0)), gLink.Cel(1), CInt(gLink.Cdd(2)), gLink.Type, Nb)
        Dim sc_gLink As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(0)) Then
          sc_gLink.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(0), Color_Link_S)
        End If
        sc_gLink.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(2)) Then
          sc_gLink.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(2), Color_Link_S)
        End If
      Next gLink

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each gCel As GCel_Excl_Cls In GRslt.CelExcl
        With gCel
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            If Pbl_Cell_Select = .Cel Then
              Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
              U_Strg_Cdd_Exc(.Cel) = Candidats
            End If
          End If
        End With
      Next gCel
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.CelExcl.Count & " candidat(s) rouge(s) à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub
  Public Sub G4_Grid_Stratégie_Gbv_g(g As Graphics)
    If Not Plcy_Strg = "Gbv" Then Exit Sub

    Try
      'Dsp_AideGraphique("Oui") 'Nécessaire pour afficher et colorier le menu contextuel
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        'Dsp_AideGraphique("Non")
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
      Next i

      ' 2 Affichage de toutes les Courbes de Bézier
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        Nb += 1
        ' Premier lien
        G0_Cdd_Bézier_g(g, gLink.Cel(0), CInt(gLink.Cdd(0)), gLink.Cel(1), CInt(gLink.Cdd(2)), gLink.Type, Nb)
        Dim sc_gLink1 As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(0)) Then
          sc_gLink1.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(0), Color_Link_S)
        End If
        sc_gLink1.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(2)) Then
          sc_gLink1.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(2), Color_Link_S)
        End If
        ' Second lien
        G0_Cdd_Bézier_g(g, gLink.Cel(0), CInt(gLink.Cdd(1)), gLink.Cel(1), CInt(gLink.Cdd(3)), gLink.Type, 0)
        Dim sc_gLink2 As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(1)) Then
          sc_gLink2.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(1), Color_Link_S)
        End If
        sc_gLink2.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(3)) Then
          sc_gLink2.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(3), Color_Link_S)
        End If
      Next gLink

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As GCel_Excl_Cls In GRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            If Pbl_Cell_Select = .Cel Then
              Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
              U_Strg_Cdd_Exc(.Cel) = Candidats
            End If
          End If
        End With
      Next XCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.CelExcl.Count & " candidat(s) rouge(s) à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_GCs_g(g As Graphics)
    If Not (Plcy_Strg = "GCs") Then Exit Sub
    Try
      'Dsp_AideGraphique("Oui") 'Nécessaire pour afficher et colorier le menu contextuel
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        ' Dsp_AideGraphique("Non")
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure_g(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        Nb += 1
        With gLink
          G0_Cdd_Bézier_g(g, .Cel(0), CInt(.Cdd(0)), .Cel(1), CInt(.Cdd(2)), .Type, Nb)
          Select Case Nb Mod 2
            Case 0
              Dim sc_Cdd As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
              sc_Cdd.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(0), Color.Green)
            Case Else
              Dim sc_cdd As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
              sc_cdd.G6_Cellule_Paint_Candidat_g(g, gLink.Cdd(0), Color.Blue)
          End Select
        End With
      Next gLink

      ' 3 Affichage des Extrémités des liens et du Dernier candidat de la chaîne
      Dim firstLien As GLink_Cls = GRslt.RoadRight.First()
      Dim lastLien As GLink_Cls = GRslt.RoadRight.Last()
      Dim sc_LCdd As New Cellule_Cls With {.Numéro = lastLien.Cel(1)}
      Dim lastIndex As Integer = GRslt.RoadRight.Count - 1
      Select Case lastIndex Mod 2
        Case 0 : sc_LCdd.G6_Cellule_Paint_Candidat_g(g, lastLien.Cdd(0), Color.Green)
        Case 1 : sc_LCdd.G6_Cellule_Paint_Candidat_g(g, lastLien.Cdd(0), Color.Blue)
      End Select
      G0_Cell_Icône_g(g, firstLien.Cel(0), "Start")
      G0_Cell_Icône_g(g, lastLien.Cel(1), "End")

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each gCelExcl As GCel_Excl_Cls In GRslt.CelExcl
        With gCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat_g(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next gCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_Obj_g(g As Graphics)
    Dim sc As New Cellule_Cls
    If Not Plcy_Strg = "Obj" Then Exit Sub
    For i As Integer = 0 To 80
      sc.Numéro = i
      sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")
    Next i

    For Each Obj As Objet_Cls In Objet_List
      With Obj

        Select Case .Forme
          Case "Cadre", "Carré", "Disque", "Cercle", "Croix"
            Select Case .Cdd_From
              Case 0    'Cellule
                G0_Cell_Figure_g(g, .Cel_From, .Forme, Color_BySymbol(.Symbol))
              Case Else 'Candidat
                G0_Cdd_Figure_g(g, .Cel_From, .Cdd_From, .Forme, Color_BySymbol(.Symbol))
            End Select
          Case "Flèche"
            G0_Cdd_Flèche_g(g, .Cel_From, .Cdd_From, .Cel_To, .Cdd_To, Color_BySymbol(.Symbol))
          Case Else
        End Select
      End With
    Next Obj
  End Sub

#End Region

#Region "Paint Figures et Formes de base"

  Public Sub G0_Cell_Icône_g(g As Graphics, Cellule As Integer, Icône As String)
    ' Place une icône de Début dans le premier candidat "libre" de la cellule
    '       une icône de Fin   dans le second candidat "libre" de la cellule

    Dim Cdd_i As Integer() = {5, 1, 3, 7, 9, 2, 4, 6, 8}
    Dim candidatsVides As Integer() = Cdd_i.
       Where(Function(c) Not U(Cellule, 3).Contains(c.ToString())).Take(2).ToArray()

    Select Case Icône
      Case "Start"
        Dim rect As Rectangle = Sqr_Cdd((Cellule * 10) + candidatsVides(0))
        Using img As Image = My.Resources.XC_Start
          g.DrawImage(img, rect)
        End Using
      Case "End"
        Dim rect As Rectangle = Sqr_Cdd((Cellule * 10) + candidatsVides(1))
        Using img As Image = My.Resources.XC_End
          g.DrawImage(img, rect)
        End Using
    End Select
  End Sub
  Public Sub G0_Cell_Figure_g(g As Graphics, Cellule As Integer, Figure As String, Couleurp As Color)
    Dim Couleur As Color = Color.FromArgb(192, Couleurp) ' 128+64 = 192

    'Dessine des figures à partir de 4 points A0, B0, C0, D0 représentant les 4 points du square A0= Left_Top, C0= Right_Bottom
    '                             déclinés en A1, B1, C1, D1 décalé de 1/4 
    '                             déclinés en A2, B2, C2, D2 décalé de 2/4 
    '                             déclinés en A3, B3, C3, D3 décalé de 3/4 
    Const Inflate As Integer = -5

    Dim Rct_Cercle As New Rectangle(x:=CInt(U_Pt20(Cellule, 0).X), CInt(U_Pt20(Cellule, 0).Y), width:=WH - 3, height:=WH - 3)
    Dim Pt1 As Point, Pt2 As Point, Pt3 As Point, Pt4 As Point
    Pt1 = New Point(CInt(U_Pt20(Cellule, 0).X), CInt(U_Pt20(Cellule, 0).Y))
    Pt2 = New Point(CInt(U_Pt20(Cellule, 1).X), CInt(U_Pt20(Cellule, 1).Y))
    Pt3 = New Point(CInt(U_Pt20(Cellule, 2).X), CInt(U_Pt20(Cellule, 2).Y))
    Pt4 = New Point(CInt(U_Pt20(Cellule, 3).X), CInt(U_Pt20(Cellule, 3).Y))

    Using pen As New Pen(Couleur, 2),
          brsh As New SolidBrush(Couleur),
          pen_Inflate As New Pen(Couleur, Math.Abs(Inflate) * 2)
      Select Case Figure
        Case "Cadre"
          ' Comme le trait du rectangle dépasse de la moitié à droite et à gauche du trait, l'inflate doit être le double
          Rct_Cercle.Inflate(Inflate, Inflate)
          g.DrawRectangle(pen_Inflate, Rct_Cercle)
        Case "Carré"
          g.FillRectangle(brsh, Rct_Cercle)
        Case "Croix"
          g.DrawLine(pen, Pt1, Pt3)
          g.DrawLine(pen, Pt2, Pt4)
        Case "Cercle"
          g.DrawArc(pen, Rct_Cercle, 0.0F, 360.0F)
        Case "Disque"
          g.FillPie(brsh, Rct_Cercle, 0.0F, 360.0F)
        Case "Ellipse"
          '28/01/2025 Cette figure n'est plus utilisée
          Dim Rct_V As New Rectangle(x:=Sqr_Cel(Cellule).X + (1 * WHthird), y:=Sqr_Cel(Cellule).Y, width:=WHthird, height:=WH - 3)
          Dim Rct_H As New Rectangle(x:=Sqr_Cel(Cellule).X + 1, y:=Sqr_Cel(Cellule).Y + (1 * WHthird), width:=WH - 3, height:=WHthird)
          g.DrawEllipse(pen, Rct_V)
          g.DrawEllipse(pen, Rct_H)
          g.DrawArc(pen, Rct_Cercle, 0.0F, 360.0F)
        Case "Double_Carré"
          g.DrawPolygon(pen, {U_Pt20(Cellule, 4), U_Pt20(Cellule, 5), U_Pt20(Cellule, 6), U_Pt20(Cellule, 7)})
          g.DrawPolygon(pen, {U_Pt20(Cellule, 12), U_Pt20(Cellule, 13), U_Pt20(Cellule, 14), U_Pt20(Cellule, 15)})
        Case Else
          Jrn_Add(, {Proc_Name_Get() & " Figure Inconnue : " & Figure}, "Erreur")
      End Select
    End Using
  End Sub

  Public Sub G0_Cdd_Figure_g(g As Graphics, Cellule As Integer, Candidat As Integer, Figure As String, Couleurp As Color)
    Dim Couleur As Color = Color.FromArgb(192, Couleurp) ' 128+64 = 192
    'Dessine des figures à partir de 4 points A0, B0, C0, D0 représentant les 4 points du square A0= Left_Top, C0= Right_Bottom
    '                             déclinés en A1, B1, C1, D1 décalé de 1/4 
    '                             déclinés en A2, B2, C2, D2 décalé de 2/4 
    '                             déclinés en A3, B3, C3, D3 décalé de 3/4 

    If Cellule < 0 Or Cellule > 80 Then Exit Sub
    If Candidat < 1 Or Candidat > 9 Then Exit Sub

    Dim Cdd_n As Integer = (Cellule * 10) + Candidat
    Dim Sqr_Cdd_n As Rectangle = Sqr_Cdd(Cdd_n)
    Sqr_Cdd_n.Inflate(-1, -1)    'Diminution du cercle du candidat  
    ' Définir les coins du rectangle
    Dim Pt1 As New Point(Sqr_Cdd_n.X, Sqr_Cdd_n.Y)                                      ' Coin supérieur gauche
    Dim Pt2 As New Point(Sqr_Cdd_n.X + Sqr_Cdd_n.Width, Sqr_Cdd_n.Y)                    ' Coin supérieur droit
    Dim Pt3 As New Point(Sqr_Cdd_n.X + Sqr_Cdd_n.Width, Sqr_Cdd_n.Y + Sqr_Cdd_n.Height) ' Coin inférieur droit
    Dim Pt4 As New Point(Sqr_Cdd_n.X, Sqr_Cdd_n.Y + Sqr_Cdd_n.Height)                   ' Coin inférieur gauche

    Using pen As New Pen(Couleur, 2),
          brsh As New SolidBrush(Couleur)
      Select Case Figure
        Case "Cadre"
          g.DrawRectangle(pen, Sqr_Cdd_n)
        Case "Carré"
          g.FillRectangle(brsh, Sqr_Cdd_n)
        Case "Croix"
          g.DrawLine(pen, Pt1, Pt3)
          g.DrawLine(pen, Pt2, Pt4)
        Case "Cercle"
          g.DrawArc(pen, Sqr_Cdd_n, 0.0F, 360.0F)
        Case "Disque"
          g.FillPie(brsh, Sqr_Cdd_n, 0.0F, 360.0F)
        Case Else
          Jrn_Add(, {Proc_Name_Get() & " Figure Inconnue : " & Figure}, "Erreur")
      End Select
    End Using
  End Sub

  Function DeCasteljau(t As Single, p0 As PointF, p1 As PointF, p2 As PointF, p3 As PointF) As PointF
    ' Algorithme de De Casteljau
    Dim x As Double = (1 - t) ^ 3 * p0.X +
          3 * (1 - t) ^ 2 * t * p1.X +
          3 * (1 - t) * t ^ 2 * p2.X +
          t ^ 3 * p3.X
    Dim y As Double = (1 - t) ^ 3 * p0.Y +
          3 * (1 - t) ^ 2 * t * p1.Y +
          3 * (1 - t) * t ^ 2 * p2.Y +
          t ^ 3 * p3.Y
    Return New PointF(CSng(x), CSng(y))
  End Function

  Public Sub G0_Cdd_Bézier_g(g As Graphics, From_Cellule As Integer, From_Candidat As Integer, To_Cellule As Integer, To_Candidat As Integer, Link_Type As String, Link_Numéro As Integer)
    ' 1 Calcul des Centres et des Points de contrôle pour une courbe de Bézier
    Dim From_Centre As PointF = Get_CentreF(From_Cellule, From_Candidat)
    Dim To_Centre As PointF = Get_CentreF(To_Cellule, To_Candidat)
    Dim Pts As Points_Struct = Get_AdjustedPoints(From_Centre, To_Centre)
    Dim Décalage As Integer = 5
    Dim From_Ctrl As New PointF(From_Centre.X + Décalage, From_Centre.Y + Décalage)
    Dim To_Ctrl As New PointF(To_Centre.X + Décalage, To_Centre.Y + Décalage)

    ' 2 Dessine la courbe de Bézier
    Dim Couleur As Color
    Dim Style As DashStyle

    Select Case Link_Type
      Case "S"
        Couleur = Color_Link_S
        Style = DashStyle.Solid
      Case "W"
        Couleur = Color_Link_W
        Style = DashStyle.Dash
    End Select

    Using Crayon As New Pen(Couleur, 2)
      Dim Cap As New AdjustableArrowCap(4, 6)
      Crayon.CustomEndCap = Cap
      Crayon.DashStyle = Style
      g.SmoothingMode = SmoothingMode.AntiAlias
      g.DrawBezier(Crayon, Pts.Pt_From, From_Ctrl, To_Ctrl, Pts.Pt_To)
    End Using

    ' 3 Colorie les candidats dans les cellules de départ et d'arrivée
    ' TODO Ce traitement sera enlevé avec la dernière stratégie.
    Select Case Plcy_Strg
      ' TODO à revoir ce devrait être autre chose 
      Case "GLk" ' Exécuter dans G4_Grid_Stratégie_GLk
      Case "Gbl" ' Exécuter dans G4_Grid_Stratégie_Gbl
      Case "Gbv" ' Exécuter dans G4_Grid_Stratégie_Gbl
      Case "GCs" ' Exécuter dans G4_Grid_Stratégie_Gbl
      Case "XNl"
        ' C'est une alternance lien-fort lien-faible 
        Select Case Link_Numéro Mod 2
          Case 0
            Dim sc As New Cellule_Cls With {.Numéro = To_Cellule}
            sc.G6_Cellule_Paint_Candidat_g(g, CStr(To_Candidat), Color.Green)
          Case Else
            Dim sc As New Cellule_Cls With {.Numéro = To_Cellule}
            sc.G6_Cellule_Paint_Candidat_g(g, CStr(To_Candidat), Color.Blue)
        End Select
      Case Else
        Dim sc As New Cellule_Cls With {.Numéro = From_Cellule}
        'Les candidats sont dessinés s'ils existent
        If U(From_Cellule, 3).Contains(CStr(From_Candidat)) Then
          sc.G6_Cellule_Paint_Candidat_g(g, CStr(From_Candidat), Color_Link_W)
        End If
        sc.Numéro = To_Cellule
        If U(To_Cellule, 3).Contains(CStr(To_Candidat)) Then
          sc.G6_Cellule_Paint_Candidat_g(g, CStr(To_Candidat), Color_Link_W)
        End If
    End Select

    ' 4 Affichage de la lettre du lien
    Dim PointMid As PointF = DeCasteljau(0.5, Pts.Pt_From, From_Ctrl, To_Ctrl, Pts.Pt_To)
    Using font As New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Italic),
          brsh As New SolidBrush(Color.Black)
      g.DrawString(ChrW(Link_Numéro + Lettre_Flèche_ChrW), font, brsh, PointMid.X, PointMid.Y)
    End Using
  End Sub

  Public Sub G0_Cell_Diagonale_g(g As Graphics, From_Cellule As Integer, To_Cellule As Integer)
    'Dessine une Flèche
    If From_Cellule = -1 And To_Cellule = -1 Then Exit Sub
    Dim Pt_From_Cellule, Pt_To_Cellule As Point
    Dim sc_From As New Cellule_Cls With {.Numéro = From_Cellule}
    Pt_From_Cellule = New Point(sc_From.Position_Center.X, sc_From.Position_Center.Y)
    Dim sc_To As New Cellule_Cls With {.Numéro = To_Cellule}
    Pt_To_Cellule = New Point(sc_To.Position_Center.X, sc_To.Position_Center.Y)
    Cells_Bresenham_Get(U_Strg, From_Cellule, To_Cellule)
    Using Pen As New Pen(Color_Stratégique, 5) With {.DashStyle = DashStyle.DashDotDot, .EndCap = LineCap.ArrowAnchor}
      g.DrawLine(Pen, Pt_From_Cellule, Pt_To_Cellule)
    End Using
  End Sub
  Public Sub Cells_Bresenham_Get(ByRef U_Strg() As Boolean, ByVal From_Cellule As Integer, ByVal To_Cellule As Integer)
    ' Documente U_Strg des cellules traversées par la ligne Pt_From Pt_To
    ' Algorithme de la droite de Bresenham IBM 1962
    ' La position Top-Left ou center des Pt_From_To donne des résultats différents, 
    '    ainsi que l'épaisseur du trait  
    Dim sc_From As New Cellule_Cls With {.Numéro = From_Cellule}
    Dim Pt_From As New Point(sc_From.Position_Center.X, sc_From.Position_Center.Y)
    Dim sc_To As New Cellule_Cls With {.Numéro = To_Cellule}
    Dim Pt_To As New Point(sc_To.Position_Center.X, sc_To.Position_Center.Y)

    Dim dx As Integer = Math.Abs(Pt_To.X - Pt_From.X)
    Dim dy As Integer = Math.Abs(Pt_To.Y - Pt_From.Y)
    Dim sx As Integer = If(Pt_From.X < Pt_To.X, 1, -1)
    Dim sy As Integer = If(Pt_From.Y < Pt_To.Y, 1, -1)
    Dim err As Integer = dx - dy

    Dim x As Integer = Pt_From.X
    Dim y As Integer = Pt_From.Y

    While True
      Dim Cellule As Integer = Array.FindIndex(Sqr_Cel, Function(cel) cel.Contains(New Point(x, y)))
      If Cellule <> -1 Then U_Strg(Cellule) = True
      If x = Pt_To.X And y = Pt_To.Y Then Exit While
      Dim e2 As Integer = 2 * err
      If e2 > -dy Then
        err -= dy
        x += sx
      End If
      If e2 < dx Then
        err += dx
        y += sy
      End If
    End While
  End Sub
  Public Sub G4_MdC_Row_Col_Box(Code_LCR As String, LCR As Integer)
    If LCR < 0 Or LCR > 8 Then Exit Sub
    Dim Grp(0 To 8) As Integer
    Dim Cellule As Integer
    Select Case Code_LCR
      Case "Row"
        Grp = U_9CelRow(LCR)
        For i As Integer = 0 To 8
          Cellule = Grp(i)
          Select Case i
            Case 0 : G4_Cell_MdC(Cellule, "PD")
            Case 8 : G4_Cell_MdC(Cellule, "PG")
            Case Else : G4_Cell_MdC(Cellule, "RH")
          End Select
        Next i
      Case "Col"
        Grp = U_9CelCol(LCR)
        For i As Integer = 0 To 8
          Cellule = Grp(i)
          Select Case i
            Case 0 : G4_Cell_MdC(Cellule, "PB")
            Case 8 : G4_Cell_MdC(Cellule, "PH")
            Case Else : G4_Cell_MdC(Cellule, "RV")
          End Select
        Next i
      Case "Box"
        Grp = U_9CelReg(LCR)
        'Les cellules de la région sont numérotées
        ' 0 1 2
        ' 3 4 5
        ' 6 7 8
        Dim b As Integer
        b = 0 : G4_Cell_MdC(Grp(b), "CHG")
        b = 1 : G4_Cell_MdC(Grp(b), "RH")
        b = 2 : G4_Cell_MdC(Grp(b), "CHD")
        b = 3 : G4_Cell_MdC(Grp(b), "RV")
        b = 4 : G4_Cell_MdC(Grp(b), "SQ")
        b = 5 : G4_Cell_MdC(Grp(b), "RV")
        b = 6 : G4_Cell_MdC(Grp(b), "CBG")
        b = 7 : G4_Cell_MdC(Grp(b), "RH")
        b = 8 : G4_Cell_MdC(Grp(b), "CBD")
      Case Else
        Exit Select
    End Select
    For i As Integer = 0 To 8 : U_Strg(Grp(i)) = True : Next i
  End Sub

  Public Sub G4_MdC_Trait_ou_Rectangle(Cellule_A As Integer, Cellule_B As Integer)
    'La procédure dessine soit un trait horizontal, soit un trait vertical, soit un rectangle
    'Toujours de 1 vers 2 en croissant, id de gauche à droite, de haut en bas, de la cellule 0 à 81
    Dim Cel, Cel_1, Cel_2 As Integer
    If Cellule_A < Cellule_B Then
      Cel_1 = Cellule_A : Cel_2 = Cellule_B
    Else
      Cel_1 = Cellule_B : Cel_2 = Cellule_A
    End If
    Dim w, h As Integer ' w et h sont toujours positifs
    w = U_Col(Cel_2) - U_Col(Cel_1)
    h = U_Row(Cel_2) - U_Row(Cel_1)

    If Cellule_A = Cellule_B Then Exit Sub

    If h = 0 Then    'Trait Horizontal 
      G4_Cell_MdC(Cel_1, "PD")
      U_Strg(Cel_1) = True
      For i As Integer = U_Col(Cel_1) + 1 To U_Col(Cel_2) - 1 Step 1
        Cel = Wh_Cellule_ColRow(i, U_Row(Cel_1))
        G4_Cell_MdC(Cel, "RH")
        U_Strg(Cel) = True
      Next i
      G4_Cell_MdC(Cel_2, "PG")
      U_Strg(Cel_2) = True
      Exit Sub
    End If '

    If w = 0 Then    'Trait vertical 
      G4_Cell_MdC(Cel_1, "PB")
      U_Strg(Cel_1) = True
      For i As Integer = U_Row(Cel_1) + 1 To U_Row(Cel_2) - 1 Step 1
        Cel = Wh_Cellule_ColRow(U_Col(Cel_1), i)
        G4_Cell_MdC(Cel, "RV")
        U_Strg(Cel) = True
      Next i
      G4_Cell_MdC(Cel_2, "PH")
      U_Strg(Cel_2) = True
      Exit Sub
    End If '

    'Forme rectangulaire
    G4_Cell_MdC(Cel_1, "CHG")
    U_Strg(Cel_1) = True
    For i As Integer = U_Col(Cel_1) + 1 To U_Col(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(i, U_Row(Cel_1))
      G4_Cell_MdC(Cel, "RH")
      U_Strg(Cel) = True
    Next i
    G4_Cell_MdC(Cel_1 + w, "CHD")
    U_Strg(Cel_1 + w) = True
    For i As Integer = U_Row(Cel_1) + 1 To U_Row(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(U_Col(Cel_1), i)
      G4_Cell_MdC(Cel, "RV")
      U_Strg(Cel) = True
    Next i
    G4_Cell_MdC(Cel_2, "CBD")
    U_Strg(Cel_2) = True
    For i As Integer = U_Col(Cel_1) + 1 To U_Col(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(i, U_Row(Cel_2))
      G4_Cell_MdC(Cel, "RH")
      U_Strg(Cel) = True
    Next i
    G4_Cell_MdC(Cel_2 - w, "CBG")
    U_Strg(Cel_2 - w) = True
    For i As Integer = U_Row(Cel_1) + 1 To U_Row(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(U_Col(Cel_2), i)
      G4_Cell_MdC(Cel, "RV")
      U_Strg(Cel) = True
    Next i
  End Sub

  Public Sub G4_Cell_MdC(Cellule As Integer, Modèle As String)
    'Exemple pour dessiner des Modèles
    'U_MdC_Init()           'Initialisation pour chaque cellule d'une suite de modèle
    'G4_Cell_MdC(40, "CHG") 'Ajout d'un modèle dans une cellule 
    'G4_Cell_MdC(41, "CHD")
    'G4_Cell_MdC(50, "CBD")
    'G4_Cell_MdC(49, "CBG")
    'G4_MdC_Row_Col_Box("Box", 7)
    'G4_MdC_Paint()         'Les figures sont dessinées et les candidats affichés

    For j As Integer = 0 To U_MdC(Cellule).MdE.Count - 1
      If U_MdC(Cellule).MdE(j) = "" Then
        U_MdC(Cellule).MdE(j) = Modèle
        U_MdC(Cellule).MdE_Exist = True
        Exit For
      End If
    Next j
  End Sub
  Public Sub G4_MdC_Paint_g(g As Graphics)
    ' Construction et peinture des Modèles Composites
    ' Un modèle est dit composite quand il s'agit de SUPERPOSER plusieurs modèles élémentaires
    '                             comme un croisement de colonne et de ligne
    '           afin qu'il n'y ait pas de superposition de couleurs
    ' Ces modèles sont construits à/p des Modèles Elémentaires
    '   un modèle composite est constitué de PLUSIEURS modèles ELEMENTAIRES
    '                       avec une Région et Union  
    ' Les candidats sont également affichés après la couche stratégique.

    Dim MdC_Exist As Boolean
    Dim MdC_Cellule As Integer
    Dim MdC_N As Integer
    Dim MdC_Modèle As String
    For i As Integer = 0 To 80
      ' Phase 1 Construit une Région d'UNION de tous les modèles à dessiner
      MdC_Exist = False
      MdC_Cellule = U_MdC(i).Cellule
      MdC_N = 0
      Dim MdC_Région As New Region
      For j As Integer = 0 To U_MdC(i).MdE.Count - 1
        If U_MdC(i).MdE(j) <> "" Then
          MdC_Exist = True
          MdC_Modèle = U_MdC(i).MdE(j)
          Dim Pth_Modèle(MdC_N) As GraphicsPath
          Pth_Modèle(MdC_N) = G0_MdE_Build(MdC_Cellule, MdC_Modèle)
          If MdC_N = 0 Then MdC_Région = New Region(Pth_Modèle(MdC_N))
          If MdC_N > 0 Then MdC_Région.Union(Pth_Modèle(MdC_N))
          Pth_Modèle(MdC_N).Dispose()
          MdC_N += 1
        End If
      Next j

      ' Phase 2 Dessine la Région et Peint les valeurs et les candidats concernés
      If MdC_Exist Then
        ' Il est à noter que l'ordre des couches est respecté,
        ' Il n'est donc pas nécessaire de conserver le composant alpha à 128 pour la couche stratégique
        ' La couche stratégique a un effet "grille". 
        Using brsh As New SolidBrush(Color_Stratégique)
          g.FillRegion(brsh, MdC_Région)                    ' G4
        End Using
        ' L'Aide Graphique comporte également l'affichage des candidats 
        Dim sc As New Cellule_Cls With {.Numéro = i}
        If sc.Valeur <> 0 Then sc.G5_Cellule_Paint_Valeur_g(g)                            ' G5
        If sc.Valeur = 0 Then sc.G6_Cellule_Paint_Candidats_g(g, "LesCandidatsEligibles")   ' G6
      End If
      MdC_Région.Dispose()
    Next i

  End Sub

  '-------------------------------------------------------------------------------
  'Création : Lundi 12/12/2020
  '           
  'Idée     : Créer des formes primaires comme un rectangle ou un pouce ou un coude
  '           Décliner ces formes en 2 formes élémentaires Rectangle Horizontal-Vertical
  '                               en 4 formes élémentaires Pouce Haut-Bas et Droit-Gauche
  '                               en 4 formes élémentaires Coude Haut-Droit, Haut-Gauche, Bas-Droit, Bas-Gauche
  '           Combiner des formes élémentaires pour créer une intersection de Rectangles, des Lettres T, 
  '                               des associations rectangles et coude,
  '           etc...
  '-------------------------------------------------------------------------------   
  'Vocabulaire:
  ' Modèle Primaire        R(V)         C (CHG)        P(B)
  '        Elémentaire     RV, RH       C H/B-D/G      P B/G/H/D/G
  '        Composite       La croix = RV Union RH
  '                        
  ' Les modèles primaires sont déterminés à/p de cell_x-y-Width-Height  G0_MdP_Build
  ' Les modèles élémentaires utilisent Matrix, RotateAt, Transform      G0_MdE_Build
  ' Les modèles composites utilisent U_Mdc() de structure MdC_Struct    G4_Cell_MdC  
  '                                                                     G4_MdC_Row_Col_Box
  '                                                                     
  Public Sub U_MdC_Init()
    'Initialisation de U_Mdc : 1 ligne par Cellule
    'Redim doit être fait pour chaque ligne
    '42 L5-C7 :   RV   PD                                         
    '43 L5-C8 :  CHG  CHD  CBD  CBG                               
    '44 L5-C9 :   RV   PG                                         
    '51 L6-C7 :  CBG
    For i As Integer = 0 To 80
      ReDim U_MdC(i).MdE(9)
      U_MdC(i).Cellule = i
      U_MdC(i).MdE_Exist = False
      For j As Integer = 0 To 9
        U_MdC(i).MdE(j) = ""
      Next j
    Next i
  End Sub
  Public Sub U_MdC_Display()
    Dim M1 As String
    Dim M2 As String
    Jrn_Add(, {"Liste de U_MdC"})
    For i As Integer = 0 To 80
      Dim MdC_Exist As Boolean = False
      M2 = ""
      For j As Integer = 0 To U_MdC(i).MdE.Count - 1
        If U_MdC(i).MdE(j) <> "" Then MdC_Exist = True
        M2 &= U_MdC(i).MdE(j).PadLeft(4) & " "
      Next j
      M1 = CStr(i).PadLeft(3) & " " & U_Coord(U_MdC(i).Cellule)
      If MdC_Exist = True Then Jrn_Add(, {M1 & " : " & M2})
    Next i
  End Sub

  Public Function G0_MdP_Build(Cellule As Integer, Modèle As String) As GraphicsPath

    'Nom des Formes, les formes diffèrent selon Format_Ang "Angles droits" ou "Coins arrondis ou biseautés"
    'C  Couronne       B/H Bas-Haut             D/G Droite-Gauche
    'R  Rectangle      H/V Horizontal-Vertical    
    'CR Croix
    'P  Pouce          BHGD
    'PT Point          (Non utilisé!)
    'T  Lettre T       B/H/D/G Bas-Haut-Droit-Gauche

    'Construction des Modèles Primaires
    '----------------------------------
    'Exemple: Modèle primaire du coin haut gauche (droit, arrondi ou biseauté)
    'à/p de ce modèle les modèles élémentaires CHG, CHD, CBD et CBG seront construits par rotation
    'G0_MdP_Build n'est utilisée que dans G0_MdE_Build

    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    Dim Cell_A As New Point(sc.Position.X, sc.Position.Y)
    Dim Cell_B As New Point(sc.Position.X + WH, sc.Position.Y)
    Dim Cell_C As New Point(sc.Position.X + WH, sc.Position.Y + WH)
    Dim Cell_D As New Point(sc.Position.X, sc.Position.Y + WH)

    Dim Pth_Modèle As New GraphicsPath
    Select Case Modèle
      Case "R" 'Rectangle Vertical
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y)
        Dim B As New Point(Cell_B.X - WHquart, Cell_B.Y)
        Dim C As New Point(Cell_C.X - WHquart, Cell_C.Y)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y)
        With Pth_Modèle
          .StartFigure()
          .AddPolygon({A, B, C, D, A})
          .CloseFigure()
        End With
      Case "S" 'Carré Central (Utilisé dans le centre des régions)
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y + WHquart)
        Dim B As New Point(Cell_B.X - WHquart, Cell_B.Y + WHquart)
        Dim C As New Point(Cell_C.X - WHquart, Cell_C.Y - WHquart)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y - WHquart)
        With Pth_Modèle
          .StartFigure()
          .AddPolygon({A, B, C, D, A})
          .CloseFigure()
        End With

      Case "P" 'Pouce Bas
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y + WHquart)
        Dim B As New Point(Cell_B.X - WHquart, Cell_B.Y + WHquart)
        Dim C As New Point(Cell_A.X + WHquart + WHhalf, Cell_C.Y)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y)
        Dim AD As New Point(Cell_A.X + WHquart, Cell_A.Y + WHhalf)
        Dim BC As New Point(Cell_A.X + WHquart + WHhalf, Cell_A.Y + WHhalf)
        Dim Rect_A_m As New Rectangle(A.X, A.Y, WHhalf, WHhalf)
        Dim A1 As New Point(Cell_A.X + WHquart, Cell_A.Y + WHhalf)
        Dim A2 As New Point(Cell_A.X + WHhalf, Cell_A.Y + WHquart)
        Dim A3 As New Point(Cell_A.X + WHhalf + WHquart, Cell_A.Y + WHhalf)
        With Pth_Modèle
          .StartFigure()
          Select Case Plcy_Format_DAB
            Case 0       ' Droits 
              .AddPolygon({A, B, C, D, A})
            Case 1, 3, 5 ' Arrondis
              .AddArc(Rect_A_m, 180, 180)
              .AddLine(BC, C)
              .AddLine(C, D)
              .AddLine(D, AD)
            Case 2, 4, 6 ' Biseautés
              .AddPolygon({A1, A2, A3, C, D, A1})
          End Select
          .CloseFigure()
        End With
      Case "C" 'Coin Haut Gauche
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y + WHquart)
        Dim B As New Point(Cell_B.X, Cell_B.Y + WHquart)
        Dim BC As New Point(Cell_C.X, Cell_C.Y - WHquart)
        Dim C As New Point(Cell_C.X - WHquart, Cell_C.Y - WHquart)
        Dim CD As New Point(Cell_C.X - WHquart, Cell_C.Y)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y)
        Dim Rect_A_g As New Rectangle(A.X, A.Y, WH + WHhalf, WH + WHhalf)
        Dim Rect_C_p As New Rectangle(C.X, C.Y, WHhalf, WHhalf)
        Dim A1 As New Point(Cell_A.X + WHquart, Cell_A.Y + WHhalf + WHquart)
        Dim A2 As New Point(Cell_A.X + WHhalf + WHquart, Cell_A.Y + WHquart)

        With Pth_Modèle
          .StartFigure()
          Select Case Plcy_Format_DAB
            Case 0         ' Droits 
              .AddPolygon({A, B, BC, C, CD, D, A})
            Case 1, 3, 5   ' Arrondis
              'Un arc est compris dans un rectangle
              'Un rectangle est défini par un Left-Top (x-y) et Width-Height
              'L'angle Start à 0 sur l'axe x et tourne dans le sens des aiguilles d'une montre
              '                90, 180 , 270
              'la longueur de l'angle est de 90° dans un sens et -90° dans l'autre                
              .AddArc(Rect_A_g, 180, 90)
              .AddLine(B, BC)
              .AddArc(Rect_C_p, 270, -90)
              .AddLine(CD, D)
            Case 2, 4, 6  ' Biseautés
              .AddPolygon({A1, A2, B, BC, CD, D, A1})
          End Select
          .CloseFigure()
        End With
      Case Else
    End Select
    Return Pth_Modèle
  End Function
  Public Function G0_MdE_Build(Cellule As Integer, Modèle As String) As GraphicsPath
    'Construction des Modèles Elémentaires
    ' Ces modèles sont construits à/p des Modèles Primaires à une lettre avec RotateAt
    ' Les modèles Elémentaires ont 2 ou 3 lettres.
    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    Dim Pth_Modèle As GraphicsPath = Nothing

    Select Case Modèle
      Case "RV" 'Rectangle Vertical
        Pth_Modèle = G0_MdP_Build(Cellule, "R")

      Case "RH" 'Rectangle Horizontal
        Pth_Modèle = G0_MdP_Build(Cellule, "R")
        Using mat As New Matrix
          mat.RotateAt(90, New PointF(sc.Position_Center.X, sc.Position_Center.Y))
          Pth_Modèle.Transform(mat)
        End Using

      Case "SQ" 'Carré central
        Pth_Modèle = G0_MdP_Build(Cellule, "S")

      Case "CHG", "CHD", "CBD", "CBG" 'Coin Haut/Bas Gauche/Droit
        Dim Angle As Integer = 0
        If Modèle = "CHG" Then Angle = 0
        If Modèle = "CHD" Then Angle = 90
        If Modèle = "CBD" Then Angle = 180
        If Modèle = "CBG" Then Angle = 270
        Pth_Modèle = G0_MdP_Build(Cellule, "C")
        Using mat As New Matrix
          mat.RotateAt(Angle, New PointF(sc.Position_Center.X, sc.Position_Center.Y))
          Pth_Modèle.Transform(mat)
        End Using

      Case "PB", "PG", "PH", "PD" 'Pouce Haut/Bas Gauche/Droit
        Dim Angle As Integer = 0
        If Modèle = "PB" Then Angle = 0
        If Modèle = "PG" Then Angle = 90
        If Modèle = "PH" Then Angle = 180
        If Modèle = "PD" Then Angle = 270
        Pth_Modèle = G0_MdP_Build(Cellule, "P")
        Using mat As New Matrix
          mat.RotateAt(Angle, New PointF(sc.Position_Center.X, sc.Position_Center.Y))
          Pth_Modèle.Transform(mat)
        End Using
      Case Else
    End Select
    Return Pth_Modèle
  End Function
#End Region
End Module