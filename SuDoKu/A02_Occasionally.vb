Imports System.Drawing.Drawing2D
Imports System.Text

' Date de création: 28/07/2024
Friend Module A02_Occasionally
  Public Sub OC_Thèmes_Couleurs(Thème As Integer)
    Obj_Symbol = My.Settings.Obj_Symbol
    Obj_Forme = My.Settings.Obj_Forme

    Select Case Thème
      Case 0
        Color_Frm_BackColor = Color.FromArgb(255, 216, 245, 216)    ' Couleur Fond du formulaire et du Grid
        Color_Trait = Color.Green                                   ' Couleur des traits du Grid
        Clr_Fnd_VI = Color.FromArgb(255, 192, 255, 192)       ' Couleur Fond Valeurs Initiales
        Clr_Fnd_VCdd = Color.FromArgb(255, 129, 224, 129)      ' Couleur Fond Cellule Remplie/Vide 
        Color_Stratégique = Color.FromArgb(128, 15, 196, 101)       ' Couleur Couche Stratégique
        Clr_VI = Color.Green                                      ' Couleur des valeurs initiales
        Clr_VCdd = Color.Blue                                     ' Couleur des valeurs et des candidats
      Case 1
        Color_Frm_BackColor = Color.FromArgb(255, 216, 245, 242)
        Color_Trait = Color.Blue
        Clr_Fnd_VI = Color.FromArgb(255, 192, 255, 250)
        Clr_Fnd_VCdd = Color.FromArgb(255, 128, 193, 225)
        Color_Stratégique = Color.FromArgb(128, 9, 89, 149)
        Clr_VI = Color.Fuchsia
        Clr_VCdd = Color.Purple
      Case 2
        Color_Frm_BackColor = Color.Beige
        Color_Trait = Color.FromArgb(255, 255, 204, 0)
        Clr_Fnd_VI = Color.FromArgb(255, 217, 179, 179) '    Color.FromArgb(255, 184, 180, 131)
        Clr_Fnd_VCdd = Color.FromArgb(255, 189, 185, 138)
        Color_Stratégique = Color.FromArgb(128, 162, 100, 94)
        Clr_VI = Color.Navy
        Clr_VCdd = Color.Red
      Case Else 'Identique à Standard
        Color_Frm_BackColor = Color.FromArgb(255, 216, 245, 216)
        Color_Trait = Color.Green
        Clr_Fnd_VI = Color.FromArgb(255, 192, 255, 192)
        Clr_Fnd_VCdd = Color.FromArgb(255, 129, 224, 129)
        Color_Stratégique = Color.FromArgb(128, 15, 196, 101)
        Clr_VI = Color.Green
        Clr_VCdd = Color.Blue
    End Select
  End Sub

  Sub OC_Plcy_Stg_UOBTXYSJZKQ()
    ' 1 à/p de Str_List, il faut déterminer le nombre de stratégies concernées par la Production-Résolution
    Dim Stg_Nb As Integer = 0
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Prd = "O" Then Stg_Nb += 1
    Next Stg
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
    Frm_SDK.Mnu08_Résolution.Text = "Résolution SDK Stratégies " & Stg_Profondeur
    Frm_SDK.Mnu08_Automate.Text = "Automate (Stratégies " & Stg_Profondeur & ")"

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
    OC_Grid_Compute_Font_Size()
    OC_Grid_Cutting_Image()
    OC_Présentation_SDK()
    OC_Présentation_Menu()
    Build_Bmp_QFVS()
  End Sub
  Public Sub OC_Présentation_SDK()
    Dim Int_Seize As Integer = 16
    'Les infos B_*
    Dim B_Width As Integer = Bld_Marge_LT
    Dim B_Top As Integer
    Dim B_Height As Integer = 20 'Hauteur des zones B_*
    Frm_SDK.BarreOutils.Visible = True
    Frm_SDK.Journal.Visible = True

    Frm_SDK.BackColor = Color_Frm_BackColor

    ' Dimensions du Formulaire Frm_SDK
    With Frm_SDK
      .Width = Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT + Bld_Journal_Width + Bld_Marge_LT + Int_Seize
      .Height = SystemInformation.CaptionHeight + Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT + Bld_WH_Grid _
              + Bld_Marge_LT + B_Height + Bld_Marge_LT + Int_Seize
    End With

    ' L'emplacement du journal dépend de l'affichage ou non de la barre d'outils
    With Frm_SDK.Journal
      .Location = New Point(Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT, Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT)
      .Size = New Drawing.Size(Bld_Journal_Width, Bld_WH_Grid + Bld_Marge_LT + B_Height)
      .BackColor = Clr_Fnd_VI
    End With

    ' les infos B_*
    B_Top = Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT
    With Frm_SDK
      .B_Famille.Location = New Point(B_Width, B_Top)
      .B_Famille.Size = New Size(20, B_Height)
      '.B_Famille.Size = New Size(40, B_Height)

      B_Width = Bld_Marge_LT + .B_Famille.Width
      .B_Position.Location = New Point(B_Width, B_Top)
      .B_Position.Size = New Size(60, B_Height)

      B_Width = Bld_Marge_LT + .B_Famille.Width + .B_Position.Width
      .B_Pourcentage.Location = New Point(B_Width, B_Top)
      .B_Pourcentage.Size = New Size(40, B_Height)

      B_Width = Bld_Marge_LT + .B_Famille.Width + .B_Position.Width + .B_Pourcentage.Width
      .B_Info.Location = New Point(B_Width, B_Top)
      .B_Info.Size = New Size(Bld_Marge_LT + Bld_WH_Grid - B_Width, B_Height)
      .B_Info.Visible = True
      ' La progress Bar se substitue à la zone B_Info dans les traitements longues durées 
      .B_ProgressBar.Location = New Point(B_Width, B_Top)
      .B_ProgressBar.Size = New Size(Bld_Marge_LT + Bld_WH_Grid - B_Width, B_Height)
      .B_ProgressBar.Visible = False
    End With
    Build_Bmp_Quadrillage() 'Exceptionnellement 
  End Sub

  Public Sub OC_Présentation_Menu()

    ' Colorisation de 4 options de menu
    With Frm_SDK
      .Mnu_EDI_Saisir_Valeur.BackColor = Clr_Fnd_VCdd
      .Mnu_EDI_Val_Normale.BackColor = Clr_Fnd_VCdd
      .Mnu_EDI_Val_Initiale.BackColor = Clr_Fnd_VI
    End With

    With Frm_SDK
      .Mnu08_TestA.Text = Msg_Read("MNU_0800A")
      .Mnu08_TestB.Text = Msg_Read("MNU_0800B")
      .Mnu08_TestC.Text = Msg_Read("MNU_0800C")
      .Mnu08_TestD.Text = Msg_Read("MNU_0800D")
      .Mnu08_TestE.Text = Msg_Read("MNU_0800E")
      .Mnu08_TestF.Text = Msg_Read("MNU_0800F")
      .Mnu08_TestG.Text = Msg_Read("MNU_0800G")
      .Mnu08_TestH.Text = Msg_Read("MNU_0800H")
      .Mnu08_TestI.Text = Msg_Read("MNU_0800I")
      .Mnu08_TestJ.Text = Msg_Read("MNU_0800J")
    End With

    ' Libellé des Menus en fonction de la taille de WH
    Dim l As Integer
    With Frm_SDK
      Select Case WH
        Case 40 To 49 : l = 3
        Case 50 To 100 : l = 10
      End Select
      'PadRight  permet d'avoir un libellé de menu inférieur à 10 caractères
      'Substring permet d'avoir un menu court et un menu long
      .Mnu01.Text = Msg_Read("MNU_01000").PadRight(l).Substring(0, l)
      .Mnu02.Text = Msg_Read("MNU_02000").PadRight(l).Substring(0, l)
      .Mnu03.Text = Msg_Read("MNU_03000").PadRight(l).Substring(0, l)
      .Mnu04.Text = Msg_Read("MNU_04000").PadRight(l).Substring(0, l)
      .Mnu08.Text = Msg_Read("MNU_08000").PadRight(l).Substring(0, l)
      .Mnu05.Text = Msg_Read("MNU_05000").PadRight(l).Substring(0, l)
      .Mnu06.Text = Msg_Read("MNU_06000").PadRight(l).Substring(0, l)
      .Mnu07.Text = Msg_Read("MNU_07000").PadRight(l).Substring(0, l)
    End With

    ' Changement de la police dans Mnu de Frm_Sdk et Mnu04
    '                         dans Mnu_Cel
    '                         dans Mnu_EDI
    '                         dans Mnu_Journal
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
    WHh = WH \ 2
    WHt = WH \ 3
    WHq = WH \ 4
    Bld_WH_Grid = (WH * 9) + 3 + 1 + 1 + 3 + 1 + 1 + 3 + 1 + 1 + 3

    ' La hauteur de la Barre_Outils est standard (25)  
    Barre_Outils_Hauteur = SystemInformation.ToolWindowCaptionHeight ' Hauteur de la Barre d'outils d'un formulaire

    ' Calcul de Gz_tl As Point   
    ' Définition du Point Top-Left, Représente une paire ordonnée de coordonnées x et y entiers
    ' Ce point est utilisé comme origine pour le tracé du quadrillage et le positionnement des cellules
    ' Le point Gz_tl est calculé à partir de la marge gauche et de la hauteur cumulée de la barre de menu et de la barre d'outils
    Gz_tl = New Point(Bld_Marge_LT,
                      Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT)
    'Gz_tl est un point fixe  (5, 60)

    ' Calcul des traits du quadrillage et des positions des cellules
    ' Traitement lourd ... mais juste pour le quadrillage droit et le quadrillage arrondi
    ' Calcul des positions des traits de la grille
    ' Calcul de la longueur d'un trait
    Gz_traits(0) = 2                         ' Epais 
    Gz_traits(1) = 1 + Gz_traits(0) + WH + 1 ' Fin
    Gz_traits(2) = 0 + Gz_traits(1) + WH + 1 ' Fin
    Gz_traits(3) = 0 + Gz_traits(2) + WH + 2 ' Epais
    Gz_traits(4) = 1 + Gz_traits(3) + WH + 1 ' Fin
    Gz_traits(5) = 0 + Gz_traits(4) + WH + 1 ' Fin
    Gz_traits(6) = 0 + Gz_traits(5) + WH + 2 ' Epais
    Gz_traits(7) = 1 + Gz_traits(6) + WH + 1 ' Fin
    Gz_traits(8) = 0 + Gz_traits(7) + WH + 1 ' Fin
    Gz_traits(9) = 0 + Gz_traits(8) + WH + 2 ' Epais

    ' Calcul des 81 Sqr_cel
    Gz_Cellxy(0) = 2 + Gz_traits(0)
    Gz_Cellxy(1) = 1 + Gz_traits(1)
    Gz_Cellxy(2) = 1 + Gz_traits(2)
    Gz_Cellxy(3) = 2 + Gz_traits(3)
    Gz_Cellxy(4) = 1 + Gz_traits(4)
    Gz_Cellxy(5) = 1 + Gz_traits(5)
    Gz_Cellxy(6) = 2 + Gz_traits(6)
    Gz_Cellxy(7) = 1 + Gz_traits(7)
    Gz_Cellxy(8) = 1 + Gz_traits(8)
    For row As Integer = 0 To 8
      For col As Integer = 0 To 8
        Dim cellule As Integer = (row * 9) + col
        Dim x1 As Integer = Gz_tl.X + Gz_Cellxy(col)
        Dim y1 As Integer = Gz_tl.Y + Gz_Cellxy(row)
        Sqr_Cel(cellule) = New Rectangle(x1, y1, WH, WH)
      Next col
    Next row
    Dim R As Integer = Rayon 'Pour alléger la lecture du code 
    Dim D As Integer = R * 2

    ' Calcul des 9 paths des régions arrondies
    ' il est possible de faire varier le rayon des régions arrondies
    Dim region As Integer = 0

    For block_row As Integer = 0 To 2
      For block_col As Integer = 0 To 2
        Dim x1 As Integer = Gz_tl.X + Gz_traits(block_col * 3)
        Dim x2 As Integer = Gz_tl.X + Gz_traits(block_col * 3 + 3)
        Dim y1 As Integer = Gz_tl.Y + Gz_traits(block_row * 3)
        Dim y2 As Integer = Gz_tl.Y + Gz_traits(block_row * 3 + 3)

        Dim pth As New GraphicsPath()
        pth.AddArc(x1, y1, D, D, 180, 90)
        pth.AddArc(x2 - D, y1, D, D, 270, 90)
        pth.AddArc(x2 - D, y2 - D, D, D, 0, 90)
        pth.AddArc(x1, y2 - D, D, D, 90, 90)
        pth.CloseFigure()
        Region_Path(region) = pth
        region += 1
      Next
    Next

    ' Calcul des 81 sqr_pth des cellules

    For row As Integer = 0 To 8
      For col As Integer = 0 To 8
        Dim cellule As Integer = row * 9 + col
        Dim rct As Rectangle = Sqr_Cel(cellule)
        Dim pth As New GraphicsPath()
        ' soit le trait est tracé, soit le coin arrondi
        Dim TL As Boolean = (row Mod 3 = 0 And col Mod 3 = 0) ' Coin en haut à gauche
        Dim TR As Boolean = (row Mod 3 = 0 And col Mod 3 = 2) ' Coin en haut à droite
        Dim BL As Boolean = (row Mod 3 = 2 And col Mod 3 = 0) ' Coin en bas à gauche
        Dim BR As Boolean = (row Mod 3 = 2 And col Mod 3 = 2) ' Coin en bas à droite
        If TL Then
          pth.AddArc(rct.X, rct.Y, D, D, 180, 90)
        Else
          pth.AddLine(rct.X, rct.Y, rct.X + R, rct.Y)
        End If
        If TR Then
          pth.AddArc(rct.Right - D, rct.Y, D, D, 270, 90)
        Else
          pth.AddLine(rct.Right - R, rct.Y, rct.Right, rct.Y)
        End If
        If BR Then
          pth.AddArc(rct.Right - D, rct.Bottom - D, D, D, 0, 90)
        Else
          pth.AddLine(rct.Right, rct.Bottom - R, rct.Right, rct.Bottom)
        End If
        If BL Then
          pth.AddArc(rct.X, rct.Bottom - D, D, D, 90, 90)
        Else
          pth.AddLine(rct.X, rct.Bottom, rct.X, rct.Bottom - R)
        End If
        pth.CloseFigure()
        Sqr_Pth(cellule) = pth
      Next
    Next

    '#736
    For i As Integer = 0 To 8
      Sqr_Celx(i) = Sqr_Cel(i + 9).X
      Sqr_Cely(i) = Sqr_Cel(i * 9).Y
    Next

    ' Calcul des 9 Squares de chaque candidat   1 2 3  ou  7 8 9
    '                                           4 5 6      4 5 6
    '                                           7 8 9      1 2 3
    ' Identification du square de chaque candidat : Cellule * 10  + Candidat 
    Dim k As Integer = 0

    For j As Integer = 0 To 8          ' Axe Y (vertical)
      For i As Integer = 0 To 8        ' Axe X (horizontal)
        Dim baseX As Integer = Sqr_Cel(k).X
        Dim baseY As Integer = Sqr_Cel(k).Y
        ' 9 sous-carrés = 3×3
        For dy As Integer = 0 To 2
          For dx As Integer = 0 To 2
            Dim idx As Integer = (k * 10) + (dy * 3 + dx + 1)
            Sqr_Cdd(idx) = New Rectangle(baseX + dx * WHt, baseY + dy * WHt, WHt, WHt)
          Next dx
        Next dy
        k += 1
      Next i
    Next j

    ' Calcul de ces mêmes Sqr_Cdd_Inf avec Inflate utilisés  
    For s As Integer = 0 To 809
      Dim Rct_s As Rectangle = Sqr_Cdd(s)
      'diminution du rectangle (-1 pixel autour)
      Rct_s.Inflate(-1, -1)
      Sqr_Cdd_Inf(s) = Rct_s
    Next s

    ' Construction des 10 images correspondantes aux polices substituées
    ' Elles sont calculées quelque soit Plcy_Fantasy et indépendamment de WH
    Using font As New Font(Fnt_Name_ValCdd, 12),
          brsh As New SolidBrush(Color.Black)      'La couleur Color.Black pour les menus contextuels et les filtres de la BO
      For i As Integer = 0 To 9
        Dim Bm_MnuBO As New Bitmap(22, 22)
        Using g As Graphics = Graphics.FromImage(image:=Bm_MnuBO)
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
      Dim Pt_Origine As New Point(Sqr_Cel(Cellule).X, Sqr_Cel(Cellule).Y)
      Dim Pt_Left As Integer = Pt_Origine.X + 1
      Dim Pt_Top As Integer = Pt_Origine.Y + 1
      Dim Pt_Right As Integer = Sqr_Cel(Cellule).Right - 3
      Dim Pt_Bottom As Integer = Sqr_Cel(Cellule).Bottom - 3
      Dim PA0, PB0, PC0, PD0, PA1, PB1, PC1, PD1, PA2, PB2, PC2, PD2, PA3, PB3, PC3, PD3 As Point
      Dim PE0, PF0, PG0, PH0 As Point

      Dim D1_4 As Integer = ((1 * WH) \ 4)
      Dim D2_4 As Integer = 2 * D1_4
      Dim D3_4 As Integer = 3 * D1_4

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

  Public Sub OC_Grid_Compute_Font_Size()
    'Calcul de la taille des polices des Valeurs et des Candidats
    Fnt_Val_Size = CalculateFontSize(Fnt_Name_ValCdd, FontStyle.Regular, WH, "1")
    Fnt_Cdd_Size = CalculateFontSize(Fnt_Name_ValCdd, FontStyle.Regular, CSng(WH / 3), "1")
    Fnt_Val.Dispose()
    Fnt_Val = New Font(Fnt_Val.FontFamily, Fnt_Val_Size, Fnt_Val.Style)
    Fnt_Cdd.Dispose()
    Fnt_Cdd = New Font(Fnt_Cdd.FontFamily, Fnt_Cdd_Size, Fnt_Cdd.Style)
  End Sub
  Private Function CalculateFontSize(fontName As String,
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

      ' --- 2) Créer une version carrée de l'image ---
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
    ' Appliquer la transparence à tous les pixels d'une image de manière rapide
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

#End Region
End Module