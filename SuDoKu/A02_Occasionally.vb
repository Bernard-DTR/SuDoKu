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
    Frm_SDK.Mnu08_Résolution.Text = "Résolution SDK _ Stratégies " & Stg_Profondeur

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
      .BackColor = Color_Fond_Typ_I
    End With

    ' les infos B_*
    B_Top = Barre_Menu_Hauteur + Barre_Outils_Hauteur + Bld_Marge_LT + Bld_WH_Grid + Bld_Marge_LT
    With Frm_SDK
      .B_Famille.Location = New Point(B_Width, B_Top)
      .B_Famille.Size = New Size(20, B_Height)

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
    Build_Quadrillage()
  End Sub

  Public Sub OC_Présentation_Menu()

    ' Colorisation de 4 options de menu
    With Frm_SDK
      .Mnu_EDI_Saisir_Valeur.BackColor = Color_Fond_Typ_RV
      .Mnu_EDI_Val_Normale.BackColor = Color_Fond_Typ_RV
      .Mnu_EDI_Val_Initiale.BackColor = Color_Fond_Typ_I
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
    WHhalf = (WH \ 2)
    WHthird = (WH \ 3)
    WHquart = (WH \ 4)
    Bld_WH_Grid = (WH * 9) + 3 + 1 + 1 + 3 + 1 + 1 + 3 + 1 + 1 + 3

    ' La hauteur de la Barre_Outils est standard (25)  
    Barre_Outils_Hauteur = SystemInformation.ToolWindowCaptionHeight ' Hauteur de la Barre d'outils d'un formulaire

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
    '#736
    For i As Integer = 0 To 8
      Sqr_Celx(i) = Sqr_Cel(i + 9).X
      Sqr_Cely(i) = Sqr_Cel(i * 9).Y
    Next

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

    '#714
    ' Calcul de ces mêmes Sqr_Cdd_Inf avec Inflate utilisés lors des clics
    For s As Integer = 0 To 809
      Dim r As Rectangle = Sqr_Cdd(s)
      'Agrandissement du rectangle (1 pixel autour)
      r.Inflate(1, 1)
      Sqr_Cdd_Inf(s) = r
    Next s

    ' Calcul des Squares Path à Coins Arrondis  
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
              .AddArc(New Rectangle(x, y, WHhalf, WHhalf), 180, 90)
              .AddLine(x + WHhalf, y, x + WH, y)
              .AddLine(x + WH, y, x + WH, y + WH)
              .AddLine(x + WH, y + WH, x, y + WH)
              .AddLine(x, y + WH, x, y + WHhalf)
              .CloseFigure() 'La figure est déjà fermée
            End With
          Case 2, 5, 8, 29, 32, 35, 56, 59, 62       ' Coin HD arrondi
            With Sqr_Pth(k)
              .StartFigure()
              .AddLine(x, y, x + WH - WHhalf, y)
              .AddArc(New Rectangle(x + WH - WHhalf, y, WHhalf, WHhalf), 270, 90)
              .AddLine(x + WH, y + WHhalf, x + WH, y + WH)
              .AddLine(x + WH, y + WH, x, y + WH)
              .AddLine(x, y + WH, x, y)
              .CloseFigure()
            End With
          Case 18, 21, 24, 45, 48, 51, 72, 75, 78    ' Coin BG arrondi
            With Sqr_Pth(k)
              .StartFigure()
              .AddLine(x, y, x + WH, y)
              .AddLine(x + WH, y, x + WH, y + WH)
              .AddLine(x + WH, y + WH, x + WHhalf, y + WH)
              .AddArc(New Rectangle(x, y + WH - WHhalf, WHhalf, WHhalf), 90, 90)
              .AddLine(x, y + WH - WHhalf, x, y)
              .CloseFigure()
            End With
          Case 20, 23, 26, 47, 50, 53, 74, 77, 80    ' Coin BD arrondi
            With Sqr_Pth(k)
              .StartFigure()
              .AddLine(x, y, x + WH, y)
              .AddLine(x + WH, y, x + WH, y + WH - WHhalf)
              .AddArc(New Rectangle(x + WH - WHhalf, y + WH - WHhalf, WHhalf, WHhalf), 0, 90)
              .AddLine(x + WH - WHhalf, y + WH, x, y + WH)
              .AddLine(x, y + WH, x, y)
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

  Public Sub OC_Grid_Compute_Font_Size()
    'Calcul de la taille des polices des Valeurs et des Candidats
    Font_Val_Size = CalculateFontSize(Font_Name_ValCdd, FontStyle.Regular, WH, "1")
    Font_Cdd_Size_Zoom = (4 * Font_Val_Size) / 6
    Font_Cdd_Size = CalculateFontSize(Font_Name_ValCdd, FontStyle.Italic, CSng(WH / 3), "1")
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