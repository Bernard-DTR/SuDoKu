Imports System.Drawing.Drawing2D
Module M03_Paint_Quadrillage
#Region "G1 Couche Quadrillage"
  Public Sub Build_Bmp_Quadrillage()
    Bmp_Quadrillage = New Bitmap(Frm_SDK.Width, Frm_SDK.Height)
    Using g As Graphics = Graphics.FromImage(Bmp_Quadrillage)
      g.SmoothingMode = SmoothingMode.AntiAlias
      g.InterpolationMode = InterpolationMode.NearestNeighbor
      g.PixelOffsetMode = PixelOffsetMode.None
      g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit

      G1_Grid_Paint(g)
    End Using
  End Sub
  Public Sub Build_Bmp_Fonds()
    'La procédure n'est appelée qu'une fois à chaque nouveau jeu
    'Le fond des cellules est peint, ainsi que les valeurs des cellules Initiales 
    Bmp_Fond = New Bitmap(Frm_SDK.Width, Frm_SDK.Height, Imaging.PixelFormat.Format32bppPArgb)
    Bmp_Fond.SetResolution(96, 96)
    Using g As Graphics = Graphics.FromImage(Bmp_Fond)
      g.SmoothingMode = SmoothingMode.None
      g.InterpolationMode = InterpolationMode.NearestNeighbor
      g.PixelOffsetMode = PixelOffsetMode.None
      g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit
      Dim sc As New Cellule_Cls
      ' Tous les fonds de cellule sont peints, ainsi que les valeurs initiales
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G2_Cellule_Paint_Fond(g)

        If U(i, 1) <> " " Then
          g.DrawString(Subst_Police(U(i, 1)),
                       Fnt_Val,
                       Brh_VI,
                       Sqr_Cel(i).X + WHhalf, Sqr_Cel(i).Y + WHhalf, Format_Center)
        End If
      Next
    End Using
  End Sub
  Public Sub Build_Bmp_Valeurs()
    ' Seules les valeurs des cellules Remplies sont peintes
    ' Cette fonction est optimisée car elle est appelée à chaque changement de valeur d'une cellule Remplie
    Bmp_Valeur = New Bitmap(Frm_SDK.Width, Frm_SDK.Height, Imaging.PixelFormat.Format32bppPArgb)
    Bmp_Valeur.SetResolution(96, 96)
    Using g As Graphics = Graphics.FromImage(Bmp_Valeur)
      g.SmoothingMode = SmoothingMode.None
      g.InterpolationMode = InterpolationMode.NearestNeighbor
      g.PixelOffsetMode = PixelOffsetMode.None
      g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit
      For i As Integer = 0 To 80
        If U(i, 1) = " " AndAlso U(i, 2) <> " " Then
          g.DrawString(Subst_Police(U(i, 2)),
                       Fnt_Val,
                       Brh_VCdd,
                       Sqr_Cel(i).X + WHhalf, Sqr_Cel(i).Y + WHhalf, Format_Center)
        End If
        If Plcy_Dernière_Valeur_Unité AndAlso U_dv(i) Then
          G0_Cell_Figure(g, i, "Ellipse", Color_Stratégique)
        End If
      Next
    End Using
  End Sub
  Public Sub Build_Bmp_Saisie()
    ' le Bmp_Fond_Saisie est calculé une seule fois dans Frm_SDK_Load
    Bmp_Fond_Saisie = New Bitmap(WH, WH)
    Dim r As New Rectangle(0, 0, WH, WH)
    r.Inflate(-1, -1)
    Using g As Graphics = Graphics.FromImage(Bmp_Fond_Saisie)
      g.SmoothingMode = SmoothingMode.None
      g.InterpolationMode = InterpolationMode.NearestNeighbor
      g.PixelOffsetMode = PixelOffsetMode.None
      g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit
      'Il n'est pas possible d'utiliser Sqr_Cdd( de 1 à 9), mais un bitmap plus large
      Using brsh As New SolidBrush(Color_Cell_Select)
        g.FillRectangle(brsh, r)
        For cdd As Integer = 1 To 9
          ' Sqr_Cel et Sqr_Cdd ne sont pas référencés par Cellule
          Dim row As Integer = (cdd - 1) \ 3
          Dim col As Integer = (cdd - 1) Mod 3
          Dim rc As New Rectangle(r.X + col * WHthird, r.Y + row * WHthird, WHthird, WHthird)
          g.DrawString(Subst_Police(CStr(cdd)), Fnt_Cdd, Brh_VCdd, rc, Format_Center)
        Next cdd
      End Using
      ' Définition des points et du style de trait
      Dim dashPattern As Single() = {1, 5}
      Using pen As New Pen(Color_Trait, Bld_Trait_1 \ 10)
        pen.DashPattern = dashPattern
        Dim x1 As Integer = WHthird
        Dim x2 As Integer = (2 * WHthird)
        g.DrawLine(pen, x1, 0, x1, WH)
        g.DrawLine(pen, x2, 0, x2, WH)
        ' Lignes horizontales
        Dim y1 As Integer = WHthird
        Dim y2 As Integer = +(2 * WHthird)
        g.DrawLine(pen, 0, y1, WH, y1)
        g.DrawLine(pen, 0, y2, WH, y2)
      End Using
    End Using
  End Sub
  Public Sub G1_Cell_Fond_Saisie(g As Graphics, Cellule As Integer)
    g.SmoothingMode = SmoothingMode.None
    g.InterpolationMode = InterpolationMode.NearestNeighbor
    g.PixelOffsetMode = PixelOffsetMode.None
    g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit
    Using brsh As New SolidBrush(Color_Cell_Select)
      g.FillRectangle(brsh, Sqr_Cel(Cellule))
      For cdd As Integer = 1 To 9
        Dim cd As Integer = Cellule * 10 + cdd
        If U(Cellule, 3).Contains(CStr(cdd)) Then
          g.DrawString(Subst_Police(CStr(cdd)), Fnt_Cdd, Brh_VCdd, Sqr_Cdd(cd), Format_Center)
        End If
      Next cdd
    End Using
    Dim X As Integer = Sqr_Cel(Cellule).X
    Dim Y As Integer = Sqr_Cel(Cellule).Y
    ' Définition des points et du style de trait
    Dim dashPattern As Single() = {1, 5}
    Using pen As New Pen(Color_Trait, Bld_Trait_1 \ 10)
      pen.DashPattern = dashPattern
      g.DrawLine(pen, X, Y + (WHthird * 1), X + WH, Y + (WHthird * 1))
      g.DrawLine(pen, X, Y + (WHthird * 2), X + WH, Y + (WHthird * 2))
      g.DrawLine(pen, X + (WHthird * 1), Y, X + (WHthird * 1), Y + WH)
      g.DrawLine(pen, X + (WHthird * 2), Y, X + (WHthird * 2), Y + WH)
    End Using
  End Sub

  Public Sub G1_Grid_Paint(g As Graphics)
    ' Cette fonction construit un seul et grand carré (taille de la grille) pour effacer l'ensemble de la grille
    ' et dessiner ensuite le quadrillage
    ' le fond du formulaire est Color_Frm_BackColor, il est défini dans Sub Présentation_SDK
    g.Clear(Color_Frm_BackColor)
    ' Efface l'intégralité de l'emplacement de la grille avec un carré unique
    Using brsh As New SolidBrush(Color_Frm_BackColor)
      g.FillRectangle(brsh, Gz_tl.X, Gz_tl.Y, Bld_WH_Grid, Bld_WH_Grid)
    End Using
    Select Case Plcy_Format_DAB
      Case 0 : G1_Grid_Paint_Quadrillage_droit(g)
      Case 1 : G1_Grid_Paint_Quadrillage_arrondi(g)
    End Select
  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_droit(g As Graphics)
    ' Tracer les traits
    For i As Integer = 0 To 9
      Dim y As Integer = Gz_tl.Y + Gz_traits(i)
      Dim x As Integer = Gz_tl.X + Gz_traits(i)

      Dim p As Pen = If(i Mod 3 = 0, Pen_épais, Pen_fin)

      g.DrawLine(p, Gz_tl.X, y, Gz_tl.X + Gz_Trait_Length, y)
      g.DrawLine(p, x, Gz_tl.Y, x, Gz_tl.Y + Gz_Trait_Length)
    Next
    '' tracer les rectangles des cellules (debug)
    'For i As Integer = 0 To 80
    '  'g.DrawRectangle(Pens.Green, Sqr_Cel(i))
    '  g.FillRectangle(Brushes.Green, Sqr_Cel(i))
    'Next
  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_arrondi(g As Graphics)
    ' Tracer les traits des régions arrondies
    For i As Integer = 0 To 8
      g.DrawPath(Pen_épais, Region_Path(i))
    Next
    ' Tracer les traits intérieurs fins
    For i As Integer = 1 To 8
      If i = 3 Then Continue For
      If i = 6 Then Continue For
      Dim y As Integer = Gz_tl.Y + Gz_traits(i)
      Dim x As Integer = Gz_tl.X + Gz_traits(i)

      'Dim p As Pen = If(i Mod 3 = 0, pen_épais, pen_fin)

      g.DrawLine(Pen_fin, Gz_tl.X, y, Gz_tl.X + Gz_Trait_Length, y)
      g.DrawLine(Pen_fin, x, Gz_tl.Y, x, Gz_tl.Y + Gz_Trait_Length)
    Next
    'For i As Integer = 0 To 80
    '  g.FillPath(Brushes.LightGreen, Sqr_Pth(i))
    '  'g.DrawPath(pen_fin, Sqr_Pth(i))
    'Next
  End Sub


  Public Sub Gz_Traits_Calcul()
    ' Calcul des positions des traits de la grille
    ' Calcul de la longueur d'un trait
    Dim pos As Integer = Trait_épais \ 2
    Gz_traits(0) = pos
    For i As Integer = 1 To 9
      If i Mod 3 = 0 Then           ' traits 3, 6, 9  
        pos += WH + Trait_épais
      Else                          ' traits 1, 2, 4, 5, 7, 8
        pos += WH + Trait_fin
      End If
      Gz_traits(i) = pos
    Next
    Gz_Trait_Length = Gz_traits(9) + (Trait_épais \ 2)
  End Sub

  Public Sub Gz_Sqr_Cel_Calcul()
    ' Calcul des 81 rectangles des cellules

    For row As Integer = 0 To 8
      For col As Integer = 0 To 8
        Dim cellule As Integer = (row * 9) + col
        Dim leftTraitWidth As Integer = If(col Mod 3 = 0, Trait_épais, Trait_fin)
        Dim rightTraitWidth As Integer = If((col + 1) Mod 3 = 0, Trait_épais, Trait_fin)
        Dim topTraitWidth As Integer = If(row Mod 3 = 0, Trait_épais, Trait_fin)
        Dim bottomTraitWidth As Integer = If((row + 1) Mod 3 = 0, Trait_épais, Trait_fin)

        Dim x1 As Integer = Gz_tl.X + Gz_traits(col) + leftTraitWidth \ 2
        Dim x2 As Integer = Gz_tl.X + Gz_traits(col + 1) - rightTraitWidth \ 2
        Dim y1 As Integer = Gz_tl.Y + Gz_traits(row) + topTraitWidth \ 2
        Dim y2 As Integer = Gz_tl.Y + Gz_traits(row + 1) - bottomTraitWidth \ 2
        Sqr_Cel(cellule) = New Rectangle(x1, y1, x2 - x1, y2 - y1)
      Next
    Next
  End Sub

  Public Sub Gz_Region_Path_Calcul()
    ' Calcul des 9 régions arrondies
    Dim region As Integer = 0

    For block_row As Integer = 0 To 2
      For block_col As Integer = 0 To 2
        Dim x1 As Integer = Gz_tl.X + Gz_traits(block_col * 3)
        Dim x2 As Integer = Gz_tl.X + Gz_traits(block_col * 3 + 3)
        Dim y1 As Integer = Gz_tl.Y + Gz_traits(block_row * 3)
        Dim y2 As Integer = Gz_tl.Y + Gz_traits(block_row * 3 + 3)

        Dim pth As New GraphicsPath()
        pth.AddArc(x1, y1, Rayon_region * 2, Rayon_region * 2, 180, 90)
        pth.AddArc(x2 - 2 * Rayon_region, y1, Rayon_region * 2, Rayon_region * 2, 270, 90)
        pth.AddArc(x2 - 2 * Rayon_region, y2 - 2 * Rayon_region, Rayon_region * 2, Rayon_region * 2, 0, 90)
        pth.AddArc(x1, y2 - 2 * Rayon_region, Rayon_region * 2, Rayon_region * 2, 90, 90)
        pth.CloseFigure()
        Region_Path(region) = pth
        region += 1
      Next
    Next
  End Sub
  Public Sub Gz_Sqr_Pth_Calcul()
    ' Calcul des 81 path des cellules

    For row As Integer = 0 To 8
      For col As Integer = 0 To 8

        Dim cellule As Integer = row * 9 + col
        Dim rct As Rectangle = Sqr_Cel(cellule)

        Dim pth As New GraphicsPath()

        Dim TL As Boolean = (row Mod 3 = 0 And col Mod 3 = 0) ' Coin en haut à gauche
        Dim TR As Boolean = (row Mod 3 = 0 And col Mod 3 = 2) ' Coin en haut à droite
        Dim BL As Boolean = (row Mod 3 = 2 And col Mod 3 = 0) ' Coin en bas à gauche
        Dim BR As Boolean = (row Mod 3 = 2 And col Mod 3 = 2) ' Coin en bas à droite
        If TL Then
          pth.AddArc(rct.X, rct.Y, 2 * Rayon_cellule, 2 * Rayon_cellule, 180, 90)
        Else
          pth.AddLine(rct.X, rct.Y, rct.X + Rayon_cellule, rct.Y)
        End If
        If TR Then
          pth.AddArc(rct.Right - 2 * Rayon_cellule, rct.Y, 2 * Rayon_cellule, 2 * Rayon_cellule, 270, 90)
        Else
          pth.AddLine(rct.Right - Rayon_cellule, rct.Y, rct.Right, rct.Y)
        End If
        If BR Then
          pth.AddArc(rct.Right - 2 * Rayon_cellule, rct.Bottom - 2 * Rayon_cellule, 2 * Rayon_cellule, 2 * Rayon_cellule, 0, 90)
        Else
          pth.AddLine(rct.Right, rct.Bottom - Rayon_cellule, rct.Right, rct.Bottom)
        End If
        If BL Then
          pth.AddArc(rct.X, rct.Bottom - 2 * Rayon_cellule, 2 * Rayon_cellule, 2 * Rayon_cellule, 90, 90)
        Else
          pth.AddLine(rct.X, rct.Bottom, rct.X, rct.Bottom - Rayon_cellule)
        End If
        pth.CloseFigure()
        Sqr_Pth(cellule) = pth
      Next
    Next
  End Sub

#End Region
End Module