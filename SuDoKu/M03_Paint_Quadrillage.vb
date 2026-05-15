Imports System.Drawing.Drawing2D
Module M03_Paint_Quadrillage
#Region "G1 Couche Quadrillage"
  Public Sub G1_Graphics_Généralité(g As Graphics)
    g.SmoothingMode = SmoothingMode.AntiAlias
    g.InterpolationMode = InterpolationMode.NearestNeighbor
    g.PixelOffsetMode = PixelOffsetMode.None
    g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit
  End Sub
  Public Sub Build_Bmp_Quadrillage()
    Bmp_Quadrillage = New Bitmap(Frm_SDK.Width, Frm_SDK.Height)
    Using g As Graphics = Graphics.FromImage(Bmp_Quadrillage)
      G1_Graphics_Généralité(g)
      G1_Grid_Paint(g)
    End Using
  End Sub
  Public Sub Build_Bmp_Fonds()
    'La procédure n'est appelée qu'une fois à chaque nouveau jeu
    'Le fond des cellules est peint, ainsi que les valeurs des cellules Initiales 
    Bmp_Fond = New Bitmap(Frm_SDK.Width, Frm_SDK.Height, Imaging.PixelFormat.Format32bppPArgb)
    Bmp_Fond.SetResolution(96, 96)
    Using g As Graphics = Graphics.FromImage(Bmp_Fond)
      G1_Graphics_Généralité(g)
      ' Tous les fonds de cellule sont peints 
      For i As Integer = 0 To 80
        Dim fondCouleur As Boolean = (Plcy_Fond_Grille = 0)
        Using brFond As New SolidBrush(U_Clr_Cell_Fond(i))
          If fondCouleur Then
            If U_dab(i) Then
              g.FillPath(brFond, Sqr_Pth(i))
            Else
              g.FillRectangle(brFond, Sqr_Cel(i))
            End If
          Else
            If U_dab(i) Then
              g.ResetClip()
              g.SetClip(Sqr_Pth(i), CombineMode.Replace)
              g.DrawImage(Sqr_Img(i), Sqr_Cel(i))
              g.ResetClip()   ' ← AJOUT ESSENTIEL
            Else
              g.DrawImage(Sqr_Img(i), Sqr_Cel(i))
            End If
          End If
          ' et les valeurs initiales
          If U(i, 1) <> " " Then
            g.DrawString(Subst_Police(U(i, 1)),
                         Fnt_Val,
                         Brh_VI,
                         Sqr_Cel(i).X + WHhalf, Sqr_Cel(i).Y + WHhalf, Format_Center)
          End If
        End Using
      Next
    End Using
  End Sub

  Public Sub Build_Bmp_Valeurs()
    ' Seules les valeurs des cellules Remplies sont peintes
    ' Cette fonction est optimisée car elle est appelée à chaque changement de valeur d'une cellule Remplie
    Bmp_Valeur = New Bitmap(Frm_SDK.Width, Frm_SDK.Height, Imaging.PixelFormat.Format32bppPArgb)
    Bmp_Valeur.SetResolution(96, 96)
    Using g As Graphics = Graphics.FromImage(Bmp_Valeur)
      G1_Graphics_Généralité(g)
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
    Bmp_Fond_Saisie = BuildBmpFondSaisie("SQ")
    Bmp_Fond_Saisie_TL = BuildBmpFondSaisie("TL")
    Bmp_Fond_Saisie_TR = BuildBmpFondSaisie("TR")
    Bmp_Fond_Saisie_BL = BuildBmpFondSaisie("BL")
    Bmp_Fond_Saisie_BR = BuildBmpFondSaisie("BR")
  End Sub

  Private Function BuildBmpFondSaisie(shape As String) As Bitmap
    Dim bmp As New Bitmap(WH, WH)
    Using g As Graphics = Graphics.FromImage(bmp)
      G1_Graphics_Généralité(g)
      Select Case shape
        Case "SQ" : g.FillRectangle(Brsh_Select, 0, 0, WH, WH)
        Case "TL" : g.FillPath(Brsh_Select, BuildPathTL())
        Case "TR" : g.FillPath(Brsh_Select, BuildPathTR())
        Case "BL" : g.FillPath(Brsh_Select, BuildPathBL())
        Case "BR" : g.FillPath(Brsh_Select, BuildPathBR())
      End Select
      Build_Bmp_Saisie_Cdd_Traits(g, New Rectangle(0, 0, WH, WH))
    End Using
    Return bmp
  End Function

  Public Sub Build_Bmp_Saisie_Cdd_Traits(g As Graphics, r As Rectangle)
    ' le Bmp_Fond_Saisie est complété des candidats et ds traits de séparation des candidats
    For cdd As Integer = 1 To 9
      ' Sqr_Cel et Sqr_Cdd ne sont pas référencés par Cellule
      Dim row As Integer = (cdd - 1) \ 3
      Dim col As Integer = (cdd - 1) Mod 3
      Dim rct As New Rectangle(r.X + (col * WHthird), r.Y + (row * WHthird), WHthird, WHthird)
      g.DrawString(Subst_Police(CStr(cdd)), Fnt_Cdd, Brh_VCdd, rct, Format_Center)
    Next cdd
    ' Définition des points et du style de trait
    Dim dashPattern As Single() = {1, 5}
    Using pen As New Pen(Color_Trait, Bld_Trait_1)
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
  End Sub

  Public Sub G1_Cell_Fond_Saisie(g As Graphics, Cellule As Integer)
    G1_Graphics_Généralité(g)
    Using brsh As New SolidBrush(Color_Cell_Select)
      'If Plcy_Format_DAB = 0 Then g.FillRectangle(brsh, Sqr_Cel(Cellule))
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
    Using pen As New Pen(Color_Trait, Bld_Trait_1)
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
    G1_Graphics_Généralité(g)
    For i As Integer = 0 To 9
      Dim x As Integer = Gz_tl.X + Gz_traits(i)
      Dim y As Integer = Gz_tl.Y + Gz_traits(i)

      Dim p As Pen = If(i Mod 3 = 0, Pen_épais, Pen_fin)

      g.DrawLine(p, Gz_tl.X + 1, y, Gz_tl.X + Gz_Trait_Length, y)
      g.DrawLine(p, x, Gz_tl.Y + 1, x, Gz_tl.Y + Gz_Trait_Length)
    Next
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

      g.DrawLine(Pen_fin, Gz_tl.X, y, Gz_tl.X + Gz_Trait_Length, y)
      g.DrawLine(Pen_fin, x, Gz_tl.Y, x, Gz_tl.Y + Gz_Trait_Length)
    Next
  End Sub

  Public Sub Gz_Region_Path_Calcul()
    ' Calcul des 9 régions arrondies
    ' il est possible de faire varier le rayon des régions arrondies
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

  Public Sub Gz_Traits_Sqr_Cel_Calcul()
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
    Gz_Trait_Length = Gz_traits(9) + 2
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
  End Sub

  Private Function BuildPathTL() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon_cellule
    Dim D As Integer = R * 2
    ' Arc en haut-gauche
    gp.AddArc(0, 0, D, D, 180, 90)
    ' Haut
    gp.AddLine(R, 0, WH, 0)
    ' Droite
    gp.AddLine(WH, 0, WH, WH)
    ' Bas
    gp.AddLine(WH, WH, 0, WH)
    gp.CloseFigure()
    Return gp
  End Function

  Private Function BuildPathTR() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon_cellule
    Dim D As Integer = R * 2
    ' Arc en haut-droite
    gp.AddArc(WH - D, 0, D, D, 270, 90)
    ' Droite
    gp.AddLine(WH, R, WH, WH)
    ' Bas
    gp.AddLine(WH, WH, 0, WH)
    ' Gauche
    gp.AddLine(0, WH, 0, 0)
    gp.CloseFigure()
    Return gp
  End Function
  Private Function BuildPathBL() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon_cellule
    Dim D As Integer = R * 2
    ' 1. Haut (gauche → droite)
    gp.AddLine(0, 0, WH, 0)
    ' 2. Droite (haut → bas)
    gp.AddLine(WH, 0, WH, WH)
    ' 3. Bas (droite → arc)
    gp.AddLine(WH, WH, R, WH)
    ' 4. Arc en bas-gauche
    gp.AddArc(0, WH - D, D, D, 90, 90)
    ' 5. Gauche (arc → haut)
    gp.AddLine(0, WH - R, 0, 0)

    gp.CloseFigure()
    Return gp
  End Function
  Private Function BuildPathBR() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon_cellule
    Dim D As Integer = R * 2
    ' Arc en bas-droite
    gp.AddArc(WH - D, WH - D, D, D, 0, 90)
    ' Bas
    gp.AddLine(WH - R, WH, 0, WH)
    ' Gauche
    gp.AddLine(0, WH, 0, 0)
    ' Haut
    gp.AddLine(0, 0, WH, 0)
    gp.CloseFigure()
    Return gp
  End Function

#End Region
End Module