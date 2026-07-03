Imports System.Drawing.Drawing2D
Module M03_Paint_Quadrillage
#Region "G1 Couche Quadrillage"
  Public Sub G1_Graphics_Généralité(g As Graphics)
    g.SmoothingMode = SmoothingMode.AntiAlias
    g.InterpolationMode = InterpolationMode.NearestNeighbor
    g.PixelOffsetMode = PixelOffsetMode.None
    g.TextRenderingHint = Text.TextRenderingHint.AntiAliasGridFit
  End Sub

  Public Sub G1_Grid_Paint(g As Graphics)
    ' Cette fonction construit un seul et grand carré (taille de la grille) pour effacer l'ensemble de la grille
    ' et dessiner ensuite le quadrillage
    ' le fond du formulaire est Color_Frm_BackColor, il est défini dans Sub Présentation_SDK
    G1_Graphics_Généralité(g)
    g.Clear(Color_Frm_BackColor)
    ' Efface l'intégralité de l'emplacement de la grille avec un carré unique
    Using brsh As New SolidBrush(Color_Frm_BackColor)
      g.FillRectangle(brsh, Gz_tl.X, Gz_tl.Y, Bld_WH_Grid, Bld_WH_Grid)
    End Using
    Select Case Plcy_Format_DAB
      Case 0 ' Quadrillage droit
        For i As Integer = 0 To 9
          Dim x As Integer = Gz_tl.X + Gz_traits(i)
          Dim y As Integer = Gz_tl.Y + Gz_traits(i)

          Dim p As Pen = If(i Mod 3 = 0, Pen_épais, Pen_fin)
          g.DrawLine(p, Gz_tl.X + 1, y, Gz_tl.X + Bld_WH_Grid, y)
          g.DrawLine(p, x, Gz_tl.Y + 1, x, Gz_tl.Y + Bld_WH_Grid)
        Next

      Case 1 ' Quadrillage arrondi
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
          g.DrawLine(Pen_fin, Gz_tl.X, y, Gz_tl.X + Bld_WH_Grid, y)
          g.DrawLine(Pen_fin, x, Gz_tl.Y, x, Gz_tl.Y + Bld_WH_Grid)
        Next

    End Select
  End Sub
  Public Sub G1_Cell_Fond_Saisie(g As Graphics, Cellule As Integer)
    'Plcy_Strg = "Sai" : Il s'agit d'afficher la grille de saisie SEULEMENT avec les candidats éligibles
    'Plcy_Strg = "CaG" : La grille m^me vide permet de saisir une valeur
    G1_Graphics_Généralité(g)
    If U_arr(Cellule) <> "SQ" Then
      g.FillPath(Brh_Select, Sqr_Pth(Cellule))
    Else
      g.FillRectangle(Brh_Select, Sqr_Cel(Cellule))
    End If

    For cdd As Integer = 1 To 9
      Dim cd As Integer = Cellule * 10 + cdd
      If U(Cellule, 3).Contains(CStr(cdd)) Then
        g.DrawString(Subst_Police(CStr(cdd)), Fnt_Cdd, Brh_VCdd, Sqr_Cdd(cd), Format_Center)
      End If
    Next cdd
    Build_Bmp_Saisie_Traits(g, Sqr_Cel(Cellule).X, Sqr_Cel(Cellule).Y)
  End Sub
  Public Sub Build_Bmp_QFVS()
    Build_Bmp_Quadrillage()
    Build_Bmp_Fonds()
    Build_Bmp_valeur_saisie()
    Build_Bmp_Saisie()
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
            If U_arr(i) <> "SQ" Then
              g.FillPath(brFond, Sqr_Pth(i))
            Else
              g.FillRectangle(brFond, Sqr_Cel(i))
            End If
          Else
            If U_arr(i) <> "SQ" Then
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
            g.DrawString(Subst_Police(U(i, 1)), Fnt_Val, Brh_VI,
                         Sqr_Cel(i), Format_Center)
          End If
        End Using
      Next
    End Using
  End Sub
  Public Sub Build_Bmp_valeur_saisie()
    ' Cette fonction est optimisée car elle est appelée à chaque changement de valeur d'une cellule Remplie
    ' Seules les valeurs des cellules Remplies sont peintes
    Bmp_Valeur = New Bitmap(Frm_SDK.Width, Frm_SDK.Height, Imaging.PixelFormat.Format32bppPArgb)
    Bmp_Valeur.SetResolution(96, 96)
    Using g As Graphics = Graphics.FromImage(Bmp_Valeur)
      G1_Graphics_Généralité(g)
      For i As Integer = 0 To 80
        If U(i, 1) = " " AndAlso U(i, 2) <> " " Then
          g.DrawString(Subst_Police(U(i, 2)), Fnt_Val, Brh_VCdd, Sqr_Cel(i), Format_Center)
        End If
        If Plcy_Dernière_Valeur_Unité AndAlso U_dv(i) Then
          G0_Cell_Figure(g, i, "Ellipse", Color_Stratégique)
        End If
      Next
    End Using
  End Sub

  Public Sub Build_Bmp_Saisie()
    Bmp_Fond_Saisie_SQ = Build_Bmp_Fond_Saisie("SQ")
    Bmp_Fond_Saisie_TL = Build_Bmp_Fond_Saisie("TL")
    Bmp_Fond_Saisie_TR = Build_Bmp_Fond_Saisie("TR")
    Bmp_Fond_Saisie_BL = Build_Bmp_Fond_Saisie("BL")
    Bmp_Fond_Saisie_BR = Build_Bmp_Fond_Saisie("BR")
  End Sub
  Public Sub Build_Bmp_Saisie_Traits(g As Graphics, x As Integer, y As Integer)
    Using pen As New Pen(Color_Trait, Trait_fin)
      pen.DashPattern = {1, 5}
      Dim x1 As Integer = x + WHt
      Dim x2 As Integer = x + (WHt * 2)
      Dim y1 As Integer = y + WHt
      Dim y2 As Integer = y + (WHt * 2)
      g.DrawLine(pen, x, y1, x + WH, y1)
      g.DrawLine(pen, x, y2, x + WH, y2)
      g.DrawLine(pen, x1, y, x1, y + WH)
      g.DrawLine(pen, x2, y, x2, y + WH)
    End Using
  End Sub

  Private Function Build_Bmp_Fond_Saisie(shape As String) As Bitmap
    Dim bmp As New Bitmap(WH, WH)
    Using g As Graphics = Graphics.FromImage(bmp)
      G1_Graphics_Généralité(g)
      Select Case shape
        Case "SQ" : g.FillRectangle(Brh_Select, 0, 0, WH, WH)
        Case "TL" : g.FillPath(Brh_Select, Build_Path_TL())
        Case "TR" : g.FillPath(Brh_Select, Build_Path_TR())
        Case "BL" : g.FillPath(Brh_Select, Build_Path_BL())
        Case "BR" : g.FillPath(Brh_Select, Build_Path_BR())
      End Select
      Build_Bmp_Saisie_Cdd(g, New Rectangle(0, 0, WH, WH))
      Build_Bmp_Saisie_Traits(g, 0, 0)
    End Using
    Return bmp
  End Function
  Public Sub Build_Bmp_Saisie_Cdd(g As Graphics, r As Rectangle)
    ' le Bmp_Fond_Saisie_SQ est complété des candidats et ds traits de séparation des candidats
    For cdd As Integer = 1 To 9
      ' Sqr_Cel et Sqr_Cdd ne sont pas référencés par Cellule
      Dim row As Integer = (cdd - 1) \ 3
      Dim col As Integer = (cdd - 1) Mod 3
      Dim rct As New Rectangle(r.X + (col * WHt), r.Y + (row * WHt), WHt, WHt)
      g.DrawString(Subst_Police(CStr(cdd)), Fnt_Cdd, Brh_VCdd, rct, Format_Center)
    Next cdd
  End Sub
  Private Function Build_Path_TL() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon
    Dim D As Integer = R * 2
    gp.AddArc(0, 0, D, D, 180, 90)
    gp.AddLine(R, 0, WH, 0)
    gp.AddLine(WH, 0, WH, WH)
    gp.AddLine(WH, WH, 0, WH)
    gp.CloseFigure()
    Return gp
  End Function

  Private Function Build_Path_TR() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon
    Dim D As Integer = R * 2
    gp.AddArc(WH - D, 0, D, D, 270, 90)
    gp.AddLine(WH, R, WH, WH)
    gp.AddLine(WH, WH, 0, WH)
    gp.AddLine(0, WH, 0, 0)
    gp.CloseFigure()
    Return gp
  End Function
  Private Function Build_Path_BL() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon
    Dim D As Integer = R * 2
    gp.AddLine(0, 0, WH, 0)
    gp.AddLine(WH, 0, WH, WH)
    gp.AddLine(WH, WH, R, WH)
    gp.AddArc(0, WH - D, D, D, 90, 90)
    gp.AddLine(0, WH - R, 0, 0)
    gp.CloseFigure()
    Return gp
  End Function
  Private Function Build_Path_BR() As GraphicsPath
    Dim gp As New GraphicsPath()
    Dim R As Integer = Rayon
    Dim D As Integer = R * 2
    gp.AddArc(WH - D, WH - D, D, D, 0, 90)
    gp.AddLine(WH - R, WH, 0, WH)
    gp.AddLine(0, WH, 0, 0)
    gp.AddLine(0, 0, WH, 0)
    gp.CloseFigure()
    Return gp
  End Function

#End Region
End Module