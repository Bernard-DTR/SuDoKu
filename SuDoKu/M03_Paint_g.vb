Option Strict On
Option Explicit On

Imports System.Drawing.Drawing2D

'-------------------------------------------------------------------------------
' Création le 16/01/2026  
'        L'axe VERTICAL   est l'axe des y
'        Le point O est situé en Top-Left
' Nommage:
' Paint
'     N° de Couche
'         Grid ou Cellule
'            Texte plus explicite
' 20/09/2022 L'ensemble des dessins sont faits à l'intérieur de Sqr_Cel
'            avec Top-Left + 1 et Width-Height - 3
' Objectif: Reprendre l'ensemble des procédures de dessin avec un g graphics passé en paramètre
'-------------------------------------------------------------------------------   

Module M03_Paint_g

  Public Sub G1_Grid_Paint_g(g As Graphics)
    Jrn_Add_Yellow(Procédure_Name_Get())
    ' Cette fonction construit un seul et grand carré (taille de la grille) pour effacer l'ensemble de la grille
    ' et dessiner ensuite le quadrillage
    ' le fond du formulaire est Color_Frm_BackColor, il est défini dans Sub Présentation_SDK
    g.Clear(Color_Frm_BackColor)
    ' Efface l'intégralité de l'emplacement de la grille avec un carré unique
    Using brsh As New SolidBrush(Color_Frm_BackColor)
      g.FillRectangle(brsh, Gz_Pt_TopLeft.X, Gz_Pt_TopLeft.Y, Bld_WH_Grid, Bld_WH_Grid)
    End Using
    Select Case Plcy_Format_DAB
      Case 0 : G1_Grid_Paint_Quadrillage_00(g)
      Case 1, 2 : G1_Grid_Paint_Quadrillage_04(g)
      Case 3, 4 : G1_Grid_Paint_Quadrillage_20(g)
      Case 5, 6 : G1_Grid_Paint_Quadrillage_36(g)
    End Select
  End Sub

End Module
