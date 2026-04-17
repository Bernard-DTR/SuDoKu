Imports System.Drawing.Drawing2D
Module M03_Paint_Quadrillage
#Region "G1 Couche Quadrillage"
  Public Sub Build_Quadrillage()
    Bmp_Quadrillage = New Bitmap(Frm_SDK.Width, Frm_SDK.Height)
    Using g As Graphics = Graphics.FromImage(Bmp_Quadrillage)
      G1_Grid_Paint(g)
    End Using
  End Sub
  Public Sub Build_Fond_Valeur()
    Bmp_Fond_Valeur = New Bitmap(Frm_SDK.Width, Frm_SDK.Height)
    Using g As Graphics = Graphics.FromImage(Bmp_Fond_Valeur)
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G2_Cellule_Paint_Fond(g)
        sc.G5_Cellule_Paint_Valeur(g)
      Next
    End Using
  End Sub

  Public Sub Build_Fond_Cellule_Survolee()
    Bmp_Fond_Cellule_Survolee = New Bitmap(WH, WH)
    Dim r As New Rectangle(0, 0, WH, WH)
    Using g As Graphics = Graphics.FromImage(Bmp_Fond_Cellule_Survolee)
      g.SmoothingMode = SmoothingMode.AntiAlias
      'Dim c As Color = Color.FromArgb(110, 80, 160, 255)
      'Dim c As Color = Color.White
      Dim c As Color = Color.FromArgb(110, 247, 238, 239)
      Using font9 As New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Regular),
                              brsh9 As New SolidBrush(c)
        For cdd As Integer = 1 To 9
          Dim row As Integer = (cdd - 1) \ 3
          Dim col As Integer = (cdd - 1) Mod 3
          Dim rc As New Rectangle(r.X + col * WHthird, r.Y + row * WHthird, WHthird, WHthird)
          g.DrawString(Subst_Police(CStr(cdd)), font9, brsh9, rc, Format_Center)
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

  Public Sub G1_Grid_Paint(g As Graphics)
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
      Case 1 : G1_Grid_Paint_Quadrillage_36(g)
    End Select
  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_00(g As Graphics)
    'Les suffixes 00,04,20,36 représentent le nombre de coins arrondis ou biseautés
    'Quadrillage de référence pour l'entourage, les régions et les cellules horizontal et vertical
    Dim Pt_H1 As Point, Pt_H2 As Point, Pt_V1 As Point, Pt_V2 As Point

    Using pen_t3 As New Pen(Color_Trait, Bld_Trait_3),
          pen_t1 As New Pen(Color_Trait, Bld_Trait_1)
      For i As Integer = 0 To 9
        Pt_H1 = New Point(x:=Gz_Pt_TopLeft.X,
                          y:=Gz_Pt_TopLeft.Y + Gz_Trait_Pos_xy(i))
        Pt_H2 = New Point(x:=Pt_H1.X + Bld_WH_Grid - 1,
                          y:=Pt_H1.Y)
        Pt_V1 = New Point(x:=Gz_Pt_TopLeft.X + Gz_Trait_Pos_xy(i),
                          y:=Gz_Pt_TopLeft.Y)
        Pt_V2 = New Point(x:=Pt_V1.X,
                          y:=Pt_V1.Y + Bld_WH_Grid - 1)
        Select Case i
          Case 0, 3, 6, 9        '  Trait de 3
            g.DrawLine(pen_t3, Pt_H1, Pt_H2)
            g.DrawLine(pen_t3, Pt_V1, Pt_V2)
          Case 1, 2, 4, 5, 7, 8  '  Trait de 1
            g.DrawLine(pen_t1, Pt_H1, Pt_H2)
            g.DrawLine(pen_t1, Pt_V1, Pt_V2)
        End Select
      Next i
    End Using
  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_36(g As Graphics)
    ' 1 Traits intérieurs fins 
    Dim Pt_H1 As Point, Pt_H2 As Point, Pt_V1 As Point, Pt_V2 As Point

    Using pen_t3 As New Pen(Color_Trait, Bld_Trait_3),
          pen_t1 As New Pen(Color_Trait, Bld_Trait_1)

      For i As Integer = 0 To 9
        Pt_H1 = New Point(x:=Gz_Pt_TopLeft.X,
                        y:=Gz_Pt_TopLeft.Y + Gz_Trait_Pos_xy(i))
        Pt_H2 = New Point(x:=Pt_H1.X + Bld_WH_Grid - 1,
                        y:=Pt_H1.Y)
        Pt_V1 = New Point(x:=Gz_Pt_TopLeft.X + Gz_Trait_Pos_xy(i),
                        y:=Gz_Pt_TopLeft.Y)
        Pt_V2 = New Point(x:=Pt_V1.X,
                        y:=Pt_V1.Y + Bld_WH_Grid - 1)
        Select Case i
          Case 1, 2, 4, 5, 7, 8  '  Trait de 1
            g.DrawLine(pen_t1, Pt_H1, Pt_H2)
            g.DrawLine(pen_t1, Pt_V1, Pt_V2)
        End Select
      Next i

      'Les 9 Régions sont arrondies
      Dim Pt1 As Point, Pt2 As Point, Pt3 As Point, Pt4 As Point
      Using Reg_Pth As New GraphicsPath

        For Région As Integer = 0 To 8
          Dim Left, Top As Integer
          Select Case Région
            Case 0 : Left = Gz_Trait_Pos_xy(0) : Top = Gz_Trait_Pos_xy(0)
            Case 1 : Left = Gz_Trait_Pos_xy(0) : Top = Gz_Trait_Pos_xy(3)
            Case 2 : Left = Gz_Trait_Pos_xy(0) : Top = Gz_Trait_Pos_xy(6)
            Case 3 : Left = Gz_Trait_Pos_xy(3) : Top = Gz_Trait_Pos_xy(0)
            Case 4 : Left = Gz_Trait_Pos_xy(3) : Top = Gz_Trait_Pos_xy(3)
            Case 5 : Left = Gz_Trait_Pos_xy(3) : Top = Gz_Trait_Pos_xy(6)
            Case 6 : Left = Gz_Trait_Pos_xy(6) : Top = Gz_Trait_Pos_xy(0)
            Case 7 : Left = Gz_Trait_Pos_xy(6) : Top = Gz_Trait_Pos_xy(3)
            Case 8 : Left = Gz_Trait_Pos_xy(6) : Top = Gz_Trait_Pos_xy(6)
          End Select
          Dim WH_Région As Integer = (WH * 3) + 5

          Pt1 = New Point(x:=Gz_Pt_TopLeft.X + Top,
                        y:=Gz_Pt_TopLeft.Y + Left)
          Pt2 = New Point(x:=Pt1.X + WH_Région,
                        y:=Pt1.Y)
          Pt3 = New Point(x:=Pt1.X + WH_Région,
                        y:=Pt1.Y + WH_Région)
          Pt4 = New Point(x:=Pt1.X,
                        y:=Pt1.Y + WH_Région)

          ' 1 Calcul avec les 4 coins arrondis
          With Reg_Pth
            .StartFigure()
            .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
            .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                    'Ligne horizontale A  B en haut
                .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y - WHhalf)                                    'Ligne verticale   B  C droite
                .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                    'Ligne horizontale C  D en bas
                .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y + WHhalf)                                    'Ligne verticale   D  A gauche
            .CloseFigure()
          End With
          ' 2 Entourage avec coins arrondis
          g.DrawPath(pen_t3, Reg_Pth)
        Next Région
      End Using
    End Using
  End Sub
#End Region

End Module