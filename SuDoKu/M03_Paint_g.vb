Imports System.Drawing.Drawing2D
' Création le 16/01/2026  
Module M03_Paint_g
#Region "G1 Couche Quadrillage"
  Public Sub G1_Grid_Paint_g(g As Graphics)
    ' Cette fonction construit un seul et grand carré (taille de la grille) pour effacer l'ensemble de la grille
    ' et dessiner ensuite le quadrillage
    ' le fond du formulaire est Color_Frm_BackColor, il est défini dans Sub Présentation_SDK
    g.Clear(Color_Frm_BackColor)
    ' Efface l'intégralité de l'emplacement de la grille avec un carré unique
    Using brsh As New SolidBrush(Color_Frm_BackColor)
      g.FillRectangle(brsh, Gz_Pt_TopLeft.X, Gz_Pt_TopLeft.Y, Bld_WH_Grid, Bld_WH_Grid)
    End Using
    Select Case Plcy_Format_DAB
      Case 0, 0 : G1_Grid_Paint_Quadrillage_00_g(g)
      Case 1, 2 : G1_Grid_Paint_Quadrillage_04_g(g)
      Case 3, 4 : G1_Grid_Paint_Quadrillage_20_g(g)
      Case 5, 6 : G1_Grid_Paint_Quadrillage_36_g(g)
    End Select
  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_00_g(g As Graphics)
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
  Public Sub G1_Grid_Paint_Quadrillage_04_g(g As Graphics)
    ' 4 Coins arrondis ou biseautés
    'Quadrillage de référence pour l'entourage, les régions et les cellules horizontal et vertical
    'Sauf les traits 0 et 9
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
          Case 3, 6              '  Trait de 3
            g.DrawLine(pen_t3, Pt_H1, Pt_H2)
            g.DrawLine(pen_t3, Pt_V1, Pt_V2)
          Case 1, 2, 4, 5, 7, 8  '  Trait de 1
            g.DrawLine(pen_t1, Pt_H1, Pt_H2)
            g.DrawLine(pen_t1, Pt_V1, Pt_V2)
        End Select
      Next i

      ' 2 Entourage avec coins arrondis ou biseautés
      Dim Pt1 As Point, Pt2 As Point, Pt3 As Point, Pt4 As Point
      Pt1 = New Point(x:=Gz_Pt_TopLeft.X,
                    y:=Gz_Pt_TopLeft.Y + Gz_Trait_Pos_xy(0))
      Pt2 = New Point(x:=Pt1.X + Bld_WH_Grid - 1,
                    y:=Pt1.Y)
      Pt4 = New Point(x:=Gz_Pt_TopLeft.X,
                    y:=Gz_Pt_TopLeft.Y + Gz_Trait_Pos_xy(9))
      Pt3 = New Point(x:=Pt4.X + Bld_WH_Grid - 1,
                    y:=Pt4.Y)

      Using Grid_Pth As New GraphicsPath
        Select Case Plcy_Format_DAB
          Case 1, 3, 5
            '   Coins arrondis
            With Grid_Pth
              .StartFigure()
              .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
              .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                    'Ligne horizontale A  B en haut
              .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
              .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y - WHhalf)                                    'Ligne verticale   B  C droite
              .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
              .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                    'Ligne horizontale C  D en bas
              .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
              .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y + WHhalf)                                    'Ligne verticale   D  A gauche
              .CloseFigure() 'La figure est déjà fermée
            End With
          Case 2, 4, 6
            '       Coins biseautés
            With Grid_Pth
              .StartFigure()
              .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
              .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
              .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
              .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y - WHquart)
              .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
              .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
              .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
              .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y + WHquart)
              .CloseFigure() 'La figure est déjà fermée
            End With
        End Select
        ' 2 Entourage avec coins arrondis ou biseautés
        g.DrawPath(pen_t3, Grid_Pth)
      End Using
    End Using

  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_20_g(g As Graphics)
    'Quadrillage Coins Arrondis
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

      ' 2 
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
          Dim Pt1 As Point, Pt2 As Point, Pt3 As Point, Pt4 As Point

          Pt1 = New Point(x:=Gz_Pt_TopLeft.X + Top,
                        y:=Gz_Pt_TopLeft.Y + Left)
          Pt2 = New Point(x:=Pt1.X + WH_Région,
                        y:=Pt1.Y)
          Pt3 = New Point(x:=Pt1.X + WH_Région,
                        y:=Pt1.Y + WH_Région)
          Pt4 = New Point(x:=Pt1.X,
                        y:=Pt1.Y + WH_Région)

          ' 1 Calcul avec coins arrondis
          With Reg_Pth
            .StartFigure()
            Select Case Plcy_Format_DAB
              Case 1, 3, 5
                Select Case Région
                  Case 0
                    .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
                    .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                    'Ligne horizontale A  B en haut
                    .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                    .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y)                                             'Ligne verticale   B  C droite
                    .AddLine(Pt3.X, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                             'Ligne horizontale C  D en bas
                    .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                    .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y + WHhalf)                                    'Ligne verticale   D  A gauche
                  Case 1
                    .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
                    .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                    'Ligne horizontale A  B en haut
                    .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                    .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y)                                             'Ligne verticale   B  C droite
                    .AddLine(Pt3.X, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                             'Ligne horizontale C  D en bas
                    .AddLine(Pt4.X, Pt4.Y, Pt1.X, Pt1.Y + WHhalf)                                             'Ligne verticale   D  A gauche
                  Case 2
                    .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
                    .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                    'Ligne horizontale A  B en haut
                    .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                    .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y - WHhalf)                                    'Ligne verticale   B  C droite
                    .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                    .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X, Pt4.Y)                                             'Ligne horizontale C  D en bas
                    .AddLine(Pt4.X, Pt4.Y, Pt1.X, Pt1.Y + WHhalf)                                             'Ligne verticale   D  A gauche
                  Case 3
                    .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
                    .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X, Pt2.Y)                                             'Ligne horizontale A  B en haut
                    .AddLine(Pt2.X, Pt2.Y, Pt3.X, Pt3.Y)                                                      'Ligne verticale   B  C droite
                    .AddLine(Pt3.X, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                             'Ligne horizontale C  D en bas
                    .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                    .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y + WHhalf)                                    'Ligne verticale   D  A gauche
                  Case 4
                    .AddLines({Pt1, Pt2, Pt3, Pt4})
                  Case 5
                    .AddLine(Pt1.X, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                             'Ligne horizontale A  B en haut
                    .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                    .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y - WHhalf)                                    'Ligne verticale   B  C droite
                    .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                    .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X, Pt4.Y)                                             'Ligne horizontale C  D en bas
                    .AddLine(Pt4.X, Pt4.Y, Pt1.X, Pt1.Y)                                                      'Ligne verticale   D  A gauche
                  Case 6
                    .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
                    .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X, Pt2.Y)                                             'Ligne horizontale A  B en haut
                    .AddLine(Pt2.X, Pt2.Y, Pt3.X, Pt3.Y - WHhalf)                                             'Ligne verticale   B  C droite
                    .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                    .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                    'Ligne horizontale C  D en bas
                    .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                    .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y + WHhalf)                                    'Ligne verticale   D  A gauche
                  Case 7
                    .AddLine(Pt1.X, Pt1.Y, Pt2.X, Pt2.Y)                                                      'Ligne horizontale A  B en haut
                    .AddLine(Pt2.X, Pt2.Y, Pt3.X, Pt3.Y - WHhalf)                                             'Ligne verticale   B  C droite
                    .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                    .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                    'Ligne horizontale C  D en bas
                    .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                    .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y)                                             'Ligne verticale   D  A gauche
                  Case 8
                    .AddLine(Pt1.X, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                             'Ligne horizontale A  B en haut
                    .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                    .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y - WHhalf)                                    'Ligne verticale   B  C droite
                    .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                    .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                    'Ligne horizontale C  D en bas
                    .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                    .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y)                                             'Ligne verticale   D  A gauche
                End Select

              Case 2, 4, 6
                Select Case Région
                  Case 0
                    .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
                    .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
                    .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
                    .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y)
                    .AddLine(Pt3.X, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                    .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
                    .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y + WHquart)
                  Case 1
                    .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
                    .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
                    .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
                    .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y)
                    .AddLine(Pt3.X, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                    .AddLine(Pt4.X, Pt4.Y, Pt1.X, Pt1.Y + WHquart)
                  Case 2
                    .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
                    .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
                    .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
                    .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y - WHquart)
                    .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
                    .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X, Pt4.Y)
                    .AddLine(Pt4.X, Pt4.Y, Pt1.X, Pt1.Y + WHquart)
                  Case 3
                    .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
                    .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X, Pt2.Y)
                    .AddLine(Pt2.X, Pt2.Y, Pt3.X, Pt3.Y)
                    .AddLine(Pt3.X, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                    .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
                    .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y + WHquart)
                  Case 4
                    .AddLines({Pt1, Pt2, Pt3, Pt4})
                  Case 5
                    .AddLine(Pt1.X, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
                    .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
                    .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y - WHquart)
                    .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
                    .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X, Pt4.Y)
                    .AddLine(Pt4.X, Pt4.Y, Pt1.X, Pt1.Y)
                  Case 6
                    .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
                    .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X, Pt2.Y)
                    .AddLine(Pt2.X, Pt2.Y, Pt3.X, Pt3.Y - WHquart)
                    .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
                    .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                    .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
                    .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y + WHquart)
                  Case 7
                    .AddLine(Pt1.X, Pt1.Y, Pt2.X, Pt2.Y)
                    .AddLine(Pt2.X, Pt2.Y, Pt3.X, Pt3.Y - WHquart)
                    .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
                    .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                    .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
                    .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y)
                  Case 8
                    .AddLine(Pt1.X, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
                    .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
                    .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y - WHquart)
                    .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
                    .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                    .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
                    .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y)
                End Select
            End Select
            .CloseFigure()
          End With
          ' 2 Entourage avec coins arrondis
          g.DrawPath(pen_t3, Reg_Pth)
        Next Région
      End Using
    End Using

  End Sub
  Public Sub G1_Grid_Paint_Quadrillage_36_g(g As Graphics)
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
            Select Case Plcy_Format_DAB
              Case 1, 3, 5
                .AddArc(New Rectangle(CInt(Pt1.X), CInt(Pt1.Y), WHhalf, WHhalf), 180, 90)                 'Coin en haut à gauche  
                .AddLine(Pt1.X + WHhalf, Pt1.Y, Pt2.X - WHhalf, Pt2.Y)                                    'Ligne horizontale A  B en haut
                .AddArc(New Rectangle(CInt(Pt2.X) - WHhalf, CInt(Pt2.Y), WHhalf, WHhalf), 270, 90)        'Coin en haut à droite  
                .AddLine(Pt2.X, Pt2.Y + WHhalf, Pt3.X, Pt3.Y - WHhalf)                                    'Ligne verticale   B  C droite
                .AddArc(New Rectangle(CInt(Pt3.X) - WHhalf, CInt(Pt3.Y) - WHhalf, WHhalf, WHhalf), 0, 90) 'Coin en bas à droite  
                .AddLine(Pt3.X - WHhalf, Pt3.Y, Pt4.X + WHhalf, Pt4.Y)                                    'Ligne horizontale C  D en bas
                .AddArc(New Rectangle(CInt(Pt4.X), CInt(Pt4.Y) - WHhalf, WHhalf, WHhalf), 90, 90)         'Coin en bas à gauche  
                .AddLine(Pt4.X, Pt4.Y - WHhalf, Pt1.X, Pt1.Y + WHhalf)                                    'Ligne verticale   D  A gauche
              Case 2, 4, 6
                .AddLine(Pt1.X, Pt1.Y + WHquart, Pt1.X + WHquart, Pt1.Y)
                .AddLine(Pt1.X + WHquart, Pt1.Y, Pt2.X - WHquart, Pt2.Y)
                .AddLine(Pt2.X - WHquart, Pt2.Y, Pt2.X, Pt2.Y + WHquart)
                .AddLine(Pt2.X, Pt2.Y + WHquart, Pt3.X, Pt3.Y - WHquart)
                .AddLine(Pt3.X, Pt3.Y - WHquart, Pt3.X - WHquart, Pt3.Y)
                .AddLine(Pt3.X - WHquart, Pt3.Y, Pt4.X + WHquart, Pt4.Y)
                .AddLine(Pt4.X + WHquart, Pt4.Y, Pt4.X, Pt4.Y - WHquart)
                .AddLine(Pt4.X, Pt4.Y - WHquart, Pt1.X, Pt1.Y + WHquart)
            End Select
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