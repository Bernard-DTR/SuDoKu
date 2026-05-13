Imports System.Drawing.Drawing2D

Module A00_Quadrillage_Test
  Public G_Mrg As Integer = 5

  Public G_tl As New Point(5, 60)

  Public trait_épais As Integer = 3
  Public trait_fin As Integer = 1
  Public pen_épais As Pen = New Pen(Color.Blue, trait_épais)
  Public pen_fin As Pen = New Pen(Color.Red, trait_fin)
  Public hw As Integer = WH

  Public G_trait(9) As Integer        ' centres des 10 traits
  Public hw_t As Integer              ' dimension totale de la grille

  Public CellRect(80) As Rectangle    ' les 81 rectangles des cellules
  Dim R_reg As Integer = 10   ' exemple
  Dim R_cell As Integer = R_reg
  Public RegionPath(8) As GraphicsPath
  Public CellPath(80) As GraphicsPath

  Public Sub Calc_RegionPaths()

    Dim idx As Integer = 0

    For br As Integer = 0 To 2
      For bc As Integer = 0 To 2

        Dim x1 As Integer = G_tl.X + G_trait(bc * 3)
        Dim x2 As Integer = G_tl.X + G_trait(bc * 3 + 3)

        Dim y1 As Integer = G_tl.Y + G_trait(br * 3)
        Dim y2 As Integer = G_tl.Y + G_trait(br * 3 + 3)

        Dim w As Integer = x2 - x1
        Dim h As Integer = y2 - y1

        Dim gp As New GraphicsPath()

        gp.AddArc(x1, y1, R_reg * 2, R_reg * 2, 180, 90)
        gp.AddArc(x2 - 2 * R_reg, y1, R_reg * 2, R_reg * 2, 270, 90)
        gp.AddArc(x2 - 2 * R_reg, y2 - 2 * R_reg, R_reg * 2, R_reg * 2, 0, 90)
        gp.AddArc(x1, y2 - 2 * R_reg, R_reg * 2, R_reg * 2, 90, 90)
        gp.CloseFigure()

        RegionPath(idx) = gp
        idx += 1

      Next
    Next

  End Sub

  Public Sub Calc_CellPaths()

    For r As Integer = 0 To 8
      For c As Integer = 0 To 8

        Dim idx As Integer = r * 9 + c
        Dim rc As Rectangle = CellRect(idx)

        Dim gp As New GraphicsPath()

        Dim TL As Boolean = (r Mod 3 = 0 And c Mod 3 = 0)
        Dim TR As Boolean = (r Mod 3 = 0 And c Mod 3 = 2)
        Dim BL As Boolean = (r Mod 3 = 2 And c Mod 3 = 0)
        Dim BR As Boolean = (r Mod 3 = 2 And c Mod 3 = 2)

        If TL Then
          gp.AddArc(rc.X, rc.Y, 2 * R_cell, 2 * R_cell, 180, 90)
        Else
          gp.AddLine(rc.X, rc.Y, rc.X + R_cell, rc.Y)
        End If

        If TR Then
          gp.AddArc(rc.Right - 2 * R_cell, rc.Y, 2 * R_cell, 2 * R_cell, 270, 90)
        Else
          gp.AddLine(rc.Right - R_cell, rc.Y, rc.Right, rc.Y)
        End If

        If BR Then
          gp.AddArc(rc.Right - 2 * R_cell, rc.Bottom - 2 * R_cell, 2 * R_cell, 2 * R_cell, 0, 90)
        Else
          gp.AddLine(rc.Right, rc.Bottom - R_cell, rc.Right, rc.Bottom)
        End If

        If BL Then
          gp.AddArc(rc.X, rc.Bottom - 2 * R_cell, 2 * R_cell, 2 * R_cell, 90, 90)
        Else
          gp.AddLine(rc.X, rc.Bottom, rc.X, rc.Bottom - R_cell)
        End If

        gp.CloseFigure()
        CellPath(idx) = gp

      Next
    Next

  End Sub

  Public Sub Calc_Traits()
    Dim pos As Integer = trait_épais \ 2
    G_trait(0) = pos

    For i As Integer = 1 To 9
      If i Mod 3 = 0 Then
        pos += hw + trait_épais
      Else
        pos += hw + trait_fin
      End If
      G_trait(i) = pos
    Next

    hw_t = G_trait(9) + (trait_épais \ 2)

  End Sub

  Public Sub Calc_CellRect()

    For r As Integer = 0 To 8
      For c As Integer = 0 To 8

        Dim idx As Integer = r * 9 + c

        Dim leftTraitWidth As Integer = If(c Mod 3 = 0, trait_épais, trait_fin)
        Dim rightTraitWidth As Integer = If((c + 1) Mod 3 = 0, trait_épais, trait_fin)

        Dim topTraitWidth As Integer = If(r Mod 3 = 0, trait_épais, trait_fin)
        Dim bottomTraitWidth As Integer = If((r + 1) Mod 3 = 0, trait_épais, trait_fin)

        Dim x1 As Integer = G_tl.X + G_trait(c) + leftTraitWidth \ 2
        Dim x2 As Integer = G_tl.X + G_trait(c + 1) - rightTraitWidth \ 2

        Dim y1 As Integer = G_tl.Y + G_trait(r) + topTraitWidth \ 2
        Dim y2 As Integer = G_tl.Y + G_trait(r + 1) - bottomTraitWidth \ 2

        CellRect(idx) = New Rectangle(x1, y1, x2 - x1, y2 - y1)

      Next
    Next

  End Sub


  Public Sub quadrillage_test()
    hw = WH

    Calc_Traits()
    Calc_CellRect()
    Using g As Graphics = Frm_Test.CreateGraphics()
      g.SmoothingMode = SmoothingMode.AntiAlias
      ' tracer les traits
      For i As Integer = 0 To 9
        Dim y As Integer = G_tl.Y + G_trait(i)
        Dim x As Integer = G_tl.X + G_trait(i)

        Dim p As Pen = If(i Mod 3 = 0, pen_épais, pen_fin)

        g.DrawLine(p, G_tl.X, y, G_tl.X + hw_t, y)
        g.DrawLine(p, x, G_tl.Y, x, G_tl.Y + hw_t)
      Next
      ' tracer les rectangles des cellules (debug)
      For i As Integer = 0 To 80
        'g.DrawRectangle(Pens.Green, CellRect(i))
        g.FillRectangle(Brushes.Green, CellRect(i))
      Next
    End Using
  End Sub

  Public Sub quadrillage_arrondi_test()
    hw = WH

    Calc_Traits()
    Calc_CellRect()
    Calc_RegionPaths()
    Calc_CellPaths()
    Using g As Graphics = Frm_Test.CreateGraphics()
      g.SmoothingMode = SmoothingMode.AntiAlias
      ' tracer les traits
      For i As Integer = 0 To 8
        g.DrawPath(pen_épais, RegionPath(i))
      Next
      For i As Integer = 1 To 8
        If i = 3 Then Continue For
        If i = 6 Then Continue For
        Dim y As Integer = G_tl.Y + G_trait(i)
        Dim x As Integer = G_tl.X + G_trait(i)

        'Dim p As Pen = If(i Mod 3 = 0, pen_épais, pen_fin)

        g.DrawLine(pen_fin, G_tl.X, y, G_tl.X + hw_t, y)
        g.DrawLine(pen_fin, x, G_tl.Y, x, G_tl.Y + hw_t)
      Next
      For i As Integer = 0 To 80
        g.FillPath(Brushes.LightGreen, CellPath(i))
        'g.DrawPath(pen_fin, CellPath(i))
      Next
    End Using
  End Sub
End Module
