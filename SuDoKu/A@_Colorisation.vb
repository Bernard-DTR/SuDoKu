Imports System.Drawing.Drawing2D

' Dimanche 27/04/2025 Colorisation
' Préfixe  
' Le module reprend pour l'instant tout ce qui a trait à la colorisation

Friend Module A__Colorisation

  Public Sub Obj_Colors_Save()
    Dim Data As String = String.Join(";", Color_List.Select(Function(c) $"{c.Symbol},{c.Color.ToArgb()}"))
    My.Settings.Obj_Colors = Data
    My.Settings.Save()
  End Sub

  Public Sub Obj_Colors_Load()
    ' Les couleurs sont chargées à/p du paramètre Obj_Color ou de Color_Originale_List
    If Not String.IsNullOrEmpty(My.Settings.Obj_Colors) Then
      Dim Data As String = My.Settings.Obj_Colors
      Dim items As String() = Data.Split(";"c)
      Color_List.Clear()

      For Each item As String In items
        Dim parts As String() = item.Split(","c)
        If parts.Length = 2 Then
          Dim Symbol As String = parts(0)
          Dim ColorValue As Integer = Integer.Parse(parts(1))
          Dim Color As Color = Color.FromArgb(ColorValue)
          Color_List.Add(New Color_Cls With {.Symbol = Symbol, .Color = Color})
        End If
      Next item
    Else
      Color_List = Color_Originale_List
    End If
  End Sub

  Public Sub Objet_List_Display()
    ' Affichage de La liste des Objets à Colorier
    If Objet_List.Count <> 0 Then
      Jrn_Add(, {"Affichage des Objets : " & Objet_List.Count})
    End If
    For Each Obj As Objet_Cls In Objet_List
      With Obj
        Jrn_Add(, {"Couleur " & .Symbol.PadLeft(1) & " " &
                                .Forme.PadRight(8) & " From " &
                                U_Coord(.Cel_From) & "-" & .Cdd_From & " To " &
                                U_Coord(.Cel_To) & "-" & .Cdd_To &
                                " Pt_From " & .Point_From.ToString() & " Pt_To " & .Point_To.ToString()})
      End With
    Next obj
  End Sub

  Function FormatARGB(a As Byte, r As Byte, g As Byte, b As Byte) As String
    Return a.ToString().PadLeft(3, "0"c) & " " &
           r.ToString().PadLeft(3, "0"c) & " " &
           g.ToString().PadLeft(3, "0"c) & " " &
           b.ToString().PadLeft(3, "0"c)
  End Function

  Public Sub Color_List_Display()
    ' Affichage de La liste de Colorisation
    If Color_List.Count <> 0 Then
      Jrn_Add(, {"Affichage des Couleurs : " & Color_List.Count})
    End If
    For Each Clr As Color_Cls In Color_List
      With Clr
        Jrn_Add(, {"Couleur " & .Symbol & " ; " & FormatARGB(.Color.A, .Color.R, .Color.G, .Color.B)})
      End With
    Next clr
  End Sub

  Public Function Color_BySymbol(Symbol As String) As Color
    ' Méthode pour trouver la couleur à/p de Symbol
    For Each Clr As Color_Cls In Color_List
      If Clr.Symbol = Symbol Then Return Clr.Color
    Next clr
    Return Color.Transparent
  End Function

  Public Sub Couleurs_ResetAndColorizsation(Items() As ToolStripMenuItem, Colors As List(Of Color_Cls))
    For i As Integer = 0 To Items.Length - 1
      Items(i).BackColor = Colors(i).Color
      Items(i).Checked = False
    Next i
  End Sub

  Public Sub Couleurs_Ckeck(Items() As ToolStripMenuItem, ActiveItem As String)
    For Each Item As ToolStripMenuItem In Items
      If Mid(Item.Text, 9, 1) = ActiveItem Then
        Item.Checked = True
        Exit For
      End If
    Next item
  End Sub

  Public Sub Objets_Reset(Items() As ToolStripMenuItem)
    For Each Item As ToolStripMenuItem In Items
      Item.Checked = False
      Item.BackColor = SystemColors.Control
    Next item
  End Sub

  Public Sub Objets_CheckAndColor(Items() As ToolStripMenuItem, ActiveItem As String, Color As Color)
    For Each Item As ToolStripMenuItem In Items
      If Item.Text = ActiveItem Then
        Item.Checked = True
        Item.BackColor = Color
        Exit For
      End If
    Next item
  End Sub

  Public Sub DrawCustomArrow(g As Graphics, startPoint As PointF, endPoint As PointF, color As Color, width As Single)
    ' Fonction auxiliaire pour dessiner une flèche personnalisée
    ' Calculer l'angle de la ligne
    Dim angle As Single = CSng(Math.Atan2(endPoint.Y - startPoint.Y, endPoint.X - startPoint.X))
    ' Définir la taille de la flèche
    Dim arrowSize As Integer = 10
    ' Calculer les points de la flèche
    Dim arrowPoint1 As New PointF(
        CInt(endPoint.X - arrowSize * Math.Cos(angle - Math.PI / 6)),
        CInt(endPoint.Y - arrowSize * Math.Sin(angle - Math.PI / 6)))
    Dim arrowPoint2 As New PointF(
        CInt(endPoint.X - arrowSize * Math.Cos(angle + Math.PI / 6)),
        CInt(endPoint.Y - arrowSize * Math.Sin(angle + Math.PI / 6)))
    ' Dessiner la flèche
    Using arrowPen As New Pen(color, width)
      g.DrawLine(arrowPen, endPoint, arrowPoint1)
      g.DrawLine(arrowPen, endPoint, arrowPoint2)
    End Using
  End Sub

  Public Function Get_Distance_Point_Flèche(p As PointF, a As PointF, b As PointF) As Double
    Dim num As Double = Math.Abs((b.Y - a.Y) * p.X - (b.X - a.X) * p.Y + b.X * a.Y - b.Y * a.X)
    Dim den As Double = Math.Sqrt((b.Y - a.Y) ^ 2 + (b.X - a.X) ^ 2)
    Return num / den
  End Function

  Public Function Get_CentreF(Cellule As Integer, Candidat As Integer) As PointF
    Dim Sqr As Rectangle = Sqr_Cdd(Cellule * 10 + Candidat)
    Return New PointF(Sqr.X + (Sqr.Width \ 2), Sqr.Y + (Sqr.Height \ 2))
  End Function

  Public Function Get_Centre(Cellule As Integer, Candidat As Integer) As Point
    Dim Sqr As Rectangle = Sqr_Cdd(Cellule * 10 + Candidat)
    Return New Point(Sqr.X + (Sqr.Width \ 2), Sqr.Y + (Sqr.Height \ 2))
  End Function

  Public Function Get_AdjustedPoints(From_Centre As PointF, To_Centre As PointF) As Points_Struct
    Dim dx As Single = To_Centre.X - From_Centre.X
    Dim dy As Single = To_Centre.Y - From_Centre.Y
    Dim dist As Single = CSng(Math.Sqrt(dx * dx + dy * dy))

    If dist = 0 Then Return New Points_Struct ' Éviter la division par zéro

    Dim t As Single = (WH \ 6) / dist
    Dim Pts As New Points_Struct With {
      .Pt_From = New PointF(From_Centre.X + t * dx, From_Centre.Y + t * dy),
      .Pt_To = New PointF(To_Centre.X - t * dx, To_Centre.Y - t * dy)
    }
    Return Pts
  End Function

  Public Sub G0_Cdd_Flèche_g(g As Graphics, From_Cellule As Integer, From_Candidat As Integer, To_Cellule As Integer, To_Candidat As Integer, Color As Color)
    If From_Cellule = -1 Or From_Candidat = 0 Or To_Cellule = -1 Or To_Candidat = 0 Then Exit Sub

    Dim From_Centre As PointF = Get_CentreF(From_Cellule, From_Candidat)
    Dim To_Centre As PointF = Get_CentreF(To_Cellule, To_Candidat)
    Dim Pts As Points_Struct = Get_AdjustedPoints(From_Centre, To_Centre)

    Using LinePen As New Pen(Color, 2) With {.DashStyle = DashStyle.Solid}
      g.DrawLine(LinePen, Pts.Pt_From, Pts.Pt_To)
    End Using

    DrawCustomArrow(g, Pts.Pt_From, Pts.Pt_To, Color, 2)
    G0_Cdd_Figure(g, From_Cellule, From_Candidat, "Cercle", Color)
    G0_Cdd_Figure(g, To_Cellule, To_Candidat, "Cercle", Color)
  End Sub

  Public Function Get_Pt_From_To_Flèche(From_Cellule As Integer, From_Candidat As Integer, To_Cellule As Integer, To_Candidat As Integer) As Points_Struct
    If From_Cellule = -1 Or To_Cellule = -1 Then Return New Points_Struct ' Retourne une structure vide

    Dim From_Centre As PointF = Get_CentreF(From_Cellule, From_Candidat)
    Dim To_Centre As PointF = Get_CentreF(To_Cellule, To_Candidat)
    Return Get_AdjustedPoints(From_Centre, To_Centre)
  End Function
End Module
