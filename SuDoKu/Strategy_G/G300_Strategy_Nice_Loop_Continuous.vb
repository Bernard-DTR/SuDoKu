'-------------------------------------------------------------------------------------------
'Mercredi 22/07/2026
' Stratégie Nice_Loop Continupus & Discontinuous
'  
' Préfixe CNL
'         CNL
'5..9.........2.4.98..7...........6...2.......4.6315....1.67..42.....3.7..4.....86
'   La stratégie CNL, DNL et AIC sont en famille 7 pour ne pas générer d'erreur dans la gestion du menu
'                                elles doivent passer en famille 4 !
'   La construction des liens forts est correctes,
'      j'ai mes liens en double, la composition peut être 024 ! 
'      les informations d'unité sont à revoir
'      les codes S et W également
'548.......3....8..1.......3.72..86........3......47..9.1.3...2.2569.........6.7..
'CNL : Exclusion(s) détectée(s) pour le candidat 3 : 3
'6..91.52......7.....................5....4.16..86..4.5.4....1..93.86......5.4138.
'CNL: Exclusion(s) détectée(s) pour le candidat 6 : 3
'CNL: Exclusion(s) détectée(s) pour le candidat 6 : 4

'-------------------------------------------------------------------------------------------
Module G300_Strategy_CNL


  ' ==========================================================================================
  ' MODULE : STRATEGIE CNL (Continuous Nice Loop)
  ' ==========================================================================================

  ' ---------------------------------------------------------
  ' Classe représentant un lien (Strong ou Weak)
  ' ---------------------------------------------------------
  Public Class GcnlLink_Cls
    Public Cel() As Integer          ' {cell1, cell2}
    Public Cdd As String             ' candidat
    Public Type As String            ' "Strong" ou "Weak"
    Public Unité As String           ' unité (L,C,B)
  End Class

  ' ---------------------------------------------------------
  ' Classe représentant une exclusion
  ' ---------------------------------------------------------
  'Public Class GCel_Excl_Cls
  '  Public Cel As Integer
  '  Public Cdd As String
  '  Public Exc() As Integer
  'End Class

  ' ---------------------------------------------------------
  ' Résultat global de la stratégie
  ' ---------------------------------------------------------
  Public Class GcnlRslt_Cls
    Public Candidat As String
    Public Productivité As Boolean
    Public CelExcl As New List(Of GCel_Excl_Cls)
    Public CelExcl_hs As New HashSet(Of Tuple(Of Integer, String))
    Public RoadRight As New List(Of GcnlLink_Cls)
  End Class

  Public GcnlRslt As New GcnlRslt_Cls()

  Public Sub GcnlRslt_Init()
    GcnlRslt.Candidat = ""
    GcnlRslt.Productivité = False
    GcnlRslt.CelExcl.Clear()
    GcnlRslt.CelExcl_hs.Clear()
    GcnlRslt.RoadRight.Clear()
  End Sub

  ' ==========================================================================================
  ' CONSTRUCTION DES LIENS FORTS
  ' ==========================================================================================
  Public Function GLinks_Build_Strong(U_temp(,) As String, candidat As String) As List(Of GcnlLink_Cls)

    Dim L As New List(Of GcnlLink_Cls)
    Dim i As Integer, j As Integer

    ' --- Lignes ---
    For i = 0 To 8
      Dim cells As New List(Of Integer)
      For j = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count = 2 Then
        L.Add(New GcnlLink_Cls With {.Cel = {cells(0), cells(1)}, .Cdd = candidat, .Type = "Strong", .Unité = "L" & i})
      End If
    Next

    ' --- Colonnes ---
    For j = 0 To 8
      Dim cells As New List(Of Integer)
      For i = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count = 2 Then
        L.Add(New GcnlLink_Cls With {.Cel = {cells(0), cells(1)}, .Cdd = candidat, .Type = "Strong", .Unité = "C" & j})
      End If
    Next

    ' --- Boîtes ---
    For b As Integer = 0 To 8
      Dim cells As New List(Of Integer)
      Dim r0 As Integer = (b \ 3) * 3
      Dim c0 As Integer = (b Mod 3) * 3
      For i = 0 To 2
        For j = 0 To 2
          Dim idx As Integer = (r0 + i) * 9 + (c0 + j)
          If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
        Next
      Next
      If cells.Count = 2 Then
        L.Add(New GcnlLink_Cls With {.Cel = {cells(0), cells(1)}, .Cdd = candidat, .Type = "Strong", .Unité = "B" & b})
      End If
    Next

    Return L
  End Function

  ' ==========================================================================================
  ' CONSTRUCTION DES LIENS FAIBLES
  ' ==========================================================================================
  Public Function GLinks_Build_Weak(U_temp(,) As String, candidat As String) As List(Of GcnlLink_Cls)

    Dim L As New List(Of GcnlLink_Cls)
    Dim i As Integer, j As Integer

    ' --- Lignes ---
    For i = 0 To 8
      Dim cells As New List(Of Integer)
      For j = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b As Integer = a + 1 To cells.Count - 1
            L.Add(New GcnlLink_Cls With {.Cel = {cells(a), cells(b)}, .Cdd = candidat, .Type = "Weak", .Unité = "L" & i})
          Next
        Next
      End If
    Next

    ' --- Colonnes ---
    For j = 0 To 8
      Dim cells As New List(Of Integer)
      For i = 0 To 8
        Dim idx As Integer = i * 9 + j
        If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b As Integer = a + 1 To cells.Count - 1
            L.Add(New GcnlLink_Cls With {.Cel = {cells(a), cells(b)}, .Cdd = candidat, .Type = "Weak", .Unité = "C" & j})
          Next
        Next
      End If
    Next

    ' --- Boîtes ---
    For b As Integer = 0 To 8
      Dim cells As New List(Of Integer)
      Dim r0 As Integer = (b \ 3) * 3
      Dim c0 As Integer = (b Mod 3) * 3
      For i = 0 To 2
        For j = 0 To 2
          Dim idx As Integer = (r0 + i) * 9 + (c0 + j)
          If U_temp(idx, 3).Contains(candidat) Then cells.Add(idx)
        Next
      Next
      If cells.Count > 2 Then
        For a As Integer = 0 To cells.Count - 2
          For b2 As Integer = a + 1 To cells.Count - 1
            L.Add(New GcnlLink_Cls With {.Cel = {cells(a), cells(b2)}, .Cdd = candidat, .Type = "Weak", .Unité = "B" & b})
          Next
        Next
      End If
    Next

    Return L
  End Function

  ' ==========================================================================================
  ' CLASSE DFS POUR CNL
  ' ==========================================================================================
  Public Class Edgecnl
    Public Neighbor As Integer
    Public Link As GcnlLink_Cls
  End Class

  Public Class DFS_CNL

    Public Graph As New Dictionary(Of Integer, List(Of Edgecnl))()
    Public AllPaths As New List(Of List(Of GcnlLink_Cls))()

    Public Sub Graph_Build_cnl(links As List(Of GcnlLink_Cls))
      Graph.Clear()

      Dim a As Integer, b As Integer

      For Each ln As GcnlLink_Cls In links
        a = ln.Cel(0)
        b = ln.Cel(1)

        If Not Graph.ContainsKey(a) Then Graph(a) = New List(Of Edgecnl)
        If Not Graph.ContainsKey(b) Then Graph(b) = New List(Of Edgecnl)

        Graph(a).Add(New Edgecnl With {.Neighbor = b, .Link = ln})
        Graph(b).Add(New Edgecnl With {.Neighbor = a, .Link = ln})
      Next
    End Sub
    Public Sub Graph_cnl_Display()
      Jrn_Add(, {Graph.Count & " entrée(s)."})
      Dim l As Integer = 0
      For Each kvp As KeyValuePair(Of Integer, List(Of Edgecnl)) In Graph
        l += 1
        Dim edges As List(Of Edgecnl) = Graph(kvp.Key)
        Dim sb As New System.Text.StringBuilder()
        For Each edge As Edgecnl In edges
          sb.AppendFormat(" → (Cel {0} )", U_Coord(edge.Neighbor))
        Next
        Dim edgeCount As String = $" ({edges.Count})"
        Jrn_Add(, {$"{l,2} Cel {U_Coord(kvp.Key)}:{edgeCount}{sb}"})
      Next
    End Sub
    Public Sub AllPaths_Build_cnl(U_temp(,) As String)
      AllPaths.Clear()

      Dim visited As HashSet(Of GcnlLink_Cls)
      Dim path As List(Of GcnlLink_Cls)

      For Each node As Integer In Graph.Keys
        visited = New HashSet(Of GcnlLink_Cls)
        path = New List(Of GcnlLink_Cls)
        DFS(U_temp, node, Nothing, visited, path)
      Next
    End Sub

    Private Sub DFS(U_temp(,) As String, current As Integer,
                    lastType As String,
                    visited As HashSet(Of GcnlLink_Cls),
                    path As List(Of GcnlLink_Cls))

      Dim e As Edgecnl
      Dim ln As GcnlLink_Cls
      Dim nextNode As Integer

      For Each e In Graph(current)

        ln = e.Link
        nextNode = e.Neighbor

        If visited.Contains(ln) Then Continue For
        If lastType IsNot Nothing AndAlso ln.Type = lastType Then Continue For

        visited.Add(ln)
        path.Add(ln)

        If Path_Is_Productive(U_temp, path, path(0).Cel(0), nextNode) Then
          AllPaths.Add(New List(Of GcnlLink_Cls)(path))
          Exit Sub
        End If

        DFS(U_temp, nextNode, ln.Type, visited, path)

        path.RemoveAt(path.Count - 1)
        visited.Remove(ln)
      Next
    End Sub

  End Class

  ' ==========================================================================================
  ' DETECTION CNL
  ' ==========================================================================================
  Private Function Path_Is_Productive(U_temp(,) As String, path As List(Of GcnlLink_Cls),
                                   startCel As Integer,
                                   currentCel As Integer) As Boolean

    Dim isLoop As Boolean = (path.Count >= 3 AndAlso currentCel = startCel)
    If Not isLoop Then Return False

    Dim evenLength As Boolean = ((path.Count Mod 2) = 0)
    If Not evenLength Then Return False

    Dim candidat As String = GcnlRslt.Candidat
    Dim weakLinks As New List(Of GcnlLink_Cls)
    Dim ln As GcnlLink_Cls

    For Each ln In path
      If ln.Type = "Weak" Then weakLinks.Add(ln)
    Next

    If weakLinks.Count = 0 Then Return False

    Dim i As Integer
    Dim celA As Integer, celB As Integer

    For Each ln In weakLinks
      celA = ln.Cel(0)
      celB = ln.Cel(1)

      For i = 0 To 80
        If i = celA OrElse i = celB Then Continue For
        If Not U_temp(i, 3).Contains(candidat) Then Continue For

        If Is_Vu(i, celA) AndAlso Is_Vu(i, celB) Then
          Dim key As Tuple(Of Integer, String) = Tuple.Create(i, candidat)
          If Not GcnlRslt.CelExcl_hs.Contains(key) Then
            GcnlRslt.CelExcl.Add(New GCel_Excl_Cls With {.Cel = i, .Cdd = candidat, .Exc = {celA, celB}})
            GcnlRslt.CelExcl_hs.Add(key)
          End If
        End If
      Next
    Next

    If GcnlRslt.CelExcl.Count > 0 Then
      GcnlRslt.Productivité = True
      Jrn_Add(, {"CNL : Exclusion(s) détectée(s) pour le candidat " & candidat & " : " & GcnlRslt.CelExcl.Count.ToString()})
      Return True
    End If

    Return False
  End Function

  ' ==========================================================================================
  ' STRATEGIE CNL
  ' ==========================================================================================
  Public Sub Strategy_CNL(U_temp(,) As String)

    GcnlRslt_Init()

    Dim Cdd As Integer
    Dim Candidat As String

    For Cdd = 1 To 9
      Jrn_Add(, {"CNL : Analyse du candidat " & Cdd.ToString()})
      If GcnlRslt.Productivité Then Exit For

      Candidat = Cdd.ToString()
      GcnlRslt_Init()
      GcnlRslt.Candidat = Candidat

      Dim Lstrong As List(Of GcnlLink_Cls) = GLinks_Build_Strong(U_temp, Candidat)
      'Jrn_Add(, {"Nombre de liens forts pour le candidat " & Candidat & " : " & Lstrong.Count.ToString()})
      'For Each ln As GcnlLink_Cls In Lstrong
      'Jrn_Add(, {"Lien : " & ln.Type & " - Candidat : " & ln.Cdd & " - Cellules : " & ln.Cel(0) & ", " & ln.Cel(1) & " - Unité : " & ln.Unité})
      'N'ext ln

      Dim Lweak As List(Of GcnlLink_Cls) = GLinks_Build_Weak(U_temp, Candidat)
      'Jrn_Add(, {"Nombre de liens faibles pour le candidat " & Candidat & " : " & Lweak.Count.ToString()})

      Dim Lall As New List(Of GcnlLink_Cls)
      Lall.AddRange(Lstrong)
      Lall.AddRange(Lweak)

      Dim Solver As New DFS_CNL()
      Solver.Graph_Build_cnl(Lall)
      'Solver.Graph_cnl_Display()
      Solver.AllPaths_Build_cnl(U_temp)

      If GcnlRslt.Productivité Then
        For Each excl As GCel_Excl_Cls In GcnlRslt.CelExcl
          Jrn_Add(, {"CNL : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & excl.Cel & " - Unités : " & String.Join(", ", excl.Exc.Select(Function(c) U_Coord(c)))})
        Next excl

        For Each hs As Tuple(Of Integer, String) In GcnlRslt.CelExcl_hs
          Jrn_Add_Yellow("CNL : Exclusion détectée pour le candidat " & Candidat & " : Cellule " & hs.Item1 & " - Unité(s) : " & hs.Item2)
        Next hs
        Exit For

      End If
    Next

  End Sub


End Module