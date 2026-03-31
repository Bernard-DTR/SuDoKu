Module DB_06

  Public Function ApplyXChain(ByVal AllCandidates() As Candidate,
                            ByVal Incompatibles() As List(Of Integer),
                            ByVal digit As Integer) As Boolean

    Dim progress As Boolean = False

    ' Liste des candidats actifs pour ce digit
    Dim activeList As New List(Of Integer)()

    For id As Integer = 0 To 728
      If AllCandidates(id).IsActive AndAlso AllCandidates(id).Digit = digit Then
        activeList.Add(id)
      End If
    Next

    ' Pour chaque candidat actif, on cherche une chaîne de liens forts
    For Each startID As Integer In activeList

      ' DFS pour liens forts uniquement
      Dim stack As New Stack(Of Integer)()
      Dim visited As New HashSet(Of Integer)()

      stack.Push(startID)

      While stack.Count > 0

        Dim cid As Integer = stack.Pop()

        If visited.Contains(cid) Then Continue While
        visited.Add(cid)

        ' Explorer uniquement les liens forts
        For Each nextID As Integer In AllCandidates(cid).StrongLinks

          ' Ne pas revenir au départ
          If nextID = startID Then Continue For

          ' Si nextID est incompatible avec startID → X-Chain trouvée
          If Incompatibles(startID).Contains(nextID) Then

            ' Éliminer le digit dans les cases qui voient les deux extrémités
            progress = progress Or EliminateXChain(AllCandidates, startID, nextID)

            If progress Then Return True
          End If

          ' Continuer la chaîne
          If Not visited.Contains(nextID) Then
            stack.Push(nextID)
          End If

        Next

      End While

    Next

    Return progress
  End Function

  Private Function EliminateXChain(ByVal AllCandidates() As Candidate,
                                 ByVal id1 As Integer,
                                 ByVal id2 As Integer) As Boolean

    Dim progress As Boolean = False

    Dim c1 As Candidate = AllCandidates(id1)
    Dim c2 As Candidate = AllCandidates(id2)

    ' Parcourir tous les candidats actifs du même digit
    For id As Integer = 0 To 728

      Dim c As Candidate = AllCandidates(id)

      If c.IsActive AndAlso c.Digit = c1.Digit Then

        ' Si c voit les deux extrémités → élimination
        If InSameUnit(c, c1) AndAlso InSameUnit(c, c2) Then

          ' Ne pas éliminer les extrémités elles-mêmes
          If id <> id1 AndAlso id <> id2 Then
            c.IsActive = False
            progress = True

            Jrn_Add(, {"X-Chain élimine → " & Describe(c)})
          End If

        End If

      End If

    Next

    Return progress
  End Function

  Private Function InSameUnit(ByVal a As Candidate, ByVal b As Candidate) As Boolean
    Return (a.Row = b.Row) OrElse (a.Col = b.Col) OrElse (a.Block = b.Block)
  End Function



End Module
