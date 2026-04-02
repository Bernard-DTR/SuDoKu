Module DB_02
  '--------------------------------------------------------------------------------
  ' Exécution des stratégies
  ' Propagation des candidats solveds
  '--------------------------------------------------------------------------------
  Public Function DB_Solution(ByVal AllCandidates() As Candidate) As Boolean

    Dim progress As Boolean
    Dim tour As Integer = 0

    Do
      progress = False
      tour += 1
      Jrn_Add(, {"--- Tour " & tour & " ---"})

      '----------------------------------------------------
      ' 1. Naked Singles
      '----------------------------------------------------
      If DetectNakedSingles(AllCandidates) Then
        progress = True
        Continue Do
      End If

      '----------------------------------------------------
      ' 2. Hidden Singles (lignes, colonnes, blocs)
      '----------------------------------------------------
      If DetectHiddenSingles(AllCandidates) Then
        PropagateSolvedCandidates(AllCandidates)
        progress = True
        Continue Do
      End If

      '----------------------------------------------------
      ' 3. Locked Candidates
      '----------------------------------------------------
      If DetectLockedCandidates(AllCandidates) Then
        progress = True
        Continue Do
      End If

    Loop While progress

    '--------------------------------------------------------
    ' 7. Contrôle final : grille résolue ?
    '--------------------------------------------------------
    Return IsSolved(AllCandidates)

  End Function

  Public Sub PropagateSolvedCandidates(ByVal AllCandidates() As Candidate)

    For Each cdd As Candidate In AllCandidates
      If cdd.IsSolved Then

        Dim r As Integer = cdd.Row
        Dim c As Integer = cdd.Col
        Dim d As Integer = cdd.Digit

        ' 1. CASE : désactiver les autres candidats
        For dd As Integer = 1 To 9
          If dd <> d Then
            Dim id As Integer = Index(r, c, dd)
            If Not AllCandidates(id).IsSolved Then
              AllCandidates(id).IsActive = False
            End If
          End If
        Next

        ' 2. LIGNE
        For cc As Integer = 1 To 9
          If cc <> c Then
            Dim id As Integer = Index(r, cc, d)
            If Not AllCandidates(id).IsSolved Then
              AllCandidates(id).IsActive = False
            End If
          End If
        Next

        ' 3. COLONNE
        For rr As Integer = 1 To 9
          If rr <> r Then
            Dim id As Integer = Index(rr, c, d)
            If Not AllCandidates(id).IsSolved Then
              AllCandidates(id).IsActive = False
            End If
          End If
        Next

        ' 4. BLOC
        Dim br As Integer = ((r - 1) \ 3) * 3 + 1
        Dim bc As Integer = ((c - 1) \ 3) * 3 + 1

        For dr As Integer = 0 To 2
          For dc As Integer = 0 To 2
            Dim rr As Integer = br + dr
            Dim cc As Integer = bc + dc

            If rr <> r OrElse cc <> c Then
              Dim id As Integer = Index(rr, cc, d)
              If Not AllCandidates(id).IsSolved Then
                AllCandidates(id).IsActive = False
              End If
            End If
          Next
        Next
      End If
    Next
  End Sub
End Module
