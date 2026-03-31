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

      '----------------------------------------------------
      ' 4. Construire le graphe (incompatibilités + liens)
      '----------------------------------------------------
      Dim Incompatibles() As List(Of Integer)
      Incompatibles = BuildIncompatibilityGraph(AllCandidates)

      BuildStrongAndWeakLinks(AllCandidates, Incompatibles)

      '----------------------------------------------------
      ' 5a. X-Chains (digit par digit)
      '----------------------------------------------------
      Dim d As Integer
      For d = 1 To 9
        If ApplyXChain(AllCandidates, Incompatibles, d) Then
          progress = True
          Exit For
        End If
      Next
      If progress Then Continue Do

      '----------------------------------------------------
      ' 5b. AIC général façon Berthier
      '----------------------------------------------------
      If ApplyAIC(AllCandidates, Incompatibles, maxDepth:=7) Then
        progress = True
        Continue Do
      End If

      '----------------------------------------------------
      ' 5c. Whips / Braids (si tu veux les réactiver)
      '----------------------------------------------------
      'If ApplyWhips(AllCandidates, Incompatibles, maxDepth:=4) Then
      '    progress = True
      '    Continue Do
      'End If
      '
      'If ApplyBraids(AllCandidates, Incompatibles, maxDepth:=4) Then
      '    progress = True
      '    Continue Do
      'End If

      '----------------------------------------------------
      ' 6. Si aucune stratégie n'a progressé, on sort
      '----------------------------------------------------
    Loop While progress

    '--------------------------------------------------------
    ' 7. Contrôle final : grille résolue ?
    '--------------------------------------------------------
    Return IsSolved(AllCandidates)

  End Function






  Public Function DB_Solution_119(ByVal AllCandidates() As Candidate) As Boolean

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

      '----------------------------------------------------
      ' 4. Construire le graphe (incompatibilités + liens)
      '----------------------------------------------------
      Dim Incompatibles() As List(Of Integer) = BuildIncompatibilityGraph(AllCandidates)
      BuildStrongAndWeakLinks(AllCandidates, Incompatibles)

      '----------------------------------------------------
      ' 5. Chaînes / AIC / X-Chains / Whips / Braids
      '    (à activer selon ce que tu veux tester)
      '----------------------------------------------------

      ' 5a. X-Chains (digit par digit)
      For d As Integer = 1 To 9
        If ApplyXChain(AllCandidates, Incompatibles, d) Then
          progress = True
          Exit For
        End If
      Next
      If progress Then Continue Do

      ' 5b. AIC général (tous digits, toutes unités)
      If ApplyAIC(AllCandidates, Incompatibles, maxDepth:=7) Then
        progress = True
        Continue Do
      End If

      ' 5c. Whips (si tu veux les garder)  
      If ApplyWhipsN(AllCandidates, Incompatibles, 1) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 2) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 3) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 4) Then
        progress = True
        Continue Do
      End If

      ' 5d. Braids (idem)
      If ApplyBraidsN(AllCandidates, Incompatibles, 1) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 2) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 3) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 4) Then
        progress = True
        Continue Do
      End If

    Loop While progress

    '--------------------------------------------------------
    ' Contrôle final : grille résolue ?
    '--------------------------------------------------------
    Return IsSolved(AllCandidates)

  End Function



  Public Sub DB_Solution2(ByVal AllCandidates() As Candidate)

    Dim progress As Boolean
    Dim Incompatibles As List(Of Integer)()

    Do
      progress = False

      ' --- Niveau 1 : stratégies simples ---
      If DetectNakedSingles(AllCandidates) Then
        PropagateSolvedCandidates(AllCandidates)
        progress = True
        Continue Do
      End If

      If DetectHiddenSingles(AllCandidates) Then
        PropagateSolvedCandidates(AllCandidates)
        progress = True
        Continue Do
      End If

      If DetectLockedCandidates(AllCandidates) Then
        progress = True
        Continue Do
      End If

      ' Si la grille est résolue, inutile d'aller plus loin
      If IsSolved(AllCandidates) Then
        Jrn_Add(, {"La grille est résolue."})
        Exit Do
      End If

      ' --- Niveau 2 : whips / braids ---
      Incompatibles = BuildIncompatibilityGraph(AllCandidates)
      BuildStrongAndWeakLinks(AllCandidates, Incompatibles)

      ' Vérifier l'existence de liens forts APRÈS la construction
      If Not HasAnyStrongLinks(AllCandidates) Then
        Jrn_Add(, {"Aucun lien fort."})
        Exit Do
      End If

      For d As Integer = 1 To 9
        If ApplyXChain(AllCandidates, Incompatibles, d) Then
          progress = True
          Continue Do
        End If
      Next

      ' Debug éventuel
      'For id As Integer = 0 To 728
      '  If AllCandidates(id).IsActive AndAlso
      ' Not AllCandidates(id).IsSolved AndAlso
      ' AllCandidates(id).StrongLinks.Count > 0 Then
      '    DebugIncompatibles(AllCandidates, Incompatibles, id)
      '  End If
      'Next

      ' Whip[1]
      If ApplyWhipsN(AllCandidates, Incompatibles, 1) Then
        progress = True
        Continue Do
      End If

      ' Whip[2]
      If ApplyWhipsN(AllCandidates, Incompatibles, 2) Then
        progress = True
        Continue Do
      End If

      ' Whip[3]
      If ApplyWhipsN(AllCandidates, Incompatibles, 3) Then
        progress = True
        Continue Do
      End If

      ' Whip[4]
      If ApplyWhipsN(AllCandidates, Incompatibles, 4) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 5) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 6) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 7) Then
        progress = True
        Continue Do
      End If
      If ApplyWhipsN(AllCandidates, Incompatibles, 8) Then
        progress = True
        Continue Do
      End If


      ' Braid[1]
      If ApplyBraidsN(AllCandidates, Incompatibles, 1) Then
        progress = True
        Continue Do
      End If

      ' Braid[2]
      If ApplyBraidsN(AllCandidates, Incompatibles, 2) Then
        progress = True
        Continue Do
      End If

      ' Braid[3]
      If ApplyBraidsN(AllCandidates, Incompatibles, 3) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 4) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 5) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 6) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 7) Then
        progress = True
        Continue Do
      End If
      If ApplyBraidsN(AllCandidates, Incompatibles, 8) Then
        progress = True
        Continue Do
      End If

      For d As Integer = 1 To 9
        If ApplyXChain(AllCandidates, Incompatibles, d) Then
          progress = True
          Continue Do
        End If
      Next

    Loop While progress

  End Sub
  Public Sub PropagateSolvedCandidates(ByVal AllCandidates() As Candidate)
    'Dim Nb_isActive As Integer
    ' Dès qu'un candidat est IsSolved, IsActive= False des candidats de la même case
    '                                                                de la même unité
    ' Convention 1 : les règles évidentes de propagation des contraintes élémentaires
    ' ECP(cellule), ECP (ligne), ECP(col) et ECP(blk) ne seront jamais affichées.
    For Each cdd As Candidate In AllCandidates
      If cdd.IsSolved Then
        Dim r As Integer = cdd.Row
        Dim c As Integer = cdd.Col
        Dim d As Integer = cdd.Digit
        ' 1. Désactiver les autres candidats de la même case
        For dd As Integer = 1 To 9
          If dd <> d Then
            AllCandidates(Index(r, c, dd)).IsActive = False
          End If
        Next
        ' 2. Désactiver la valeur dans la même ligne
        For cc As Integer = 1 To 9
          If cc <> c Then
            AllCandidates(Index(r, cc, d)).IsActive = False
          End If
        Next
        ' 3. Désactiver la valeur dans la même colonne
        For rr As Integer = 1 To 9
          If rr <> r Then
            AllCandidates(Index(rr, c, d)).IsActive = False
          End If
        Next
        ' 4. Désactiver la valeur dans le même bloc
        Dim br As Integer = ((r - 1) \ 3) * 3 + 1
        Dim bc As Integer = ((c - 1) \ 3) * 3 + 1
        For dr As Integer = 0 To 2
          For dc As Integer = 0 To 2
            Dim rr As Integer = br + dr
            Dim cc As Integer = bc + dc
            If rr <> r OrElse cc <> c Then
              AllCandidates(Index(rr, cc, d)).IsActive = False
            End If
          Next
        Next
      End If
    Next
    'For Each cdd As Candidate In AllCandidates
    '  If cdd.IsActive And Not cdd.IsSolved Then Nb_isActive += 1
    'Next
    'Jrn_Add_Yellow(Proc_Name_Get() & " Nb candidats isActive et Not isSolved : " & Nb_isActive & ".")
  End Sub
  Public Sub DebugIncompatibles(ByVal AllCandidates() As Candidate,
                                ByVal Incompatibles() As List(Of Integer),
                                ByVal id As Integer)

    Dim cand As Candidate = AllCandidates(id)
    Jrn_Add(, {"Candidat ID " & Describe(cand)})
    Jrn_Add(, {"Incompatibles :" & Incompatibles(id).Count})

    For Each otherId As Integer In Incompatibles(id)
      Dim c As Candidate = AllCandidates(otherId)
      Jrn_Add(, {"     -> ID " & Describe(c)})
    Next
  End Sub
End Module
