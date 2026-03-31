Module DB_05

  '  Moteur DFS générique, propre, typé, et fidèle à Berthier

  Public Function ExploreChain(ByVal AllCandidates() As Candidate,
                             ByVal Incompatibles() As List(Of Integer),
                             ByVal startID As Integer,
                             ByVal maxDepth As Integer) As Integer

    'Jrn_Add(, {"Exp C ExploreChain"})

    Dim stack As New Stack(Of ChainNode)()
    Dim state(728) As Nullable(Of Boolean)   ' Nothing = pas encore visité, True/Faux = état logique

    ' ---------------------------------------------------------
    ' 1. Hypothèse de départ : startID est VRAI
    ' ---------------------------------------------------------
    state(startID) = True

    ' ---------------------------------------------------------
    ' 2. Les voisins forts du départ deviennent FAUX
    ' ---------------------------------------------------------
    For Each nextID As Integer In AllCandidates(startID).StrongLinks

      ' Si déjà incompatible avec le départ → impossible comme maillon
      If Incompatibles(startID).Contains(nextID) Then Continue For

      stack.Push(New ChainNode(nextID, startID, 1, True, False))
      Jrn_Add(, {"  Start strong → " & Describe(AllCandidates(nextID)) & " (False)"})
    Next

    ' ---------------------------------------------------------
    ' 3. Parcours DFS
    ' ---------------------------------------------------------
    While stack.Count > 0

      Dim node As ChainNode = stack.Pop()

      If node.Depth > maxDepth Then Continue While

      Dim cid As Integer = node.CandidateID

      ' ---------------------------------------------------------
      ' 3a. Vérifier si ce candidat a déjà un état
      ' ---------------------------------------------------------
      If state(cid).HasValue Then

        ' Contradiction : forcé à la fois vrai et faux
        If state(cid).Value <> node.IsTrue Then
          Jrn_Add(, {"    CONTRADICTION : " & Describe(AllCandidates(cid)) &
                          " forcé à la fois VRAI et FAUX"})
          Return cid
        End If

        ' Sinon rien à faire
        Continue While
      End If

      ' ---------------------------------------------------------
      ' 3b. Affecter l'état logique
      ' ---------------------------------------------------------
      state(cid) = node.IsTrue

      Jrn_Add(, {"    Depth " & node.Depth &
                   "  " & If(node.IsStrong, "Strong", "Weak") &
                   "  → " & Describe(AllCandidates(cid)) &
                   "  (" & If(node.IsTrue, "True", "False") & ")"})

      ' ---------------------------------------------------------
      ' 3c. Si ce candidat est VRAI, vérifier incompatibilités
      ' ---------------------------------------------------------
      If node.IsTrue Then
        For Each other As Integer In Incompatibles(cid)
          If state(other).HasValue AndAlso state(other).Value = True Then
            Jrn_Add(, {"    CONTRADICTION : " &
                               Describe(AllCandidates(cid)) &
                               " et " &
                               Describe(AllCandidates(other)) &
                               " tous deux VRAIS"})
            Return cid
          End If
        Next
      End If

      ' ---------------------------------------------------------
      ' 3d. Propagation logique
      ' ---------------------------------------------------------
      If node.IsTrue Then
        ' VRAI → via lien faible → FAUX
        For Each wid As Integer In AllCandidates(cid).WeakLinks
          If Not state(wid).HasValue Then
            stack.Push(New ChainNode(wid, cid, node.Depth + 1, False, False))
          End If
        Next
      Else
        ' FAUX → via lien fort → VRAI
        For Each sid As Integer In AllCandidates(cid).StrongLinks
          If Not state(sid).HasValue Then
            stack.Push(New ChainNode(sid, cid, node.Depth + 1, True, True))
          End If
        Next
      End If

    End While

    Return -1
  End Function


  Public Function ApplyWhipsN(ByVal AllCandidates() As Candidate,
                            ByVal Incompatibles() As List(Of Integer),
                            ByVal n As Integer) As Boolean

    For id As Integer = 0 To 728

      Dim cand As Candidate = AllCandidates(id)

      If cand.IsActive AndAlso Not cand.IsSolved Then
        ' Jrn_Add(, {"Whip[" & n & "] testing start → " & Describe(cand)})

        Dim contradictionID As Integer =
                ExploreChain(AllCandidates, Incompatibles, id, n)

        If contradictionID <> -1 Then
          Jrn_Add(, {"Whip[" & n & "] eliminates → " & Describe(cand)})
          AllCandidates(id).IsActive = False
          Return True
        End If

      End If

    Next

    Return False
  End Function
  Public Function ApplyBraidsN(ByVal AllCandidates() As Candidate,
                             ByVal Incompatibles() As List(Of Integer),
                             ByVal n As Integer) As Boolean

    Dim id As Integer

    For id = 0 To 728

      Dim cand As Candidate = AllCandidates(id)

      ' On ne teste que les candidats actifs non résolus
      If cand.IsActive AndAlso Not cand.IsSolved Then

        'Jrn_Add(, {"Braid[" & n & "] testing start → " & Describe(cand)})

        Dim contradictionID As Integer =
                ExploreBraid(AllCandidates, Incompatibles, id, n)

        If contradictionID <> -1 Then
          Jrn_Add(, {"Braid[" & n & "] eliminates → " & Describe(cand)})
          AllCandidates(id).IsActive = False
          Return True
        End If

      End If

    Next id

    Return False

  End Function


  Public Function ExploreBraid(ByVal AllCandidates() As Candidate,
                             ByVal Incompatibles() As List(Of Integer),
                             ByVal startID As Integer,
                             ByVal maxDepth As Integer) As Integer

    'Jrn_Add(, {"Exp B ExploreBraid"})

    Dim stack As New Stack(Of ChainNode)()
    Dim state(728) As Nullable(Of Boolean)

    Dim nextID As Integer
    Dim cid As Integer
    Dim other As Integer

    ' ---------------------------------------------------------
    ' 1. Hypothèse de départ : startID est VRAI
    ' ---------------------------------------------------------
    state(startID) = True

    ' ---------------------------------------------------------
    ' 2. Départ : tous les liens (forts + faibles)
    ' ---------------------------------------------------------
    For Each nextID In AllCandidates(startID).StrongLinks
      If Not Incompatibles(startID).Contains(nextID) Then
        stack.Push(New ChainNode(nextID, startID, 1, True, False))
        Jrn_Add(, {"  BRAID start strong → " & Describe(AllCandidates(nextID)) & " (False)"})
      End If
    Next

    For Each nextID In AllCandidates(startID).WeakLinks
      If Not Incompatibles(startID).Contains(nextID) Then
        stack.Push(New ChainNode(nextID, startID, 1, False, False))
        Jrn_Add(, {"  BRAID start weak → " & Describe(AllCandidates(nextID)) & " (False)"})
      End If
    Next

    ' ---------------------------------------------------------
    ' 3. Parcours DFS
    ' ---------------------------------------------------------
    While stack.Count > 0

      Dim node As ChainNode = stack.Pop()

      If node.Depth > maxDepth Then Continue While

      cid = node.CandidateID

      ' ---------------------------------------------------------
      ' 3a. Si déjà visité avec un état
      ' ---------------------------------------------------------
      If state(cid).HasValue Then

        ' Contradiction : forcé à la fois vrai et faux
        If state(cid).Value <> node.IsTrue Then
          Jrn_Add(, {"    BRAID CONTRADICTION : " &
                           Describe(AllCandidates(cid)) &
                           " forcé à la fois VRAI et FAUX"})
          Return cid
        End If

        Continue While
      End If

      ' ---------------------------------------------------------
      ' 3b. Affecter l'état logique
      ' ---------------------------------------------------------
      state(cid) = node.IsTrue

      Jrn_Add(, {"    BRAID depth " & node.Depth &
                   "  " & If(node.IsStrong, "Strong", "Weak") &
                   " → " & Describe(AllCandidates(cid)) &
                   " (" & If(node.IsTrue, "True", "False") & ")"})

      ' ---------------------------------------------------------
      ' 3c. Si ce candidat est VRAI, vérifier incompatibilités
      ' ---------------------------------------------------------
      If node.IsTrue Then
        For Each other In Incompatibles(cid)
          If state(other).HasValue AndAlso state(other).Value = True Then
            Jrn_Add(, {"    BRAID CONTRADICTION : " &
                               Describe(AllCandidates(cid)) &
                               " et " &
                               Describe(AllCandidates(other)) &
                               " tous deux VRAIS"})
            Return cid
          End If
        Next
      End If

      ' ---------------------------------------------------------
      ' 3d. Propagation logique (braid = tous les liens)
      ' ---------------------------------------------------------

      ' Si VRAI → via faibles → FAUX
      If node.IsTrue Then
        For Each nextID In AllCandidates(cid).WeakLinks
          If Not state(nextID).HasValue Then
            stack.Push(New ChainNode(nextID, cid, node.Depth + 1, False, False))
          End If
        Next
      End If

      ' Si FAUX → via forts → VRAI
      If Not node.IsTrue Then
        For Each nextID In AllCandidates(cid).StrongLinks
          If Not state(nextID).HasValue Then
            stack.Push(New ChainNode(nextID, cid, node.Depth + 1, True, True))
          End If
        Next
      End If

    End While

    Return -1
  End Function



End Module
