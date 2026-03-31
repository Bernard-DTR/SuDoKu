Module DB_07
  Public Function ApplyAIC_old(ByVal AllCandidates() As Candidate,
                         ByVal Incompatibles() As List(Of Integer),
                         ByVal maxDepth As Integer) As Boolean

    Dim progress As Boolean = False

    ' On teste chaque candidat actif comme point de départ
    For startID As Integer = 0 To 728

      If Not AllCandidates(startID).IsActive Then Continue For

      ' AIC = on suppose startID TRUE
      If ExploreAIC(AllCandidates, Incompatibles, startID, maxDepth) Then
        Return True
      End If

    Next

    Return progress
  End Function

  Private Function ExploreAIC(ByVal AllCandidates() As Candidate,
                            ByVal Incompatibles() As List(Of Integer),
                            ByVal startID As Integer,
                            ByVal maxDepth As Integer) As Boolean

    Dim stack As New Stack(Of ChainNode)()
    Dim state(728) As Nullable(Of Boolean)

    ' Hypothèse : startID est TRUE
    state(startID) = True

    ' Départ : liens faibles → FALSE
    For Each wid As Integer In AllCandidates(startID).WeakLinks
      stack.Push(New ChainNode(wid, startID, 1, False, False))
    Next

    ' Départ : liens forts → FALSE
    For Each sid As Integer In AllCandidates(startID).StrongLinks
      stack.Push(New ChainNode(sid, startID, 1, True, False))
    Next

    While stack.Count > 0

      Dim node As ChainNode = stack.Pop()

      If node.Depth > maxDepth Then Continue While

      Dim cid As Integer = node.CandidateID

      ' Déjà visité ?
      If state(cid).HasValue Then

        ' Contradiction logique : TRUE et FALSE
        If state(cid).Value <> node.IsTrue Then
          Return EliminateAIC(AllCandidates, startID)
        End If

        Continue While
      End If

      ' Affecter l'état logique
      state(cid) = node.IsTrue

      ' Si TRUE → vérifier incompatibilités
      If node.IsTrue Then
        For Each other As Integer In Incompatibles(cid)
          If state(other).HasValue AndAlso state(other).Value = True Then
            Return EliminateAIC(AllCandidates, startID)
          End If
        Next
      End If

      ' Propagation alternée
      If node.IsStrong Then
        ' Strong → Weak
        For Each wid As Integer In AllCandidates(cid).WeakLinks
          If Not state(wid).HasValue Then
            stack.Push(New ChainNode(wid, cid, node.Depth + 1, False, Not node.IsTrue))
          End If
        Next
      Else
        ' Weak → Strong
        For Each sid As Integer In AllCandidates(cid).StrongLinks
          If Not state(sid).HasValue Then
            stack.Push(New ChainNode(sid, cid, node.Depth + 1, True, Not node.IsTrue))
          End If
        Next
      End If

    End While

    Return False
  End Function
  Private Function EliminateAIC(ByVal AllCandidates() As Candidate,
                              ByVal startID As Integer) As Boolean

    Dim c As Candidate = AllCandidates(startID)

    ' On élimine le candidat de départ
    c.IsActive = False

    Jrn_Add(, {"AIC élimine → " & Describe(c)})

    Return True
  End Function


End Module
