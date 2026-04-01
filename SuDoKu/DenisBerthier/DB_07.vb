Module DB_07
  '--------------------------------------------------------------------------------
  ' AIC
  '   ApplyAIC
  '   ExploreAIC
  '   EliminateAIC
  '--------------------------------------------------------------------------------

  Public Function ApplyAIC(ByVal AllCandidates() As Candidate,
                         ByVal Incompatibles() As List(Of Integer),
                         ByVal maxDepth As Integer) As Boolean
    Jrn_Add(, {"AIC     " & Proc_Name_Get()})
    Dim startID As Integer

    For startID = 0 To 728

      If Not AllCandidates(startID).IsActive Then Continue For

      If ExploreAIC(AllCandidates, Incompatibles, startID, maxDepth) Then
        Return True
      End If

    Next

    Return False
  End Function

  Private Function ExploreAIC(ByVal AllCandidates() As Candidate,
                            ByVal Incompatibles() As List(Of Integer),
                            ByVal startID As Integer,
                            ByVal maxDepth As Integer) As Boolean
    'Jrn_Add(, {"        " & Proc_Name_Get()})
    Dim stack As New Stack(Of ChainNode)()
    Dim state(728) As Nullable(Of Boolean)

    Dim cid As Integer
    Dim other As Integer
    Dim wid As Integer
    Dim sid As Integer

    ' Hypothèse de départ : startID est TRUE
    state(startID) = True

    ' Départ : depuis TRUE, on ne propage que via liens faibles → FALSE
    For Each wid In AllCandidates(startID).WeakLinks
      stack.Push(New ChainNode(wid, startID, 1, False, False))
    Next

    While stack.Count > 0

      Dim node As ChainNode = stack.Pop()

      If node.Depth > maxDepth Then Continue While

      cid = node.CandidateID

      ' Déjà visité ?
      If state(cid).HasValue Then
        ' Si même état, rien à faire
        If state(cid).Value = node.IsTrue Then
          Continue While
        Else
          ' Deux états différents atteints par des chemins différents :
          ' on ignore ce chemin, mais on ne déclare PAS de contradiction
          Continue While
        End If
      End If

      ' Affecter l'état logique
      state(cid) = node.IsTrue

      ' Si ce candidat est TRUE, vérifier les incompatibilités
      If node.IsTrue Then
        For Each other In Incompatibles(cid)
          If state(other).HasValue AndAlso state(other).Value = True Then
            ' Deux candidats incompatibles forcés TRUE → contradiction valide
            Return EliminateAIC(AllCandidates, startID)
          End If
        Next
      End If

      ' Propagation booléenne façon Berthier
      If node.IsTrue Then
        ' TRUE → via liens faibles → FALSE
        For Each wid In AllCandidates(cid).WeakLinks
          If Not state(wid).HasValue Then
            stack.Push(New ChainNode(wid, cid, node.Depth + 1, False, False))
          End If
        Next
      Else
        ' FALSE → via liens forts → TRUE
        For Each sid In AllCandidates(cid).StrongLinks
          If Not state(sid).HasValue Then
            stack.Push(New ChainNode(sid, cid, node.Depth + 1, True, True))
          End If
        Next
      End If

    End While

    Return False
  End Function

  Private Function EliminateAIC(ByVal AllCandidates() As Candidate,
                              ByVal startID As Integer) As Boolean
    AllCandidates(startID).IsActive = False
    Trace("AIC", AllCandidates(startID))
    Controle_E(AllCandidates(startID))
    Return True
  End Function

End Module