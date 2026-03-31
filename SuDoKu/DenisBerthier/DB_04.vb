Module DB_04
  Public Function BuildIncompatibilityGraph(ByVal AllCandidates() As Candidate) As List(Of Integer)()
    Jrn_Add(, {"      " & Proc_Name_Get()})

    ' Tableau de 729 listes
    Dim Incompatibles(728) As List(Of Integer)

    Dim i As Integer
    For i = 0 To 728
      Incompatibles(i) = New List(Of Integer)()
    Next i

    Dim cand1 As Candidate
    Dim cand2 As Candidate

    ' Comparer chaque paire de candidats
    For i = 0 To 728
      cand1 = AllCandidates(i)

      If cand1.IsActive Then

        Dim j As Integer
        For j = 0 To 728
          If j <> i Then
            cand2 = AllCandidates(j)

            If cand2.IsActive Then

              ' ---------------------------------------------------------
              ' Règle 1 : même case, digit différent
              ' ---------------------------------------------------------
              If cand1.Row = cand2.Row AndAlso cand1.Col = cand2.Col AndAlso
                 cand1.Digit <> cand2.Digit Then

                Incompatibles(i).Add(j)
                Continue For
              End If

              ' ---------------------------------------------------------
              ' Règle 2 : même ligne, même digit
              ' ---------------------------------------------------------
              If cand1.Row = cand2.Row AndAlso cand1.Digit = cand2.Digit AndAlso
                 cand1.Col <> cand2.Col Then

                Incompatibles(i).Add(j)
                Continue For
              End If

              ' ---------------------------------------------------------
              ' Règle 3 : même colonne, même digit
              ' ---------------------------------------------------------
              If cand1.Col = cand2.Col AndAlso cand1.Digit = cand2.Digit AndAlso
                 cand1.Row <> cand2.Row Then

                Incompatibles(i).Add(j)
                Continue For
              End If

              ' ---------------------------------------------------------
              ' Règle 4 : même bloc, même digit
              ' ---------------------------------------------------------
              If cand1.Block = cand2.Block AndAlso cand1.Digit = cand2.Digit AndAlso
                 (cand1.Row <> cand2.Row OrElse cand1.Col <> cand2.Col) Then

                Incompatibles(i).Add(j)
                Continue For
              End If

            End If
          End If
        Next j

      End If
    Next i

    Return Incompatibles

  End Function

  Public Function BuildStrongAndWeakLinks(ByVal AllCandidates() As Candidate,
                                          ByVal Incompatibles() As List(Of Integer)) As Boolean
    Jrn_Add(, {"      " & Proc_Name_Get()})

    Dim progress As Boolean = False

    ' Nettoyer les liens existants
    Dim i As Integer
    For i = 0 To 728
      AllCandidates(i).StrongLinks.Clear()
      AllCandidates(i).WeakLinks.Clear()
    Next i

    ' ---------------------------------------------------------
    ' 1. LIGNES
    ' ---------------------------------------------------------
    Dim r As Integer, d As Integer, c As Integer

    For r = 1 To 9
      For d = 1 To 9

        Dim activeList As New List(Of Integer)()

        For c = 1 To 9
          Dim id As Integer = ((r - 1) * 81) + ((c - 1) * 9) + (d - 1)
          If AllCandidates(id).IsActive Then
            activeList.Add(id)
          End If
        Next

        If activeList.Count = 2 Then
          ' Lien fort
          AddStrongLink(AllCandidates, activeList(0), activeList(1))
          progress = True
        ElseIf activeList.Count > 2 Then
          ' Liens faibles
          AddWeakLinks(AllCandidates, activeList)
          progress = True
        End If

      Next d
    Next r

    ' ---------------------------------------------------------
    ' 2. COLONNES
    ' ---------------------------------------------------------
    'Dim rr As Integer

    For c = 1 To 9
      For d = 1 To 9

        Dim activeList As New List(Of Integer)()

        For rr As Integer = 1 To 9
          Dim id As Integer = ((rr - 1) * 81) + ((c - 1) * 9) + (d - 1)
          If AllCandidates(id).IsActive Then
            activeList.Add(id)
          End If
        Next

        If activeList.Count = 2 Then
          AddStrongLink(AllCandidates, activeList(0), activeList(1))
          progress = True
        ElseIf activeList.Count > 2 Then
          AddWeakLinks(AllCandidates, activeList)
          progress = True
        End If

      Next d
    Next c

    ' ---------------------------------------------------------
    ' 3. BLOCS
    ' ---------------------------------------------------------
    Dim b As Integer, r0 As Integer, c0 As Integer

    For b = 1 To 9
      For d = 1 To 9

        Dim activeList As New List(Of Integer)()

        r0 = ((b - 1) \ 3) * 3 + 1
        c0 = ((b - 1) Mod 3) * 3 + 1

        For rr As Integer = r0 To r0 + 2
          For cc As Integer = c0 To c0 + 2
            Dim id As Integer = ((rr - 1) * 81) + ((cc - 1) * 9) + (d - 1)
            If AllCandidates(id).IsActive Then
              activeList.Add(id)
            End If
          Next
        Next

        If activeList.Count = 2 Then
          AddStrongLink(AllCandidates, activeList(0), activeList(1))
          progress = True
        ElseIf activeList.Count > 2 Then
          AddWeakLinks(AllCandidates, activeList)
          progress = True
        End If

      Next d
    Next b

    ' ---------------------------------------------------------
    ' 4. CASES (toujours 9 candidats → liens faibles)
    ' ---------------------------------------------------------
    For r = 1 To 9
      For c = 1 To 9

        Dim activeList As New List(Of Integer)()

        For d = 1 To 9
          Dim id As Integer = ((r - 1) * 81) + ((c - 1) * 9) + (d - 1)
          If AllCandidates(id).IsActive Then
            activeList.Add(id)
          End If
        Next

        If activeList.Count > 1 Then
          AddWeakLinks(AllCandidates, activeList)
          progress = True
        End If

      Next c
    Next r

    Return progress

  End Function

  Private Sub AddStrongLink(ByVal AllCandidates() As Candidate,
                            ByVal id1 As Integer, ByVal id2 As Integer)
    If Not AllCandidates(id1).StrongLinks.Contains(id2) Then
      AllCandidates(id1).StrongLinks.Add(id2)
    End If
    If Not AllCandidates(id2).StrongLinks.Contains(id1) Then
      AllCandidates(id2).StrongLinks.Add(id1)
    End If
  End Sub


  Private Sub AddWeakLinks(ByVal AllCandidates() As Candidate,
                           ByVal listIds As List(Of Integer))
    Dim i As Integer, j As Integer

    For i = 0 To listIds.Count - 1
      For j = i + 1 To listIds.Count - 1
        Dim id1 As Integer = listIds(i)
        Dim id2 As Integer = listIds(j)
        If Not AllCandidates(id1).WeakLinks.Contains(id2) Then
          AllCandidates(id1).WeakLinks.Add(id2)
        End If
        If Not AllCandidates(id2).WeakLinks.Contains(id1) Then
          AllCandidates(id2).WeakLinks.Add(id1)
        End If
      Next j
    Next i
  End Sub
End Module
