Module DB_03
  '--------------------------------------------------------------------------------
  ' Stratégies de base
  '   NakedSingles
  '   HiddenSingles (Row, Col, Block)
  '   LockedCandidates
  '--------------------------------------------------------------------------------

  Public Function DetectNakedSingles(ByVal AllCandidates() As Candidate) As Boolean
    Jrn_Add(, {"NS      " & Proc_Name_Get()})
    Dim found As Boolean = False
    For row As Integer = 1 To 9
      For col As Integer = 1 To 9
        ' Vérifier si la case est déjà résolue
        Dim alreadySolved As Boolean = False
        For d As Integer = 1 To 9
          If AllCandidates(Index(row, col, d)).IsSolved Then
            alreadySolved = True
            Exit For
          End If
        Next
        If alreadySolved Then Continue For
        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0
        For val As Integer = 1 To 9
          Dim c As Candidate = AllCandidates(Index(row, col, val))
          If c.IsActive Then
            activeCount += 1
            lastActive = c
          End If
        Next
        If activeCount = 1 Then
          lastActive.IsSolved = True
          found = True
          Trace("NS", lastActive)
          Controle_P(lastActive)
        End If
      Next
    Next
    Return found
  End Function
  Public Function DetectHiddenSingles(ByVal AllCandidates() As Candidate) As Boolean
    Jrn_Add(, {"HS      " & Proc_Name_Get()})
    Dim found As Boolean = False
    If DetectHiddenSinglesInRows(AllCandidates) Then found = True
    If DetectHiddenSinglesInCols(AllCandidates) Then found = True
    If DetectHiddenSinglesInBlocks(AllCandidates) Then found = True
    Return found
  End Function

  Private Function DetectHiddenSinglesInBlocks(ByVal AllCandidates() As Candidate) As Boolean
    Dim found As Boolean = False

    For block As Integer = 1 To 9

      Dim startRow As Integer = ((block - 1) \ 3) * 3 + 1
      Dim startCol As Integer = ((block - 1) Mod 3) * 3 + 1

      For val As Integer = 1 To 9

        ' 1. Vérifier si le bloc contient déjà val
        Dim valAlreadyInBlock As Boolean = False
        For dr As Integer = 0 To 2
          For dc As Integer = 0 To 2
            If AllCandidates(Index(startRow + dr, startCol + dc, val)).IsSolved Then
              valAlreadyInBlock = True
              Exit For
            End If
          Next
          If valAlreadyInBlock Then Exit For
        Next
        If valAlreadyInBlock Then Continue For

        ' 2. Compter les candidats actifs
        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0

        For dr As Integer = 0 To 2
          For dc As Integer = 0 To 2
            Dim row As Integer = startRow + dr
            Dim col As Integer = startCol + dc
            Dim cand As Candidate = AllCandidates(Index(row, col, val))
            If cand.IsActive Then
              activeCount += 1
              lastActive = cand
            End If
          Next
        Next

        ' 3. Hidden Single valide
        If activeCount = 1 Then

          If lastActive.IsSolved Then Continue For

          ' Vérifier ligne
          For cc As Integer = 1 To 9
            If AllCandidates(Index(lastActive.Row, cc, val)).IsSolved Then
              GoTo SkipHS
            End If
          Next

          ' Vérifier colonne
          For rr As Integer = 1 To 9
            If AllCandidates(Index(rr, lastActive.Col, val)).IsSolved Then
              GoTo SkipHS
            End If
          Next

          lastActive.IsSolved = True
          found = True
          Trace("HS B", lastActive)
          Controle_P(lastActive)

SkipHS:
        End If

      Next
    Next

    Return found
  End Function

  Private Function DetectHiddenSinglesInCols(ByVal AllCandidates() As Candidate) As Boolean
    Dim found As Boolean = False

    For col As Integer = 1 To 9
      For val As Integer = 1 To 9

        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0

        ' 1. Vérifier si la colonne contient déjà val
        Dim valAlreadyInCol As Boolean = False
        For row As Integer = 1 To 9
          If AllCandidates(Index(row, col, val)).IsSolved Then
            valAlreadyInCol = True
            Exit For
          End If
        Next
        If valAlreadyInCol Then Continue For

        ' 2. Compter les candidats actifs
        For row As Integer = 1 To 9
          Dim c As Candidate = AllCandidates(Index(row, col, val))
          If c.IsActive Then
            activeCount += 1
            lastActive = c
          End If
        Next

        ' 3. Hidden Single valide
        If activeCount = 1 Then

          ' Vérifier que la case n'est pas déjà résolue
          If lastActive.IsSolved Then Continue For

          ' Vérifier que la ligne ne contient pas déjà val
          For cc As Integer = 1 To 9
            If AllCandidates(Index(lastActive.Row, cc, val)).IsSolved Then
              GoTo SkipHS
            End If
          Next

          ' Vérifier que le bloc ne contient pas déjà val
          Dim br As Integer = ((lastActive.Row - 1) \ 3) * 3 + 1
          Dim bc As Integer = ((lastActive.Col - 1) \ 3) * 3 + 1
          For dr As Integer = 0 To 2
            For dc As Integer = 0 To 2
              If AllCandidates(Index(br + dr, bc + dc, val)).IsSolved Then
                GoTo SkipHS
              End If
            Next
          Next

          ' OK : Hidden Single réel
          lastActive.IsSolved = True
          found = True
          Trace("HS C", lastActive)
          Controle_P(lastActive)

SkipHS:
        End If

      Next
    Next

    Return found
  End Function

  Private Function DetectHiddenSinglesInRows(ByVal AllCandidates() As Candidate) As Boolean
    Dim found As Boolean = False

    For row As Integer = 1 To 9
      For val As Integer = 1 To 9

        ' 1. Vérifier si la ligne contient déjà val
        Dim valAlreadyInRow As Boolean = False
        For col As Integer = 1 To 9
          If AllCandidates(Index(row, col, val)).IsSolved Then
            valAlreadyInRow = True
            Exit For
          End If
        Next
        If valAlreadyInRow Then Continue For

        ' 2. Compter les candidats actifs
        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0

        For col As Integer = 1 To 9
          Dim c As Candidate = AllCandidates(Index(row, col, val))
          If c.IsActive Then
            activeCount += 1
            lastActive = c
          End If
        Next

        ' 3. Hidden Single valide
        If activeCount = 1 Then

          ' Vérifier que la case n'est pas déjà résolue
          If lastActive.IsSolved Then Continue For

          ' Vérifier colonne
          For rr As Integer = 1 To 9
            If AllCandidates(Index(rr, lastActive.Col, val)).IsSolved Then
              GoTo SkipHS
            End If
          Next

          ' Vérifier bloc
          Dim br As Integer = ((row - 1) \ 3) * 3 + 1
          Dim bc As Integer = ((lastActive.Col - 1) \ 3) * 3 + 1
          For dr As Integer = 0 To 2
            For dc As Integer = 0 To 2
              If AllCandidates(Index(br + dr, bc + dc, val)).IsSolved Then
                GoTo SkipHS
              End If
            Next
          Next

          ' OK : Hidden Single réel
          lastActive.IsSolved = True
          found = True
          Trace("HS R", lastActive)
          Controle_P(lastActive)

SkipHS:
        End If

      Next
    Next

    Return found
  End Function


  Public Function DetectLockedCandidates(ByVal AllCandidates() As Candidate) As Boolean
    Jrn_Add(, {"LC      " & Proc_Name_Get()})
    Dim found As Boolean = False

    ' ---------------------------------------------------------
    ' 1. POINTING : bloc → ligne/colonne
    ' ---------------------------------------------------------
    Dim b As Integer, d As Integer, r As Integer, c As Integer

    For b = 1 To 9
      For d = 1 To 9

        Dim rows As New HashSet(Of Integer)()
        Dim cols As New HashSet(Of Integer)()

        ' Trouver les candidats actifs (b,d)
        For Each cand As Candidate In AllCandidates
          If cand.IsActive AndAlso cand.Block = b AndAlso cand.Digit = d Then
            rows.Add(cand.Row)
            cols.Add(cand.Col)
          End If
        Next

        ' POINTING : bloc → ligne
        If rows.Count = 1 Then
          Dim targetRow As Integer = rows.First()

          Dim blockStartRow As Integer = ((b - 1) \ 3) * 3 + 1
          Dim blockStartCol As Integer = ((b - 1) Mod 3) * 3 + 1

          For c = 1 To 9
            If Not (targetRow >= blockStartRow AndAlso targetRow <= blockStartRow + 2 AndAlso
                            c >= blockStartCol AndAlso c <= blockStartCol + 2) Then

              Dim id As Integer = Index(targetRow, c, d)

              ' PROTECTION : ne jamais supprimer un solved
              If AllCandidates(id).IsSolved Then Continue For

              If AllCandidates(id).IsActive Then
                AllCandidates(id).IsActive = False
                Trace("LC PL", AllCandidates(id))
                found = True
              End If
            End If
          Next
        End If

        ' POINTING : bloc → colonne
        If cols.Count = 1 Then
          Dim targetCol As Integer = cols.First()

          Dim blockStartRow As Integer = ((b - 1) \ 3) * 3 + 1
          Dim blockStartCol As Integer = ((b - 1) Mod 3) * 3 + 1

          For r = 1 To 9
            If Not (r >= blockStartRow AndAlso r <= blockStartRow + 2 AndAlso
                            targetCol >= blockStartCol AndAlso targetCol <= blockStartCol + 2) Then

              Dim id As Integer = Index(r, targetCol, d)

              If AllCandidates(id).IsSolved Then Continue For

              If AllCandidates(id).IsActive Then
                AllCandidates(id).IsActive = False
                Trace("LC PC", AllCandidates(id))
                found = True
              End If
            End If
          Next
        End If

      Next d
    Next b

    ' ---------------------------------------------------------
    ' 2. CLAIMING : ligne/colonne → bloc
    ' ---------------------------------------------------------

    ' CLAIMING LIGNE
    For r = 1 To 9
      For d = 1 To 9

        Dim blocks As New HashSet(Of Integer)()

        For Each cand As Candidate In AllCandidates
          If cand.IsActive AndAlso cand.Row = r AndAlso cand.Digit = d Then
            blocks.Add(cand.Block)
          End If
        Next

        If blocks.Count = 1 Then
          Dim targetBlock As Integer = blocks.First()

          Dim blockStartRow As Integer = ((targetBlock - 1) \ 3) * 3 + 1
          Dim blockStartCol As Integer = ((targetBlock - 1) Mod 3) * 3 + 1

          For rr As Integer = blockStartRow To blockStartRow + 2
            For cc As Integer = blockStartCol To blockStartCol + 2
              If rr <> r Then
                Dim id As Integer = Index(rr, cc, d)

                If AllCandidates(id).IsSolved Then Continue For

                If AllCandidates(id).IsActive Then
                  AllCandidates(id).IsActive = False
                  Trace("LC CL", AllCandidates(id))
                  found = True
                End If
              End If
            Next
          Next
        End If

      Next d
    Next r

    ' CLAIMING COLONNE
    For c = 1 To 9
      For d = 1 To 9

        Dim blocks As New HashSet(Of Integer)()

        For Each cand As Candidate In AllCandidates
          If cand.IsActive AndAlso cand.Col = c AndAlso cand.Digit = d Then
            blocks.Add(cand.Block)
          End If
        Next

        If blocks.Count = 1 Then
          Dim targetBlock As Integer = blocks.First()

          Dim blockStartRow As Integer = ((targetBlock - 1) \ 3) * 3 + 1
          Dim blockStartCol As Integer = ((targetBlock - 1) Mod 3) * 3 + 1

          For rr As Integer = blockStartRow To blockStartRow + 2
            For cc As Integer = blockStartCol To blockStartCol + 2
              If cc <> c Then
                Dim id As Integer = Index(rr, cc, d)

                If AllCandidates(id).IsSolved Then Continue For

                If AllCandidates(id).IsActive Then
                  AllCandidates(id).IsActive = False
                  Trace("LC CC", AllCandidates(id))
                  found = True
                End If
              End If
            Next
          Next
        End If

      Next d
    Next c

    Return found
  End Function


End Module