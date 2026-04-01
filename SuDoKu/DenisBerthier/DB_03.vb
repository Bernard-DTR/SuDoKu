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

  Private Function DetectHiddenSinglesInRows(ByVal AllCandidates() As Candidate) As Boolean
    'Jrn_Add(, {"HS    " & Proc_Name_Get()})
    Dim found As Boolean = False
    For row As Integer = 1 To 9
      For val As Integer = 1 To 9
        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0
        For col As Integer = 1 To 9
          Dim c As Candidate = AllCandidates(Index(row, col, val))
          If c.IsActive Then
            activeCount += 1
            lastActive = c
          End If
        Next
        ' Hidden Single : une seule position possible pour cette valeur dans la ligne
        If activeCount = 1 Then
          If Not lastActive.IsSolved Then
            lastActive.IsSolved = True
            found = True
            Trace("HS R", lastActive)
            Controle_P(lastActive)
          End If
        End If
      Next
    Next
    Return found
  End Function
  Private Function DetectHiddenSinglesInCols(ByVal AllCandidates() As Candidate) As Boolean
    'Jrn_Add(, {"HS    " & Proc_Name_Get()})
    Dim found As Boolean = False
    For col As Integer = 1 To 9
      For val As Integer = 1 To 9
        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0
        For row As Integer = 1 To 9
          Dim c As Candidate = AllCandidates(Index(row, col, val))
          If c.IsActive Then
            activeCount += 1
            lastActive = c
          End If
        Next
        If activeCount = 1 Then
          If Not lastActive.IsSolved Then
            lastActive.IsSolved = True
            found = True
            Trace("HS C", lastActive)
            Controle_P(lastActive)
          End If
        End If
      Next
    Next
    Return found
  End Function
  Private Function DetectHiddenSinglesInBlocks(ByVal AllCandidates() As Candidate) As Boolean
    'Jrn_Add(, {"HS    " & Proc_Name_Get()})
    Dim found As Boolean = False
    For block As Integer = 1 To 9
      Dim startRow As Integer = ((block - 1) \ 3) * 3 + 1
      Dim startCol As Integer = ((block - 1) Mod 3) * 3 + 1
      For val As Integer = 1 To 9
        Dim lastActive As Candidate = Nothing
        Dim activeCount As Integer = 0
        For r As Integer = 0 To 2
          For c As Integer = 0 To 2
            Dim row As Integer = startRow + r
            Dim col As Integer = startCol + c
            Dim cand As Candidate = AllCandidates(Index(row, col, val))
            If cand.IsActive Then
              activeCount += 1
              lastActive = cand
            End If
          Next
        Next
        If activeCount = 1 Then
          If Not lastActive.IsSolved Then
            lastActive.IsSolved = True
            found = True
            Trace("HS B", lastActive)
            Controle_P(lastActive)
          End If
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

        ' Trouver les candidats actifs (b,d)
        Dim rows As New HashSet(Of Integer)()
        Dim cols As New HashSet(Of Integer)()

        For Each cand As Candidate In AllCandidates
          If cand.IsActive AndAlso cand.Block = b AndAlso cand.Digit = d Then
            rows.Add(cand.Row)
            cols.Add(cand.Col)
          End If
        Next

        ' POINTING : si tous les candidats sont dans UNE SEULE ligne
        If rows.Count = 1 Then
          Dim targetRow As Integer = rows.First()

          For c = 1 To 9
            ' Exclure les cases du bloc
            Dim blockStartRow As Integer = ((b - 1) \ 3) * 3 + 1
            Dim blockStartCol As Integer = ((b - 1) Mod 3) * 3 + 1

            If Not (targetRow >= blockStartRow AndAlso targetRow <= blockStartRow + 2 AndAlso
                    c >= blockStartCol AndAlso c <= blockStartCol + 2) Then

              Dim id As Integer = ((targetRow - 1) * 81) + ((c - 1) * 9) + (d - 1)

              If AllCandidates(id).IsActive Then
                AllCandidates(id).IsActive = False
                Trace("LC PL", AllCandidates(id))
                Controle_E(AllCandidates(id))
                found = True
              End If
            End If
          Next
        End If

        ' POINTING : si tous les candidats sont dans UNE SEULE colonne
        If cols.Count = 1 Then
          Dim targetCol As Integer = cols.First()

          For r = 1 To 9
            Dim blockStartRow As Integer = ((b - 1) \ 3) * 3 + 1
            Dim blockStartCol As Integer = ((b - 1) Mod 3) * 3 + 1

            If Not (r >= blockStartRow AndAlso r <= blockStartRow + 2 AndAlso
                    targetCol >= blockStartCol AndAlso targetCol <= blockStartCol + 2) Then

              Dim id As Integer = ((r - 1) * 81) + ((targetCol - 1) * 9) + (d - 1)

              If AllCandidates(id).IsActive Then
                AllCandidates(id).IsActive = False
                Trace("LC PC", AllCandidates(id))
                Controle_E(AllCandidates(id))

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
                Dim id As Integer = ((rr - 1) * 81) + ((cc - 1) * 9) + (d - 1)
                If AllCandidates(id).IsActive Then
                  AllCandidates(id).IsActive = False
                  Trace("LC CL", AllCandidates(id))
                  Controle_E(AllCandidates(id))

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
                Dim id As Integer = ((rr - 1) * 81) + ((cc - 1) * 9) + (d - 1)
                If AllCandidates(id).IsActive Then
                  AllCandidates(id).IsActive = False
                  Trace("LC CC", AllCandidates(id))
                  Controle_E(AllCandidates(id))

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