Module DB_02

  Public Function DetectNakedSingles(ByVal AllCandidates() As Candidate) As Boolean
    Jrn_Add_Orange("Passage : " & CStr(DB_Passage) & " " & Proc_Name_Get())
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
          Trace(lastActive)
        End If

      Next
    Next
    Return found
  End Function
  Public Sub Trace(Cdd As Candidate)
    With Cdd
      Dim S As String
      Dim Cellule As Integer = Wh_Cellule_RowCol(.Row - 1, .Col - 1)
      S = $"{CStr(.ID),3} R{ .Row}_C{ .Col} Digit { .Digit}  Block { .Block} IsActive = { CStr(.IsActive),5} IsSolved = { CStr(.IsSolved),5}  {U_Coord(Cellule)}"
      Jrn_Add_Orange(S)
    End With

  End Sub
  Public Sub DB_Solution(ByVal AllCandidates() As Candidate)
    Dim progress As Boolean
    Do
      DB_Passage += 1
      progress = False

      If DetectNakedSingles(AllCandidates) Then
        PropagateSolvedCandidates(AllCandidates)
        progress = True
      End If

      'If DetectHiddenSingles(AllCandidates) Then
      'PropagateSolvedCandidates(AllCandidates)
      'progress = True
      'End If

    Loop While progress
  End Sub

  Private Function DetectHiddenSinglesInRows(ByVal AllCandidates() As Candidate) As Boolean

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
          End If
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
          End If
        End If

      Next
    Next

    Return found

  End Function

  Private Function DetectHiddenSinglesInBlocks(ByVal AllCandidates() As Candidate) As Boolean

    Dim found As Boolean = False

    For block As Integer = 0 To 8

      Dim startRow As Integer = (block \ 3) * 3 + 1
      Dim startCol As Integer = (block Mod 3) * 3 + 1

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
          End If
        End If

      Next
    Next

    Return found

  End Function

  Public Function DetectHiddenSingles(ByVal AllCandidates() As Candidate) As Boolean

    Dim found As Boolean = False

    If DetectHiddenSinglesInRows(AllCandidates) Then found = True
    If DetectHiddenSinglesInCols(AllCandidates) Then found = True
    If DetectHiddenSinglesInBlocks(AllCandidates) Then found = True

    Return found

  End Function


End Module
