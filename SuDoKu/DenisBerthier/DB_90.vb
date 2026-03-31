Module DB_90
  Public Sub AllCandidates_Display(AllCandidates() As Candidate)
    Dim Nb_IsActive As Integer
    Dim Nb_IsActiveSolved As Integer
    Dim Nb_IsActiveNotsolved As Integer
    Dim Nb_IsSolved As Integer
    For Each Cdd As Candidate In AllCandidates
      If Cdd.IsActive Then Nb_IsActive += 1
      If Cdd.IsActive And Cdd.IsSolved Then Nb_IsActiveSolved += 1
      If Cdd.IsActive And Not Cdd.IsSolved Then Nb_IsActiveNotsolved += 1
      If Cdd.IsSolved Then Nb_IsSolved += 1
      Trace("Dsp  ", Cdd)
    Next
    Jrn_Add_Yellow("IsActive           " & Nb_IsActive)
    Jrn_Add_Yellow("IsActive_Solved    " & Nb_IsActiveSolved)
    Jrn_Add_Yellow("IsActive_NotSolved " & Nb_IsActiveNotsolved)
    Jrn_Add_Yellow("IsSolved           " & Nb_IsSolved)
  End Sub

  Public Sub AllCandidates_Display_IsSolved(AllCandidates() As Candidate)
    Dim Nb_IsActive As Integer
    Dim Nb_IsSolved As Integer
    For Each Cdd As Candidate In AllCandidates
      If Cdd.IsSolved Then
        If Cdd.IsActive Then Nb_IsActive += 1
        If Cdd.IsSolved Then Nb_IsSolved += 1
        Trace("Dsp", Cdd)
      End If
    Next
    Jrn_Add_Yellow("IsActive " & Nb_IsActive)
    Jrn_Add_Yellow("IsSolved " & Nb_IsSolved)
  End Sub

End Module
