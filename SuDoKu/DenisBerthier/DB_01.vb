Imports SuDoKu.DancingLink

Module DB_01
  '--------------------------------------------------------------------------------
  ' Objectifs:
  '    Transférer les données entre l'application SDK <---> Denis_Berthier
  '--------------------------------------------------------------------------------
  Public Sub AllCandidates_Init(AllCandidates() As Candidate)
    ' Initialisation de AllCandidates(728)
    ' id varie de 0 à 728, (9 * 9 * 9 = 729)
    ' Block = ((r - 1) \ 3) * 3 + ((c - 1) \ 3) + 1
    ' IsActive = True et IsSolved = False
    Dim id As Integer = 0
    For r As Integer = 1 To 9
      For c As Integer = 1 To 9
        For d As Integer = 1 To 9
          AllCandidates(id) = New Candidate(id, r, c, d)
          id += 1
        Next
      Next
    Next
  End Sub
  Public Sub SDK_AllCandidate(grid As String, AllCandidates() As Candidate)
    ' Copie U (SDK) ---> AllCandidates(Denis Berthier) 
    Dim Cellule As Integer = 0
    Dim Nb_ValeursInitiales As Integer
    For r As Integer = 1 To 9
      For c As Integer = 1 To 9
        Dim ch As Char = grid(Cellule)
        Cellule += 1
        If ch >= "1"c AndAlso ch <= "9"c Then
          Dim solvedDigit As Integer = CInt(ch.ToString())
          For d As Integer = 1 To 9
            Dim id As Integer = Index(r, c, d)
            If d = solvedDigit Then
              AllCandidates(id).IsActive = True
              AllCandidates(id).IsSolved = True
              Nb_ValeursInitiales += 1
              Trace("vi", AllCandidates(id))
            Else
              AllCandidates(id).IsActive = False
              AllCandidates(id).IsSolved = False
            End If
          Next
        Else
          ' Case vide : tous les candidats restent actifs
          For d As Integer = 1 To 9
            Dim id As Integer = Index(r, c, d)
            AllCandidates(id).IsActive = True
            AllCandidates(id).IsSolved = False
          Next
        End If
      Next
    Next
    Jrn_Add(, {Proc_Name_Get() & " " & Nb_ValeursInitiales & " valeurs initiales chargées."})
    Jrn_Add(, {"Solution attendue calculée avec Dancing_Link"})
    Jrn_Add(, {XSolution})
  End Sub
  Public Sub AllCandidates_SDK(AllCandidates() As Candidate)
    ' Copie AllCandidates(Denis Berthier) ---> U (SDK)
    ' Tableau temporaire : 81 cellules × 3 champs
    Dim U_temp(80, 3) As String
    ' Initialisation
    For i As Integer = 0 To 80
      U_temp(i, 1) = " "
      U_temp(i, 2) = " "
      U_temp(i, 3) = Cnddts_Blancs
    Next
    ' Remplissage
    For Each Cdd As Candidate In AllCandidates
      If Not Cdd.IsActive Then Continue For
      Dim Cellule As Integer = Wh_Cellule_RowCol(Cdd.Row - 1, Cdd.Col - 1)
      If Cdd.IsSolved Then
        U_temp(Cellule, 2) = Cdd.Digit.ToString()
        U_temp(Cellule, 3) = Cnddts_Blancs
      Else
        ' Candidat non résolu : on ajoute le digit dans la chaîne
        Dim Candidats As String = U_temp(Cellule, 3)
        Candidats = ReplaceDigit(Candidats, Cdd.Digit)
        U_temp(Cellule, 3) = Candidats
      End If
    Next
    For i As Integer = 0 To 80
      'U(i, 1) = U_temp(i, 1) '= " "
      U(i, 2) = U_temp(i, 2) '= " "
      U(i, 3) = U_temp(i, 3) '= Cnddts_Blancs
    Next
    Frm_SDK.Invalidate()
  End Sub
  '--------------------------------------------------------------------------------
  ' Autres Fonctions 
  '--------------------------------------------------------------------------------
  Public Function IsSolved(ByVal AllCandidates() As Candidate) As Boolean
    'Dès qu'une cellule a plus d'un candidat actif, le grille n'est pas résolue
    Dim activeCount As Integer
    For r As Integer = 1 To 9
      For c As Integer = 1 To 9
        activeCount = 0
        For d As Integer = 1 To 9
          Dim id As Integer = Index(r, c, d)
          If AllCandidates(id).IsActive Then
            activeCount += 1
          End If
        Next d
        ' Une case doit avoir exactement 1 candidat actif
        If activeCount <> 1 Then
          Return False
        End If
      Next c
    Next r
    'Chaque cellule a 1 seul candidat actif: la grille est résolue
    Return True
  End Function
  '--------------------------------------------------------------------------------
  ' Fonctions utilitaires
  '--------------------------------------------------------------------------------
  Public Function Index(r As Integer, c As Integer, d As Integer) As Integer
    Return (r - 1) * 81 + (c - 1) * 9 + (d - 1)
  End Function
  Private Function ReplaceDigit(source As String, digit As Integer) As String
    Dim chars() As Char = source.ToCharArray()
    chars(digit - 1) = ChrW(48 + digit)   ' Remplace par "1"…"9"
    Return New String(chars)
  End Function
  Public Sub Trace(DB_Stg As String, Cdd As Candidate)
    Jrn_Add(, {DB_Stg.PadRight(7) & Describe(Cdd)})
  End Sub
  Public Function Describe(ByVal Cdd As Candidate) As String
    With Cdd
      Return $" {CStr(.ID),3} L{ .Row}_C{ .Col} { .Digit}  Act={ CStr(.IsActive),-5} Sol={ CStr(.IsSolved),-5} Strong:{ .StrongLinks.Count} Weak:{ .WeakLinks.Count}"
    End With
  End Function
  Public Function Controle_P(Cdd As Candidate) As Boolean
    Dim Cellule As Integer = Wh_Cellule_RowCol(Cdd.Row - 1, Cdd.Col - 1)
    If XSolution(Cellule) <> CStr(Cdd.Digit) Then
      Jrn_Add(, {"⛔" & "   Erreur en " & U_Coord(Cellule) & " " & XSolution(Cellule) & " est attendu au lieu de " & CStr(Cdd.Digit) & "."})
      Return False
    End If
    Return True
  End Function
  Public Function Controle_E(Cdd As Candidate) As Boolean
    Dim Cellule As Integer = Wh_Cellule_RowCol(Cdd.Row - 1, Cdd.Col - 1)
    If XSolution(Cellule) = CStr(Cdd.Digit) Then
      Jrn_Add(, {"⛔" & "   Erreur en " & U_Coord(Cellule) & " " & CStr(Cdd.Digit) & " est la solution. "})
      Return False
    End If
    Return True
  End Function
End Module