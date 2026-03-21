Module DB_01
  Public Class Candidate
    ' Identifiant unique (0 à 728)
    Public Property ID As Integer
    ' Position dans la grille
    Public Property Row As Integer   ' 1 à 9
    Public Property Col As Integer   ' 1 à 9
    Public Property Block As Integer ' 1 à 9
    ' Valeur du candidat (1 à 9)
    Public Property Digit As Integer
    ' État logique
    Public Property IsActive As Boolean   ' encore possible ?
    Public Property IsSolved As Boolean   ' fait partie de la solution 
    ' Liens logiques (pour whips/braids)
    Public Property StrongLinks As New List(Of Integer) ' IDs des candidats liés
    Public Property WeakLinks As New List(Of Integer)
    ' Pour les chaînes : parent, profondeur, etc.
    Public Property ChainParent As Integer = -1
    Public Property ChainDepth As Integer = 0
    Public Sub New(id As Integer, r As Integer, c As Integer, d As Integer)
      Me.ID = id
      Me.Row = r
      Me.Col = c
      Me.Digit = d
      Me.Block = ((r - 1) \ 3) * 3 + ((c - 1) \ 3) + 1
      Me.IsActive = True
      Me.IsSolved = False
    End Sub
  End Class

  Dim AllCandidates(728) As Candidate ' Tableau des 729 Candidate
  Public DB_Passage As Integer

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
      Trace(Cdd)
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
        Trace(Cdd)
      End If
    Next
    Jrn_Add_Yellow("IsActive " & Nb_IsActive)
    Jrn_Add_Yellow("IsSolved " & Nb_IsSolved)
  End Sub
  Public Sub AllCandidates_SDK(AllCandidates() As Candidate, VI As Boolean)
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
        If VI Then U_temp(Cellule, 1) = Cdd.Digit.ToString()
        U_temp(Cellule, 2) = Cdd.Digit.ToString()
        U_temp(Cellule, 3) = Cnddts_Blancs
      Else
        ' Candidat non résolu : on ajoute le digit dans la chaîne
        Dim Candidats As String = U_temp(Cellule, 3)
        Candidats = ReplaceDigit(Candidats, Cdd.Digit)
        U_temp(Cellule, 3) = Candidats
      End If
    Next

    Array.Copy(U_temp, U, UNbCopy)
    Event_OnPaint_MAP = Proc_Name_Get()
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub
  Private Function ReplaceDigit(source As String, digit As Integer) As String
    Dim chars() As Char = source.ToCharArray()
    chars(digit - 1) = ChrW(48 + digit)   ' Remplace par "1"…"9"
    Return New String(chars)
  End Function

  Public Sub AllCandidates_Init(AllCandidates() As Candidate)
    ' Initialisation de AllCandidates(728)
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
    Dim index As Integer = 0
    For r As Integer = 1 To 9
      For c As Integer = 1 To 9
        Dim ch As Char = grid(index)
        index += 1
        If ch >= "1"c AndAlso ch <= "9"c Then
          Dim solvedDigit As Integer = CInt(ch.ToString())
          For d As Integer = 1 To 9
            Dim id As Integer = ((r - 1) * 81) + ((c - 1) * 9) + (d - 1)
            If d = solvedDigit Then
              AllCandidates(id).IsActive = True
              AllCandidates(id).IsSolved = True
            Else
              AllCandidates(id).IsActive = False
              AllCandidates(id).IsSolved = False
            End If
          Next
        Else
          ' Case vide : tous les candidats restent actifs
          For d As Integer = 1 To 9
            Dim id As Integer = ((r - 1) * 81) + ((c - 1) * 9) + (d - 1)
            AllCandidates(id).IsActive = True
            AllCandidates(id).IsSolved = False
          Next
        End If
      Next
    Next
  End Sub

  Public Function Index(r As Integer, c As Integer, d As Integer) As Integer
    Return (r - 1) * 81 + (c - 1) * 9 + (d - 1)
  End Function

  Public Sub PropagateSolvedCandidates(ByVal AllCandidates() As Candidate)
    ' Dès qu'un candidat est IsSolved, IsActive= False des candidats de la même case
    '                                                                de la même unité
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
  End Sub

End Module