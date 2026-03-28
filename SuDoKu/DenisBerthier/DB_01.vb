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

  Public Class DB_Stg_Cls 'Classe structurant les stratégies de Denis Berthier
    Public Property Code As String         ' 0 Code
    Public Property Type As String         ' 1 Type: Insertion, Exclusion, Non 
    Public Property Texte As String        ' 2 Texte                            
    'Constructeur paramétré
    Sub New(New_Code As String,
            New_Type As String,
            New_Texte As String)
      Code = New_Code
      Type = New_Type
      Texte = New_Texte
    End Sub
  End Class
  Public DB_Stg_List As New List(Of DB_Stg_Cls)        ' List comportant les stratégies de Denis Berthier

  Public Sub DB_Stg_List_Init()
    'Initialisation de la Liste des Stratégies de Denis Berthier
    With DB_Stg_List
      .Add(New DB_Stg_Cls("vi   ", "N", "SDK_AllCandidate"))
      .Add(New DB_Stg_Cls("NS   ", "V", "NakedSingles"))
      .Add(New DB_Stg_Cls("HS   ", "V", "HiddenSingles"))
      .Add(New DB_Stg_Cls("LC   ", "C", "LockedCandidates"))

    End With
  End Sub




  Dim AllCandidates(728) As Candidate ' Tableau des 729 Candidate

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
        Trace("dsp", Cdd)
      End If
    Next
    Jrn_Add_Yellow("IsActive " & Nb_IsActive)
    Jrn_Add_Yellow("IsSolved " & Nb_IsSolved)
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
              Trace("vi", AllCandidates(id))
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

  Public Sub Trace(DB_Stg As String, Cdd As Candidate)
    With Cdd
      Dim S As String
      Dim Cellule As Integer = Wh_Cellule_RowCol(.Row - 1, .Col - 1)
      S = $"{DB_Stg,-5} {CStr(.ID),3} R{ .Row}_C{ .Col} Digit { .Digit}  Block { .Block} IsActive = { CStr(.IsActive),-5} IsSolved = { CStr(.IsSolved),-5}  {U_Coord(Cellule)} :{ .Digit}"
      Jrn_Add_Orange(S)
    End With
  End Sub


End Module