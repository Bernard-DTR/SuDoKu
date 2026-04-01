Module DB_00
  '--------------------------------------------------------------------------------
  ' Classe Candidate
  '--------------------------------------------------------------------------------
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

  '--------------------------------------------------------------------------------
  ' Classe ChainNode
  '--------------------------------------------------------------------------------
  Public Class ChainNode
    Public Property CandidateID As Integer
    Public Property ParentID As Integer
    Public Property Depth As Integer
    Public Property IsStrong As Boolean   ' True = lien fort, False = lien faible
    Public IsTrue As Boolean
    Public Sub New(cid As Integer, parent As Integer, depth As Integer, strong As Boolean, isTrue As Boolean)
      Me.CandidateID = cid
      Me.ParentID = parent
      Me.Depth = depth
      Me.IsStrong = strong
      Me.IsTrue = isTrue
    End Sub
  End Class

End Module