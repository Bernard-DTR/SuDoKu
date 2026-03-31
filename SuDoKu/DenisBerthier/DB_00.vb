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

  '--------------------------------------------------------------------------------
  ' Classe Denis_Berthier_Stratégie
  '--------------------------------------------------------------------------------
  'Public Class DB_Stg_Cls 'Classe structurant les stratégies de Denis Berthier
  '  Public Property Code As String         ' 0 Code
  '  Public Property Type As String         ' 1 Type: Insertion, Exclusion, Non 
  '  Public Property Texte As String        ' 2 Texte                            
  '  'Constructeur paramétré
  '  Sub New(New_Code As String,
  '          New_Type As String,
  '          New_Texte As String)
  '    Code = New_Code
  '    Type = New_Type
  '    Texte = New_Texte
  '  End Sub
  'End Class

  'Public DB_Stg_List As New List(Of DB_Stg_Cls)        ' Liste des stratégies de Denis Berthier
  Public Solution As String

  'Public Sub DB_Stg_List_Init()
  '  'Initialisation de la Liste des Stratégies de Denis Berthier
  '  With DB_Stg_List
  '    .Add(New DB_Stg_Cls("vi   ", "N", "SDK_AllCandidate"))
  '    .Add(New DB_Stg_Cls("NS   ", "V", "NakedSingles"))
  '    .Add(New DB_Stg_Cls("HS   ", "V", "HiddenSingles"))
  '    .Add(New DB_Stg_Cls("LC   ", "C", "LockedCandidates"))

  '  End With
  'End Sub


End Module


'Quand tu reviendras avec tes observations, On pourra :
'affiner le traçage des whips/braids
'reconstruire les chaînes complètes pour les afficher proprement
'vérifier la cohérence des liens forts/faibles
'optimiser les DFS pour les profondeurs élevées
'ou même ajouter les variantes avancées (t-whips, z-chains, g-whips…)
'Tu avances avec une rigueur d'ingénieur, et ça se voit dans la qualité de ton code.


';C\Développement\SuDoKu_2026\S50_SDK\Interactif_Poubelle\SDK_E_00523.txt
'PR_Nom = SDK_E_00523
'PR_Ini = ...9....5.8..1..72..1.7.34.8..4..23.....671....5.....41....94.7.5...1....742.....
'PR_Val = ...9....5.8..1..72..1.7.34.8..4..23.....671....5.....41....94.7.5...1....742.....
'PR_Sol = 7..9...15.8..1..72..1.7.34.817495236....6715...51..7.41...594.7.5.741..3.742..5.1
'PR_Dlu = Dlu
'PR_Dls = 746923815389514672521678349817495236492367158635182794163859427258741963974236581
';Date_Heure de création       : 2025_11_16_04_59_25_091
';Version                      : V2026_2510 #429
';Plcy_Strg_Profondeur         : UOBTXYSJZKQ
';Contrainte                   : S3
';Nb_Cellules_Demandées        : 28
';Str: 0-CdU 1-CdO 2-Cbl 3-Tpl 4-Xwg 5-XYw 6-Swf 7-Jly 8-XYZ 9-SKy 10-Unq 
';Crt:     0     0     0     0     0     0     0     0     0     0     0 
';Slv:     5    11     3     0     5     0     0     0     0     0     0 
';...

' je ne comprends pas le traçage 
' il y a une erreur c
' comment tester tous les cas de figures ? 