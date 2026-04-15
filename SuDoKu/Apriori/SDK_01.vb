Module SDK_01
  '--------------------------------------------------------------------------------
  ' Classe SdkU
  '--------------------------------------------------------------------------------
  Public Class SdkU
    Public Property ID As Integer          ' Identifiant unique (0 à 80)
    Public Property Row As Integer         ' 1 à 9
    Public Property Col As Integer         ' 1 à 9
    Public Property Coord As String        ' Coordonnées Rr_Cc (ID)
    Public Property Block As Integer       ' 1 à 9
    Public Property Valeur As Integer      ' Valeur donnée ou trouvée
    Public Property Candidats As String    ' Candidats éligibles "123456789"  ou "         " si résolu
    Public Property IsGiven As Boolean     ' Valeur initiale
    Public Property IsActive As Boolean    ' encore possible ?
    Public Property IsSolved As Boolean    ' fait partie de la solution 
    ' Liens logiques (pour whips/braids)
    Public Sub New(id As Integer, r As Integer, c As Integer, v As Integer)
      Me.ID = id
      Me.Row = r
      Me.Col = c
      Me.Valeur = v
      Me.Block = ((r - 1) \ 3) * 3 + ((c - 1) \ 3) + 1
      Me.IsActive = True
      Me.IsSolved = False
    End Sub
  End Class

End Module
