Option Strict On
Option Explicit On
Public Class XLink_Cls
  ' Classe structurant un Lien
  '   La structure d'un Lien comporte 2 cellules, le lien sera Cel(0) → Cel(1)
  '   puis un tableau de 9 candidats.
  '      En général Cdd(0) et Cdd(1) concerneront Cel(0)
  '                 Cdd(2) et Cdd(3) concerneront Cel(1) et Cdd(4) sera le candidat commun à Cel(0) et à Cel(1)
  '      Le Type de Lien sera "S" pour Lien Fort (Strong Link) ou "W" pour Lien Faible (Weak Link)
  '      L'Unité sera Rowx/Colx/Regx c'est à dire l'emplacement du lien dans la rangée x, dans la colonne x ou dans le région x
  '      La Composition sera une chaîne de caractères indiquant la composition du lien
  '      (ex: "024"  pour un lien fort entre 2 cellules contenant les candidats 0, 2 et 4)
  '           "0123" pour un lien faible entre 2 cellules contenant les candidats 0, 1, 2 et 3
  '           XLink_Str_New(XLink As XLink_Cls) As String retourne une chaîne lisible de XLink_Cls en utilisant la composition
  Public Property Cel As Integer() = {-1, -1}
  Public Property Cdd As String() = Enumerable.Repeat("0", 9).ToArray()
  Public Property Type As String
  Public Property Unité As String = "#"
  Public Property Composition As String = "#"
End Class


Public Class GLink_Cls
  ' Mise en place le 19/12/2025
  ' Classe structurant un Lien
  '   La structure d'un Lien comporte 2 cellules, le lien sera Cel(0) → Cel(1)
  '   puis un tableau de 9 candidats.
  '      En général Cdd(0) et Cdd(1) concerneront Cel(0)
  '                 Cdd(2) et Cdd(3) concerneront Cel(1) et Cdd(4) sera le candidat commun à Cel(0) et à Cel(1)
  '      Le Type de Lien sera "S" pour Lien Fort (Strong Link) ou "W" pour Lien Faible (Weak Link)
  '      L'Unité sera Rowx/Colx/Regx c'est à dire l'emplacement du lien dans la rangée x, dans la colonne x ou dans le région x
  '      La Composition sera une chaîne de caractères indiquant la composition des candidats du lien
  Public Property Cel As Integer() = {-1, -1}
  Public Property Cdd As String() = Enumerable.Repeat("0", 9).ToArray()
  '               équivalent à      Cdd = {"0", "0", "0", "0", "0", "0", "0", "0", "0"}
  Public Property Type As String = "#"
  Public Property Unité As String = "#"
  Public Property Cdd_Composition As String = "#"
End Class
