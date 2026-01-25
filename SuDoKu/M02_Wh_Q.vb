Option Strict On
Option Explicit On

Module M02_Wh_Q
  '-------------------------------------------------------------------------------
  ' Wh What Réservé aux fonctions
  ' Hw How  Réservé aux procédures
  '-------------------------------------------------------------------------------


  Public Function Wh_3RowReg(Région As Integer) As Integer()
    ' Retourne les 3 Rangées d'une Région
    Dim Wh_3Row(2) As Integer
    Wh_3Row(0) = -1 : Wh_3Row(1) = -1 : Wh_3Row(2) = -1
    Select Case Région
      Case 0, 1, 2 : Wh_3Row(0) = 0 : Wh_3Row(1) = 1 : Wh_3Row(2) = 2
      Case 3, 4, 5 : Wh_3Row(0) = 3 : Wh_3Row(1) = 4 : Wh_3Row(2) = 5
      Case 6, 7, 8 : Wh_3Row(0) = 6 : Wh_3Row(1) = 7 : Wh_3Row(2) = 8
    End Select
    Return Wh_3Row
  End Function

  Public Function Wh_3ColReg(Région As Integer) As Integer()
    ' Retourne les 3 Colonnes d'une Région
    Dim Wh_3Col(2) As Integer
    Wh_3Col(0) = -1 : Wh_3Col(1) = -1 : Wh_3Col(2) = -1
    Select Case Région
      Case 0, 3, 6 : Wh_3Col(0) = 0 : Wh_3Col(1) = 1 : Wh_3Col(2) = 2
      Case 1, 4, 7 : Wh_3Col(0) = 3 : Wh_3Col(1) = 4 : Wh_3Col(2) = 5
      Case 2, 5, 8 : Wh_3Col(0) = 6 : Wh_3Col(1) = 7 : Wh_3Col(2) = 8
    End Select
    Return Wh_3Col
  End Function

  Public Function Wh_Cell_Nb_Candidats(U_temp(,) As String, Cellule As Integer) As Integer
    ' Retourne le nombre de candidats dans une cellule
    ' =  0 La cellule est remplie
    ' = -1 La cellule est vide et n'a plus de candidats (#Erreur)
    ' =  x (1 à 9) la cellule a x candidats
    ' = -2 autres cas
    If U_temp(Cellule, 2) <> " " Then Return 0
    If U_temp(Cellule, 2) = " " And U_temp(Cellule, 3) = Cnddts_Blancs Then Return -1
    If U_temp(Cellule, 2) = " " And U_temp(Cellule, 3) <> Cnddts_Blancs Then
      'Return U_temp(Cellule, 3).Count(Function(ch As Char) Char.IsDigit(ch))
      'c précise que c'est un Char et non un String
      Return U_temp(Cellule, 3).Count(Function(ch As Char) ch >= "1"c AndAlso ch <= "9"c)
    End If
    Return -2
  End Function

  Public Function Wh_Candidat_Unité(U_temp(,) As String, Cellule As Integer, Candidat As String, Code_LCR As String) As Integer
    ' Correction 02/07/2025 Code_LCR récupère désormais Row, Col ou Reg
    Dim Nb As Integer = 0
    Dim Grp(0 To 8) As Integer
    Select Case Code_LCR
      Case "Row" : Grp = U_9CelRow(U_Row(Cellule)) 'Comporte les 9 cellules de la ligne
      Case "Col" : Grp = U_9CelCol(U_Col(Cellule)) '                              colonne
      Case "Reg" : Grp = U_9CelReg(U_Reg(Cellule)) '                              région
    End Select
    Dim Cel As Integer
    For j As Integer = 0 To 8
      Cel = Grp(j)
      If U_temp(Cel, 3).Contains(Candidat) Then Nb += 1
    Next j
    Return Nb
  End Function

  Public Function Wh_Grid_Nb_Cellules_Initiales(U_temp(,) As String) As Integer
    'Retourne le Nombre de Cellules Initiales dans une grille
    Dim Nb_Cellules_Initiales As Integer = 0
    For i As Integer = 0 To 80
      If U_temp(i, 1) <> " " Then Nb_Cellules_Initiales += 1
    Next i
    Return Nb_Cellules_Initiales
  End Function
  Public Function Wh_Grid_Nb_Cellules_Vides(U_temp(,) As String) As Integer
    'Retourne le Nombre de Cellules vides dans une grille
    Dim Nb_Cellules_Vides As Integer = 0
    For i As Integer = 0 To 80
      If U_temp(i, 2) = " " Then Nb_Cellules_Vides += 1
    Next i
    Return Nb_Cellules_Vides
  End Function
  Public Function Wh_Grid_Nb_Cellules_Remplies(U_temp(,) As String) As Integer
    'Retourne le Nombre de Cellules remplies dans une grille
    Dim Nb_Cellules_Remplies As Integer = 0
    For i As Integer = 0 To 80
      If U_temp(i, 2) <> " " Then Nb_Cellules_Remplies += 1
    Next i
    Return Nb_Cellules_Remplies
  End Function
  Public Function Wh_Grid_Nb_Candidats(U_temp(,) As String) As Integer
    'La fonction retourne le nombre total de candidats de la grille
    Dim Nb_Candidats As Integer = 0

    For i As Integer = 0 To 80
      Dim n As Integer = Wh_Cell_Nb_Candidats(U_temp, i)
      If n > 0 Then Nb_Candidats += n
    Next i
    Return Nb_Candidats
  End Function
  Public Function Wh_Nb_Cells_Nb_Candidats_Identiques(U_temp(,) As String, Cellule_A As Integer, Cellule_B As Integer) As Integer
    'La fonction retourne le nombre de candidats communs entre 2 cellules
    Dim n As Integer = 0
    Dim Cdd_A As String = U_temp(Cellule_A, 3)
    Dim Cdd_B As String = U_temp(Cellule_B, 3)
    For cd As Integer = 1 To 9
      If Cdd_A.Substring(cd - 1, 1) = Cdd_B.Substring(cd - 1, 1) Then
        If (Cdd_A.Substring(cd - 1, 1) <> " " And Cdd_B.Substring(cd - 1, 1) <> " ") Then n += 1
      End If
    Next cd
    Return n
  End Function

  Public Function Wh_Pourcentage() As String
    'Calcule le Pourcentage de cellules remplies dans la grille
    Return CStr((Wh_Nb_Cell(U).Remplies * 100) \ 81) & " %"
  End Function
  Public Function Wh_RandomCelluleVide() As Integer
    'Détermine une Cellule Vide au hasard
    'Si aucune cellule n'est vide, retourne -1
    Dim Cdd_Vide_Clct As New Collection
    '1 Remplissage d'une collection des cellules vides
    For i As Integer = 0 To 80
      If U(i, 2) = " " Then Clct_Add(Cdd_Vide_Clct, i)
    Next i
    '2 Détermination d'une cellule vide au hasard
    Wh_RandomCelluleVide = Clct_Random(Cdd_Vide_Clct)
  End Function
  Public Function Wh_9CellulesRow_Row(Row As Integer) As Integer()
    'Détermine les 9 Cellules de la Ligne, à partir de la Ligne
    Dim Grp(0 To 8) As Integer
    For i As Integer = 0 To 8 : Grp(i) = (Row * 9) + i : Next i
    Return Grp
  End Function
  Public Function Wh_9CellulesColumn_Col(Col As Integer) As Integer()
    'Détermine les 9 Cellules de la Colonne, à partir de la Colonne
    Dim Grp(0 To 8) As Integer
    For i As Integer = 0 To 8 : Grp(i) = Col + (9 * i) : Next i
    Return Grp
  End Function
  Public Function Wh_9CellulesRégion_Rég(Région As Integer) As Integer()
    'Détermine les 9 Cellules de la Région, à partir de la Région
    Dim Grp(0 To 8) As Integer
    Select Case Région
      Case 0 : Grp = {0, 1, 2, 9, 10, 11, 18, 19, 20}
      Case 1 : Grp = {3, 4, 5, 12, 13, 14, 21, 22, 23}
      Case 2 : Grp = {6, 7, 8, 15, 16, 17, 24, 25, 26}
      Case 3 : Grp = {27, 28, 29, 36, 37, 38, 45, 46, 47}
      Case 4 : Grp = {30, 31, 32, 39, 40, 41, 48, 49, 50}
      Case 5 : Grp = {33, 34, 35, 42, 43, 44, 51, 52, 53}
      Case 6 : Grp = {54, 55, 56, 63, 64, 65, 72, 73, 74}
      Case 7 : Grp = {57, 58, 59, 66, 67, 68, 75, 76, 77}
      Case 8 : Grp = {60, 61, 62, 69, 70, 71, 78, 79, 80}
    End Select
    Return Grp
  End Function
  Public Function Wh_Région_Cel(Cellule As Integer) As Integer
    'Détermine la Région de 0 à 8 de la cellule
    Dim Région As Integer = -1
    Select Case Cellule
      Case 0, 1, 2, 9, 10, 11, 18, 19, 20 : Région = 0
      Case 3, 4, 5, 12, 13, 14, 21, 22, 23 : Région = 1
      Case 6, 7, 8, 15, 16, 17, 24, 25, 26 : Région = 2
      Case 27, 28, 29, 36, 37, 38, 45, 46, 47 : Région = 3
      Case 30, 31, 32, 39, 40, 41, 48, 49, 50 : Région = 4
      Case 33, 34, 35, 42, 43, 44, 51, 52, 53 : Région = 5
      Case 54, 55, 56, 63, 64, 65, 72, 73, 74 : Région = 6
      Case 57, 58, 59, 66, 67, 68, 75, 76, 77 : Région = 7
      Case 60, 61, 62, 69, 70, 71, 78, 79, 80 : Région = 8
    End Select
    Return Région
  End Function
  Public Function Wh_Cellule_ColRow(Cl As Integer, Rw As Integer) As Integer
    'Donne le numéro de la cellule
    Return (Rw * 9) + Cl
  End Function
  Public Function Wh_Cellule_RowCol(Rw As Integer, Cl As Integer) As Integer
    'Donne le numéro de la cellule
    Return (Rw * 9) + Cl
  End Function
  Public Function Is_SameRow(Cel1 As Integer, Cel2 As Integer) As Boolean
    Return U_Row(Cel1) = U_Row(Cel2)
  End Function
  Public Function Is_SameCol(Cel1 As Integer, Cel2 As Integer) As Boolean
    Return U_Col(Cel1) = U_Col(Cel2)
  End Function
  Public Function Is_SameReg(Cel1 As Integer, Cel2 As Integer) As Boolean
    Return U_Reg(Cel1) = U_Reg(Cel2)
  End Function
  Public Function Is_SameUnité(Cel1 As Integer, Cel2 As Integer) As String()
    ' Correction 02/07/2025 La fonction retourne le Code_LCR Row, Col ou Reg et LCR 
    ' La fonction retourne le Code_LCR et LCR; LCR est augmenté de 1, pour simplifier la lecture.
    ' à chaque utilisation, il est testé auparavant Cel1 <> Cel2
    If Cel1 = Cel2 Then Return {"=", "="}
    If Is_SameRow(Cel1, Cel2) Then Return {"Row", CStr(U_Row(Cel1) + 1)}
    If Is_SameCol(Cel1, Cel2) Then Return {"Col", CStr(U_Col(Cel1) + 1)}
    ' Si les 2 cellules sont dans la même région, les tests effectués avant
    '                   entraînent que les 2 cellules sont dans la même région,
    '                   MAIS pas dans la même Row, ni dans la même Col
    If Is_SameReg(Cel1, Cel2) Then Return {"Reg", CStr(U_Reg(Cel1) + 1)}
    Return {"#", "#"}
  End Function

  Public Function Is_LinkType(U_temp(,) As String, Cel1 As Integer, Cel2 As Integer, Candidat As String) As String
    'Retourne le type de lien entre les 2 cellules 
    Dim Code_LCR As String = ""

    If Is_SameRow(Cel1, Cel2) Then
      Code_LCR = "Row"
    ElseIf Is_SameCol(Cel1, Cel2) Then
      Code_LCR = "Col"
    ElseIf Is_SameReg(Cel1, Cel2) Then
      Code_LCR = "Reg"
    Else
      Return "#" ' Aucun lien
    End If

    Dim Grp(8) As Integer
    Select Case Code_LCR
      Case "Row" : Grp = U_9CelRow(U_Row(Cel1))
      Case "Col" : Grp = U_9CelCol(U_Col(Cel1))
      Case "Reg" : Grp = U_9CelReg(U_Reg(Cel1))
    End Select

    Dim Nb As Integer = Grp.Count(Function(c) U_temp(c, 3).Contains(Candidat))

    Return If(Nb = 2, "S", "W")
  End Function
  Public Function Wh_NumérodelaCellule_RegNreg(Région As Integer, NRégion As Integer) As Integer
    'Donne le numéro de la cellule à partir de la région et du N° de Région
    Dim Grp(0 To 8) As Integer
    Select Case Région
      Case 0 : Grp = {0, 1, 2, 9, 10, 11, 18, 19, 20}
      Case 1 : Grp = {3, 4, 5, 12, 13, 14, 21, 22, 23}
      Case 2 : Grp = {6, 7, 8, 15, 16, 17, 24, 25, 26}
      Case 3 : Grp = {27, 28, 29, 36, 37, 38, 45, 46, 47}
      Case 4 : Grp = {30, 31, 32, 39, 40, 41, 48, 49, 50}
      Case 5 : Grp = {33, 34, 35, 42, 43, 44, 51, 52, 53}
      Case 6 : Grp = {54, 55, 56, 63, 64, 65, 72, 73, 74}
      Case 7 : Grp = {57, 58, 59, 66, 67, 68, 75, 76, 77}
      Case 8 : Grp = {60, 61, 62, 69, 70, 71, 78, 79, 80}
    End Select
    Wh_NumérodelaCellule_RegNreg = Grp(NRégion)
  End Function

#Region "Fonctions diverses"
  Public Function Wh_Cellule_Pt_IA(Pt As Point) As Integer
    'Retourne le numéro de la cellule en fonction du point clické
    'En fonction de e, le point peut être un Point (MouseDown, MouseClick)
    '                           ou        un PointToClient (Drag And Drop)
    'Il faut que le point clické soit dans un des 81 squares
    ' -1 est retourné si le point n'est pas dans la grille 
    '                             ou sur un séparateur de cellule
    '                             ou sur un séparateur de région
    '                             ou sur l'entourage de la grille

    Return Array.FindIndex(Sqr_Cel, Function(cel) cel.Contains(Pt))
  End Function

  Public Function Wh_Cellule_Candidat_Pt_IA(Cellule As Integer, Pt As Point) As Integer
    'Donne la zone du Candidat clické en fonction du point ou retourne -1
    'En fonction de e, le point peut être un Point (MouseDown, MouseClick)
    '                           ou        un PointToClient (Drag And Drop)
    Dim Candidat_Clic As Integer = -1
    If Cellule < 0 Or Cellule > 80 Then
      Return Candidat_Clic
    End If

    '29/07/2023 Agrandissement du rectangle du candidat  
    For cdd As Integer = 1 To 9
      Dim cdd_n As Integer = (Cellule * 10) + cdd
      Dim Sqr_Cdd_n As New Rectangle(Sqr_Cdd(cdd_n).X - 1, Sqr_Cdd(cdd_n).Y - 1, Sqr_Cdd(cdd_n).Width + 2, Sqr_Cdd(cdd_n).Height + 2)
      If Sqr_Cdd_n.Contains(Pt) Then
        Candidat_Clic = cdd
        Exit For
      End If
    Next cdd
    Return Candidat_Clic
  End Function

  Public Function Wh_Candidats_Unité(Grp() As Integer) As String
    'Retourne les candidats restants d'une Unité Colonne, Ligne ou Région
    Dim Cdd_Unité As String = Cnddts
    For n As Integer = 0 To 8
      If U(Grp(n), 2) <> " " Then
        Dim V As Integer = CInt(U(Grp(n), 2))
        Mid$(Cdd_Unité, V, 1) = " "
      End If
    Next n
    'Si l'unité est remplie, un ToolTipText vide est affiché
    'If Cdd_Unité = Cnddts_Blancs Then Cdd_Unité = Nothing
    Return Cdd_Unité
  End Function
#End Region
End Module