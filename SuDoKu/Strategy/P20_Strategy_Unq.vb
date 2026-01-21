Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------
' Stratégie des Uniqueness Test                           
' Préfixe       Unq
' Documentation issue de http://sudopedia.enjoysudoku.com/Uniqueness_Test.html
' La stratégie Unq R1 ne peut pas être intégrée dans le calcul de Résolution Strategy_Upd_BTXYSJZKQ
' car elle repose sur le fait que la solution est unique
' UTILISATION:
' Stratégie des Uniqueness G4_Grid_Stratégie_Unq() , Mnu07_Strategy("Q"), Strategy_Unq_Test()
'   L'affichage de la stratégie Unq est effectuée à/p du Menu Stratégie ou de la BO lettre Q
'   L'affichage est correct en Affichage Simple et Aide Graphique
' Production d'une Grille 
'   Pzzl_Crt  Production_Type = "P"
'   Pzzl_Slv  Production_Type = "P"
' Résolution d'une Grille 
'   Pzzl_Slv  Production_Type = "S"
'
'-------------------------------------------------------------------------------
Friend Module P20_Strategy_Unq

  Function Strategy_Unq(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "Unq"
    Dim Sous_Stratégie As String = "Unq"
    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Unq
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 
    Strategy_Unq_Rectangle_1(U_temp, Strategy_Rslt)
    Strategy_Unq_Rectangle_2(U_temp, Strategy_Rslt)
    Return Strategy_Rslt
  End Function

  Sub Strategy_Unq_Rectangle_1(U_temp(,) As String, ByRef Strategy_Rslt(,) As String)
    'La stratégie Unq R1 a été testée avec 15 exemples de Hodoku
    Dim Stratégie As String = "Unq"
    Dim Sous_Stratégie As String = "R1"
    Dim Typologie As Integer

    Dim Unq1_List As New List(Of Unq1_Cls)
    Dim Unq2_List As New List(Of Unq2_Cls)
    'Unq2_List.Count = 0 lorsque la List est vide

    ' Analyse des Cellules à 2 candidats
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = 2 Then
        Unq1_List.Add(New Unq1_Cls(New_Cellule:=i, New_Candidats:=U_temp(i, 3), New_Occurences:=0))
      End If
    Next i
    If Unq1_List.Count = 0 Then GoTo Strategy_Unq_End

    ' Calcul du nombre d'occurences
    For i As Integer = 0 To Unq1_List.Count - 1
      Dim n As Integer = 0
      For j As Integer = 0 To Unq1_List.Count - 1
        If Unq1_List.Item(i).Candidats = Unq1_List.Item(j).Candidats Then n += 1
      Next j
      Unq1_List.Item(i).Occurences = n
    Next i

    Dim Cel_A, Cel_B, Cel_C, Cel_D As Integer

    ' Etude des groupes de 3 _ Il faut 3 Cellules
    For i As Integer = 0 To Unq1_List.Count - 1
      If Unq1_List.Item(i).Occurences <= 2 Then Continue For

      Cel_A = Unq1_List.Item(i).Cellule
      Dim Cdd As String = Unq1_List.Item(i).Candidats
      For j As Integer = i + 1 To Unq1_List.Count - 1
        If Cdd = Unq1_List.Item(j).Candidats Then
          Cel_B = Unq1_List.Item(j).Cellule
          For k As Integer = j + 1 To Unq1_List.Count - 1
            If Cdd = Unq1_List.Item(k).Candidats Then
              Cel_C = Unq1_List.Item(k).Cellule
              Unq2_List.Add(New Unq2_Cls(New_Cellule_A:=Cel_A, New_Cellule_B:=Cel_B, New_Cellule_C:=Cel_C, New_Candidats:=Cdd))
            End If
          Next k
        End If
      Next j
    Next i

    ' Etude rectangulaire des points A, B et C 
    '
    '  Pour que les 3 points soient en rectangle, CNS : une seule Ligne et une seule Colonne
    '  Seuls 4 Shémas favorables sont possibles
    '  Cellule_A < Cellule_B < Cellule_C
    '  A-----B       A-----B        A                A
    '  |1                 2|        |                |
    '  |                   |        |3              4|
    '  C                   C        B-----C    B-----C
    'L'ordonnancement des cellules sera important pour tracer un rectangle qui est programmé "Croissant"
    For i As Integer = 0 To Unq2_List.Count - 1
      Cel_A = Unq2_List.Item(i).Cellule_A
      Cel_B = Unq2_List.Item(i).Cellule_B
      Cel_C = Unq2_List.Item(i).Cellule_C
      Dim Row_A As Integer = U_Row(Cel_A) : Dim Row_B As Integer = U_Row(Cel_B)
      Dim Row_C As Integer = U_Row(Cel_C) : Dim Col_A As Integer = U_Col(Cel_A)
      Dim Col_B As Integer = U_Col(Cel_B) : Dim Col_C As Integer = U_Col(Cel_C)
      Typologie = -1
      Cel_D = -1
      'Les 4 cellules sont en rectangle si
      If (Row_A = Row_B And Col_A = Col_C) Then
        Cel_D = Wh_Cellule_RowCol(Row_C, Col_B) ' Rectangle 1
        Typologie = 1
      End If
      If (Row_A = Row_B And Col_B = Col_C) Then
        Cel_D = Wh_Cellule_RowCol(Row_C, Col_A) ' Rectangle 2
        Typologie = 2
      End If
      If (Row_B = Row_C And Col_A = Col_B) Then
        Cel_D = Wh_Cellule_RowCol(Row_A, Col_C) ' Rectangle 3
        Typologie = 3
      End If
      If (Row_B = Row_C And Col_A = Col_C) Then
        Cel_D = Wh_Cellule_RowCol(Row_A, Col_B) ' Rectangle 4
        Typologie = 4
      End If
      If Cel_D = -1 Then Continue For

      'Il faut encore que les 4 cellules n'appartiennent qu'à 2 régions différentes
      Dim Reg_A As Integer = U_Reg(Cel_A)
      Dim Reg_B As Integer = U_Reg(Cel_B)
      Dim Reg_C As Integer = U_Reg(Cel_C)
      Dim Reg_D As Integer = U_Reg(Cel_D)
      If (Reg_A = Reg_B And Reg_C = Reg_D And Reg_A <> Reg_C) _
      Or (Reg_A = Reg_C And Reg_B = Reg_D And Reg_A <> Reg_B) _
      Or (Reg_A = Reg_D And Reg_B = Reg_C And Reg_A <> Reg_B) Then
      Else Continue For
      End If

      'La Cellule_D DOIT à présent comporter au moins 3 candidats
      If Wh_Cell_Nb_Candidats(U_temp, Cel_D) < 3 Then Continue For
      'Les Cellule_A_B_C comportent chacune les mêmes 2 candidats
      'la  Cellule_D doit comporter ces 2 candidats et d'autres
      '                             ces 2 candidats seront à supprimer
      Dim Cdd_A As String = U_temp(Cel_A, 3)
      Dim Cdd_D As String = U_temp(Cel_D, 3)

      Dim Cdd_Identique_Nb As Integer
      Dim Cdd_Différent_Nb As Integer
      Dim Cdd_Différent As String = Cnddts_Blancs
      For c As Integer = 1 To 9
        'Les espaces ne sont pas traités
        If Mid$(Cdd_D, c, 1) <> " " And Mid$(Cdd_D, c, 1) = Mid$(Cdd_A, c, 1) Then
          Cdd_Identique_Nb += 1
        Else
          Cdd_Différent_Nb += 1
          Mid$(Cdd_Différent, c, 1) = CStr(c)
        End If
      Next c

      If Cdd_Identique_Nb = 2 Then          'Cas Uniqueness 1
        Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
        Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
        Select Case Typologie
          Case 1     ' A, B, D, C
            Cel45(0) = CStr(Cel_A)
            Cel45(1) = CStr(Cel_B)
            Cel45(2) = CStr(Cel_D)
            Cel45(3) = CStr(Cel_C)
          Case 2     ' A, B, C, D
            Cel45(0) = CStr(Cel_A)
            Cel45(1) = CStr(Cel_B)
            Cel45(2) = CStr(Cel_C)
            Cel45(3) = CStr(Cel_D)
          Case 3     ' A, D, C, B
            Cel45(0) = CStr(Cel_A)
            Cel45(1) = CStr(Cel_D)
            Cel45(2) = CStr(Cel_C)
            Cel45(3) = CStr(Cel_B)
          Case 4     ' D, A, C, B
            Cel45(0) = CStr(Cel_D)
            Cel45(1) = CStr(Cel_A)
            Cel45(2) = CStr(Cel_C)
            Cel45(3) = CStr(Cel_B)
        End Select
        Exc45(0) = CStr(Cel_D)
        'Premier cas où Strategy_Rslt(5, x) qui comporte Le Candidat concerné par la stratégie comporte 2 candidats
        Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie & CStr(Typologie), "X", "-1", Cdd_A, Cel45, Exc45)
      End If
    Next i
Strategy_Unq_End:
    Exit Sub
  End Sub

  Sub Strategy_Unq_Rectangle_2(U_temp(,) As String, ByRef Strategy_Rslt(,) As String)
    Dim Stratégie As String = "Unq"
    Dim Sous_Stratégie As String = "R2"
    Dim Typologie As String

    Dim Unq1_List As New List(Of Unq1_Cls)
    Dim Unq3_List As New List(Of Unq3_Cls)
    'Unq3_List.Count.ToString() = 0 lorsque la List est vide

    ' La typologie Unq R2 est soit Horizontale, soit Verticale, avec pour H les 3 candidats à  Droite ou à  Gauche
    '                                                           et   pour V les 3 candidats en Haut   ou en Bas.
    '
    ' AB     ABC       ABC      AB        AB        AB         ABC       ABC
    '
    '
    '
    ' AB     ABC       ABC      AB        ABC       ABC        AB        AB
    '
    ' R2HD             R2HG               R2VB                 R2VH
    '
    ' A      C         C        A         A         B          C         D
    '
    '
    '
    ' B      D         D        B         C         D          A         B
    '
    ' Analyse des Cellules à 2 candidats
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = 2 Then
        Unq1_List.Add(New Unq1_Cls(New_Cellule:=i, New_Candidats:=U_temp(i, 3), New_Occurences:=0))
      End If
    Next i
    If Unq1_List.Count = 0 Then GoTo Strategy_Unq_End

    ' Calcul du nombre d'occurences
    For i As Integer = 0 To Unq1_List.Count - 1
      Dim n As Integer = 0
      For j As Integer = 0 To Unq1_List.Count - 1
        If Unq1_List.Item(i).Candidats = Unq1_List.Item(j).Candidats Then n += 1
      Next j
      Unq1_List.Item(i).Occurences = n
    Next i
    'Strategy_Unq_Unq1List_Display(Unq1_List)

    Dim Cel_A, Cel_B, Cel_C, Cel_D As Integer

    ' Etude des groupes de 2 cellules à 2 candidats indentiques 
    For i As Integer = 0 To Unq1_List.Count - 1
      If Unq1_List.Item(i).Occurences <= 1 Then Continue For 'il faut au moins 2 cellules
      Cel_A = Unq1_List.Item(i).Cellule
      Dim Cdd As String = Unq1_List.Item(i).Candidats

      For j As Integer = i + 1 To Unq1_List.Count - 1
        Cel_B = Unq1_List.Item(j).Cellule
        'les candidats de Cel_A et Cel_B doivent être identiques 
        If Cdd <> Unq1_List.Item(j).Candidats Then Continue For
        Dim Row_A As Integer = U_Row(Cel_A) : Dim Row_B As Integer = U_Row(Cel_B)
        Dim Col_A As Integer = U_Col(Cel_A) : Dim Col_B As Integer = U_Col(Cel_B)
        Typologie = "#"       ' Typologie provisoire et incomplète
        If (Row_A = Row_B) Then Typologie = "V"  'Combinaison: AB  Verticale
        If (Col_A = Col_B) Then Typologie = "H"  'Combinaison: AB  Horizontale
        If Typologie <> "#" Then
          Unq3_List.Add(New Unq3_Cls(Cel_A, Cel_B, -1, Cdd, Typologie))
        End If
      Next j
    Next i
    'Strategy_Unq_Unq3List_Display(Unq3_List)

    ' Etude des groupes de 2 cellules à 2 candidats indentiques avec 2 autres cellules à 3 candidats
    ' Il faut trouver pour Cel_A une cellule avec un troisième candidat 
    '                      cel_B une cellule avec un troisième candidat
    ' et que ces 2 cellules soient dans une même région, dans la même ligne ou dans la même colonne
    For i As Integer = 0 To Unq3_List.Count - 1
      Cel_A = Unq3_List.Item(i).Cellule_A               ' Les 2 cellules 
      Cel_B = Unq3_List.Item(i).Cellule_B               ' à l'horizontale ou à la verticale
      Dim Cdd2 As String = Unq3_List.Item(i).Candidats  ' et les candidats identiques de ces 2 cellules
      Typologie = Unq3_List.Item(i).Typologie           ' pour le moment H ou V uniquement

      'On cherche une cellule avec un troisième candidat
      For cd As Integer = 1 To 9
        ' If Cdd2.Contains(CStr(cd)) Then Continue For ' est équivalent à la condition InStr
        If InStr(1, Cdd2, CStr(cd)) <> 0 Then Continue For
        'Le 3 èm candidat est ajouté aux 2 candidats recherchés
        Dim Cdd3 As String = Cdd2
        Mid$(Cdd3, cd, 1) = CStr(cd) ' Donc : Cdd3 = Cdd2 + cd
        Dim Row_A As Integer = U_Row(Cel_A) : Dim Row_B As Integer = U_Row(Cel_B)
        Dim Col_A As Integer = U_Col(Cel_A) : Dim Col_B As Integer = U_Col(Cel_B)
        Dim Grp(0 To 8) As Integer
        Dim Cellule As Integer
        Select Case Typologie
          Case "V" : Grp = U_9CelCol(Col_A)
          Case "H" : Grp = U_9CelRow(Row_A)
        End Select
        For j As Integer = 0 To 8
          Cellule = Grp(j)
          ' La cellule ne doit pas être Cel_A, elle doit être vide et le nombre de candidats doit être égal à 3
          If (Cellule = Cel_A) Or (U_temp(Cellule, 2) <> " ") Or (Wh_Cell_Nb_Candidats(U_temp, Cellule) <> 3) Then Continue For
          ' Il faut encore que les candidats recherchés soient Cdd3 ' Cdd3 = Cdd2 + cd
          If U_temp(Cellule, 3) <> Cdd3 Then Continue For
          Cel_C = Cellule     ' Cel_C est la première cellule trouvée
          ' Cel_D doit être positionnée "rectanglement parlant" et comporter également Cdd3
          Cel_D = -1
          If Typologie = "V" Then Cel_D = Wh_Cellule_ColRow(Col_B, U_Row(Cel_C))
          If Typologie = "H" Then Cel_D = Wh_Cellule_RowCol(Row_B, U_Col(Cel_C))
          If U_temp(Cel_D, 3) <> Cdd3 Then Continue For
          If U_Reg(Cel_A) <> U_Reg(Cel_B) And U_Reg(Cel_A) <> U_Reg(Cel_C) Then Continue For
          If U_Reg(Cel_A) = U_Reg(Cel_D) Then Continue For 'ce ne doit pas être une seule région

          'A ce niveau, on est dans une stratégie Unq R2
          If U_Row(Cel_A) < U_Row(Cel_C) Then Typologie &= "B"
          If U_Row(Cel_A) > U_Row(Cel_C) Then Typologie &= "H"
          If U_Col(Cel_A) < U_Col(Cel_C) Then Typologie &= "D"
          If U_Col(Cel_A) > U_Col(Cel_C) Then Typologie &= "G"
          Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
          Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice
          Select Case Typologie
            Case "HD"
              Cel45(0) = CStr(Cel_A)
              Cel45(1) = CStr(Cel_C)
              Cel45(2) = CStr(Cel_D)
              Cel45(3) = CStr(Cel_B)
            Case "HG"
              Cel45(0) = CStr(Cel_C)
              Cel45(1) = CStr(Cel_A)
              Cel45(2) = CStr(Cel_B)
              Cel45(3) = CStr(Cel_D)
            Case "VB"
              Cel45(0) = CStr(Cel_A)
              Cel45(1) = CStr(Cel_B)
              Cel45(2) = CStr(Cel_D)
              Cel45(3) = CStr(Cel_C)
            Case "VH"
              Cel45(0) = CStr(Cel_C)
              Cel45(1) = CStr(Cel_D)
              Cel45(2) = CStr(Cel_B)
              Cel45(3) = CStr(Cel_A)
          End Select
          Dim Grp_Exc(0 To 8) As Integer
          Dim Cel_Exc As Integer = -1
          ' 1 Exclusion pour la région
          If U_Reg(Cel_C) = U_Reg(Cel_D) Then
            Grp_Exc = U_9CelReg(U_Reg(Cel_D))
            For k As Integer = 0 To 8
              Cel_Exc = Grp_Exc(k)
              If Cel_Exc = Cel_C Or Cel_Exc = Cel_D Or U_temp(Cel_Exc, 2) <> " " Then Continue For
              If U_temp(Cel_Exc, 3).Contains(CStr(cd)) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_Exc)
              End If
            Next k
          End If
          ' 2 Exclusion pour la ligne ou pour la colonne
          Select Case Typologie
            Case "VB", "VH" : Grp_Exc = U_9CelRow(U_Row(Cel_D))
            Case "HD", "HG" : Grp_Exc = U_9CelCol(U_Col(Cel_D))
          End Select
          For k As Integer = 0 To 8
            Cel_Exc = Grp_Exc(k)
            If Cel_Exc = Cel_C Or Cel_Exc = Cel_D Or U_temp(Cel_Exc, 2) <> " " Then Continue For
            If U_temp(Cel_Exc, 3).Contains(CStr(cd)) Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_Exc)
            End If
          Next k
          If Exc45_n <> -1 Then
            Dim Candidats As String = Cnddts_Blancs
            Mid$(Candidats, cd, 1) = CStr(cd)
            Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie & Typologie, "X", "-1", Candidats, Cel45, Exc45)
          End If
        Next j
      Next cd
    Next i
Strategy_Unq_End:
    Exit Sub
  End Sub

  'Utilisation d'une Classe plutôt qu'une structure
  'https://learn.microsoft.com/fr-fr/dotnet/visual-basic/programming-guide/language-features/data-types/structures-And-classes

  Public Class Unq1_Cls 'Classe structurant les résultats de niveau 1
    Public Property Cellule As Integer               ' 1 Cellule
    Public Property Candidats As String              ' 3 Les Candidats 
    Public Property Occurences As Integer            ' 4 Le nombre d'occurences
    Sub New(New_Cellule As Integer,
            New_Candidats As String,
            New_Occurences As Integer)
      Cellule = New_Cellule
      Candidats = New_Candidats
      Occurences = New_Occurences
    End Sub
  End Class

  Public Class Unq2_Cls 'Classe structurant les résultats de niveau 2
    Public Property Cellule_A As Integer             ' 1 Cellule A
    Public Property Cellule_B As Integer             ' 2 Cellule B
    Public Property Cellule_C As Integer             ' 2 Cellule C
    Public Property Candidats As String              ' 3 Les Candidats 
    Sub New(New_Cellule_A As Integer,
            New_Cellule_B As Integer,
            New_Cellule_C As Integer,
            New_Candidats As String)
      Cellule_A = New_Cellule_A
      Cellule_B = New_Cellule_B
      Cellule_C = New_Cellule_C
      Candidats = New_Candidats
    End Sub
  End Class

  Public Class Unq3_Cls 'Classe structurant les résultats de niveau 2
    Public Property Cellule_A As Integer             ' 1 Cellule A
    Public Property Cellule_B As Integer             ' 2 Cellule B
    Public Property Cellule_C As Integer             ' 2 Cellule C
    Public Property Candidats As String              ' 3 Les Candidats 
    Public Property Typologie As String              ' 4 Typologie H/V GD HB
    Sub New(New_Cellule_A As Integer,
            New_Cellule_B As Integer,
            New_Cellule_C As Integer,
            New_Candidats As String,
            New_Typologie As String)
      Cellule_A = New_Cellule_A
      Cellule_B = New_Cellule_B
      Cellule_C = New_Cellule_C
      Candidats = New_Candidats
      Typologie = New_Typologie
    End Sub
  End Class

  Sub Strategy_Unq_Unq1List_Display(Unq1_list As List(Of Unq1_Cls))
    Jrn_Add(, {"Liste de Unq1_list : Nombre de Postes " & CStr(Unq1_list.Count)})
    Dim S As String
    For i As Integer = 0 To Unq1_list.Count - 1
      S = CStr(i).PadRight(5) & "  " & U_cr(Unq1_list.Item(i).Cellule) & " " &
                                       Unq1_list.Item(i).Candidats &
                                       " Occurences: " & Unq1_list.Item(i).Occurences
      Jrn_Add(, {S})
    Next i
  End Sub
  Sub Strategy_Unq_Unq2List_Display(Unq2_list As List(Of Unq2_Cls))
    Jrn_Add(, {"Liste de Unq2_list : Nombre de Postes " & CStr(Unq2_list.Count)})
    Dim S As String
    For i As Integer = 0 To Unq2_list.Count - 1
      S = CStr(i).PadRight(5) & "  " & U_cr(Unq2_list.Item(i).Cellule_A) & " " &
                                       U_cr(Unq2_list.Item(i).Cellule_B) & " " &
                                       U_cr(Unq2_list.Item(i).Cellule_C) & " " &
                                       Unq2_list.Item(i).Candidats
      Jrn_Add(, {S})
    Next i
  End Sub
  Sub Strategy_Unq_Unq3List_Display(Unq3_list As List(Of Unq3_Cls))
    Jrn_Add(, {"Liste de Unq3_list : Nombre de Postes " & CStr(Unq3_list.Count)})
    Dim S As String
    For i As Integer = 0 To Unq3_list.Count - 1
      S = CStr(i).PadRight(5) & "  " & U_cr(Unq3_list.Item(i).Cellule_A) & " " &
                                       U_cr(Unq3_list.Item(i).Cellule_B) & " " &
                                       U_cr(Unq3_list.Item(i).Cellule_C) & " " &
                                       Unq3_list.Item(i).Candidats & " " &
                                       Unq3_list.Item(i).Typologie
      Jrn_Add(, {S})
    Next i
  End Sub
End Module