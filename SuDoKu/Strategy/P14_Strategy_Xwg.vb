Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------
'
' Stratégie des XWing
' - dans 2 colonnes, en supprimant les candidats dans 2 lignes
' - dans 2 lignes,   en supprimant les candidats dans 2 colonnes
'
' La stratégie XWing  est notée Xwg 
' La stratégie XYWing est notée XYw
'-------------------------------------------------------------------------------
'La stratégie XWg utilise Strategy_Rslt qui stocke les candidats et les cellules des candidats bloqués
'
'
'  XU est utilisé pour placer en colonne 10 le candidat commun sur la ligne
'                          et en ligne   10 le candidat commun sur la colonne
'  XU représente donc la grille d'un candidat en 9*9 avec les résultats en ligne et colonne 10
'  Il ne doit y avoir AUCUNE variables PUBLIQUES dans le module 
'  XU et XU_New est remplacé par U_ps(10,10) as string
'  Code_LCR Contient L au lieu de Ligne
'                    C            Colonne
'  Mid$(Code_LCR, 1, 1) devient Code_LCR
'  Le traitement des Régions est supprimé

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Friend Module P14_Strategy_Xwg
  ' 06/02/2025  Stratégie Finned Ailettes Finned-X_Wing
  ' 11/03/2025  Stratégie Finned Sashimi X_Wing

  Public Function Strategy_Xwg(U_temp(,) As String) As String(,)
    ' Rappel Stratégie / Sous_Stratégie et racine de traitement
    '        Xwg         Xwg
    '        Xwg         Fnd
    '        Xwg         Shm
    '

    Dim Stratégie As String = "Xwg"
    Dim Sous_Stratégie As String = "Xwg"
    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Cbl
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    'Stratégie Xwg Xwg (_C, _L)
    For candidat As Integer = 1 To 9
      Strategy_Xwg(U_temp, Strategy_Rslt, "L", CStr(candidat))
      Strategy_Xwg(U_temp, Strategy_Rslt, "C", CStr(candidat))
    Next candidat

    'Stratégie Xwg Fnd
    For candidat As Integer = 1 To 9
      Strategy_Fnd(U_temp, Strategy_Rslt, CStr(candidat))
    Next candidat

    'Stratégie Xwg Sashimi
    For candidat As Integer = 1 To 9
      Strategy_Shm(U_temp, Strategy_Rslt, CStr(candidat))
    Next candidat

    Return Strategy_Rslt
  End Function

  Public Sub Strategy_Fnd(U_temp(,) As String,
                          ByRef Strategy_Rslt(,) As String,
                          Candidat As String)
    'U_ps comporte le candidat étudié et en row 9 et lig 9 le nombre de candidats par colonne et par ligne
    Dim U_ps(0 To 9, 0 To 9) As String 'Ce tableau est ensuite transmis  
    U_ps_Calcul(U_temp, Candidat, U_ps)
    'Déterminer une position de XWing 2 et 3 
    'Jrn_Add("SDK_Space")
    Dim CelA1, CelA2 As Integer        ' 2 cellules pour la base  
    Dim CelB1, CelB2, CelB3 As Integer ' 3 pour l'autre côté      
    ' Traitement des lignes et des colonnes
    Strategy_Fnd_LC(U_ps, U_temp, Candidat, Strategy_Rslt, CelA1, CelA2, CelB1, CelB2, CelB3, True)  ' True pour les lignes
    Strategy_Fnd_LC(U_ps, U_temp, Candidat, Strategy_Rslt, CelA1, CelA2, CelB1, CelB2, CelB3, False) ' False pour les colonnes
  End Sub

  Private Sub Strategy_Fnd_LC(ByRef U_ps(,) As String, U_temp(,) As String, Candidat As String, ByRef Strategy_Rslt(,) As String, ByRef CelA1 As Integer, ByRef CelA2 As Integer, ByRef CelB1 As Integer, ByRef CelB2 As Integer, ByRef CelB3 As Integer, isRow As Boolean)
    Dim Stratégie As String = "Xwg"
    'La sous-stratégie sera complétée ensuite.
    Dim Sous_Stratégie As String = If(isRow, "Fnd_L", "Fnd_C")
    Dim Code_LCR As String = If(isRow, "C", "L")
    Dim LCR As Integer
    Dim Ai, Le, Ro As Integer
    For i As Integer = 0 To 8
      If U_ps(If(isRow, i, 9), If(isRow, 9, i)) = "2" Then
        CelA1 = -1 : CelA2 = -1
        For j As Integer = 0 To 8
          Dim Cel As Integer = Wh_Cellule_RowCol(If(isRow, i, j), If(isRow, j, i))
          If U_ps(If(isRow, i, j), If(isRow, j, i)) = Candidat Then
            If CelA1 = -1 Then
              CelA1 = Cel
            ElseIf CelA2 = -1 Then
              CelA2 = Cel
            End If
          End If
        Next j
        For k As Integer = 0 To 8
          If k = i Then Continue For
          If U_ps(If(isRow, k, 9), If(isRow, 9, k)) = "3" Then
            CelB1 = -1 : CelB2 = -1 : CelB3 = -1
            For l As Integer = 0 To 8
              Dim Cel As Integer = Wh_Cellule_RowCol(If(isRow, k, l), If(isRow, l, k))
              If U_ps(If(isRow, k, l), If(isRow, l, k)) = Candidat Then
                If CelB1 = -1 Then
                  CelB1 = Cel
                ElseIf CelB2 = -1 Then
                  CelB2 = Cel
                ElseIf CelB3 = -1 Then
                  CelB3 = Cel
                End If
              End If
            Next l
            Dim Cel45(0 To 44) As String : For k1 As Integer = 0 To 44 : Cel45(k1) = "__" : Next k1
            Dim Exc45(0 To 44) As String : For k1 As Integer = 0 To 44 : Exc45(k1) = "__" : Next k1
            Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice
            '
            '                                                               B1 x
            '                                                                  |
            '               A1              A2                     A1 x     B2 x   B1  x   B1 x
            '               x---------------x                         |        |       |      | 
            '                                                         |        |       |      |
            '          B1   B2              B3                        |        |       |      |
            '           x---x---------------x                         |        |       |      |
            '               B1          B2  B3                        |        |       |      |
            '               x-----------x---x                         |        |   B2  x      |  
            '               B1              B2  B3                    |        |       |      |
            '               x---------------x---x                  A2 x     B3 x   B3  x   B2 x
            '                                                                                 | 
            '                                                                              B3 x
            '
            LCR = -1
            If isRow Then  'Si le traitement concerne les rangées, je dois tester les colonnes
              If Is_SameCol(CelA1, CelB2) And Is_SameCol(CelA2, CelB3) And Is_SameReg(CelB1, CelB2) Then
                '.5267.3.8.3...562767..325.128...61.5.6....2.4714523869827314956.9.267483346958712
                LCR = U_Col(CelA1) : Ai = CelA1 : Le = CelB2 : Ro = CelB1 : Sous_Stratégie = "Fnd_La"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

              If Is_SameCol(CelA1, CelB1) And Is_SameCol(CelA2, CelB3) And Is_SameReg(CelB2, CelB3) Then
                '.83.4926..7...2...2....35.......6....57.316.26..275...76532814939.1.4.....1..7...
                LCR = U_Col(CelA2) : Ai = CelA2 : Le = CelB3 : Ro = CelB2 : Sous_Stratégie = "Fnd_LbR"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

              If Is_SameCol(CelA1, CelB1) And Is_SameCol(CelA2, CelB3) And Is_SameReg(CelB2, CelB1) Then
                '.6294.38....2...7...53....2...6.....2.613.75....572..6941823567...4.1.93...7..1..
                LCR = U_Col(CelA1) : Ai = CelA1 : Le = CelB1 : Ro = CelB2 : Sous_Stratégie = "Fnd_LbL"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

              If Is_SameCol(CelA1, CelB1) And Is_SameCol(CelA2, CelB2) And Is_SameReg(CelB3, CelB2) Then
                '217859643384762.9.6594137289683254174.2....6.5.16...821.523..767265...3.8.3.7625.
                LCR = U_Col(CelA2) : Ai = CelA2 : Le = CelB2 : Ro = CelB3 : Sous_Stratégie = "Fnd_Lc"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

            Else           'Si le traitement concerne les colonnes, je dois tester les rangées

              If Is_SameRow(CelA1, CelB2) And Is_SameRow(CelA2, CelB3) And Is_SameReg(CelB1, CelB2) Then
                '3.87.26..4921687356.74....29235....65612..3.78743.625.7498215631856...2.236945178
                LCR = U_Row(CelA1) : Ai = CelA1 : Le = CelB2 : Ro = CelB1 : Sous_Stratégie = "Fnd_Ca"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

              If Is_SameRow(CelA1, CelB1) And Is_SameRow(CelA2, CelB3) And Is_SameReg(CelB1, CelB2) Then
                '..9.2......4.....6..1.6.5.2748516329..273...4.132.....1.5.7...3.96.5..78.376..2..
                LCR = U_Row(CelA1) : Ai = CelA1 : Le = CelB1 : Ro = CelB2 : Sous_Stratégie = "Fnd_CbT"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

              If Is_SameRow(CelA1, CelB1) And Is_SameRow(CelA2, CelB3) And Is_SameReg(CelB2, CelB3) Then
                '.376..2...96.5..781.5.7...3.132.......273...4748516329..1.6.5.2..4.....6..9.2....
                LCR = U_Row(CelA2) : Ai = CelA2 : Le = CelB3 : Ro = CelB2 : Sous_Stratégie = "Fnd_CbB"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If

              If Is_SameRow(CelA1, CelB1) And Is_SameRow(CelA2, CelB2) And Is_SameReg(CelB3, CelB2) Then
                '217859643384762.9.6594137289683254174.2....6.5.16...821.523..767265...3.8.3.7625.
                LCR = U_Row(CelA2) : Ai = CelA2 : Le = CelB2 : Ro = CelB3 : Sous_Stratégie = "Fnd_Cc"
                Strategy_Fnd_Add(Sous_Stratégie, LCR, U_temp, Candidat, Ai, Le, Ro, Exc45, Exc45_n)
              End If
            End If
            If Ai <> -1 And Ro <> -1 And Exc45_n <> -1 Then
              Cel45(0) = CStr(CelA1)      ' Petite Branche
              Cel45(1) = CStr(CelA2)
              Cel45(2) = CStr(CelB1)      ' Grande branche
              Cel45(3) = CStr(CelB2)
              Cel45(4) = CStr(CelB3)
              Cel45(5) = CStr(Ai)         ' Aileron
              Cel45(6) = CStr(Ro)
              Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, CStr(LCR), Candidat, Cel45, Exc45)
            End If
          End If
        Next k
      End If
    Next i
  End Sub

  Public Sub Strategy_Fnd_Add(Sous_stratégie As String, LCR As Integer, U_temp(,) As String, Candidat As String, Ai As Integer, Le As Integer, Ro As Integer, ByRef Exc45() As String, ByRef Exc45_n As Integer)
    Dim Grp As Integer() = New Integer(8) {}
    Select Case Mid(Sous_stratégie, 5, 1)
      Case "L" : Grp = U_9CelCol(LCR)
      Case "C" : Grp = U_9CelRow(LCR)
    End Select
    For c As Integer = 0 To 8
      Dim Cellule As Integer = Grp(c)
      If Cellule = Ai Or Cellule = Le Then Continue For
      If Not Is_SameReg(Cellule, Ro) Then Continue For
      If U_temp(Cellule, 3).Contains(Candidat) Then
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next c
  End Sub


  Public Sub Strategy_Shm(U_temp(,) As String,
                          ByRef Strategy_Rslt(,) As String,
                          Candidat As String)
    'U_ps comporte le candidat étudié et en row 9 et lig 9 le nombre de candidats par colonne et par ligne
    Dim U_ps(0 To 9, 0 To 9) As String 'Ce tableau est ensuite transmis  
    U_ps_Calcul(U_temp, Candidat, U_ps)
    Strategy_Shm_LC(U_ps, U_temp, Candidat, Strategy_Rslt, False)  ' False pour les colonnes
    Strategy_Shm_LC(U_ps, U_temp, Candidat, Strategy_Rslt, True)   ' True pour les lignes
  End Sub

  Private Sub Strategy_Shm_LC(ByRef U_ps(,) As String, U_temp(,) As String, Candidat As String, ByRef Strategy_Rslt(,) As String,
                              isRow As Boolean)
    Dim Stratégie As String = "Xwg", Sous_Stratégie As String = "Shm_#"
    Dim Code_LCR As String = "#", LCR As Integer = -1
    ' Déclaration explicite des tableaux
    Dim A(1) As Integer ' Ligne ou Colonne à 2 candidats
    Dim B(2) As Integer ' Ligne ou Colonne à 3 candidats
    Dim A0, A1 As Integer
    Dim B0, B1, B2 As Integer

    'Pour un même candidat,
    'la procédure passe 2 fois, la première fois pour traiter les rangées
    '                           la seconde  fois pour traiter les colonnes
    For i As Integer = 0 To 8
      If U_ps(If(isRow, i, 9), If(isRow, 9, i)) = "2" Then
        A(0) = -1 : A(1) = -1 ' Initialisation à -1, car 0 représente L1-C1
        For j As Integer = 0 To 8
          Dim Cel As Integer = Wh_Cellule_RowCol(If(isRow, i, j), If(isRow, j, i))
          If U_ps(If(isRow, i, j), If(isRow, j, i)) = Candidat Then
            If A(0) = -1 Then
              A(0) = Cel
            ElseIf A(1) = -1 Then
              A(1) = Cel
            End If
          End If
        Next j
        For k As Integer = 0 To 8
          If k = i Then Continue For
          If U_ps(If(isRow, k, 9), If(isRow, 9, k)) = "3" Then
            B(0) = -1 : B(1) = -1 : B(2) = -1
            For l As Integer = 0 To 8
              Dim Cel As Integer = Wh_Cellule_RowCol(If(isRow, k, l), If(isRow, l, k))
              If U_ps(If(isRow, k, l), If(isRow, l, k)) = Candidat Then
                If B(0) = -1 Then
                  B(0) = Cel
                ElseIf B(1) = -1 Then
                  B(1) = Cel
                ElseIf B(2) = -1 Then
                  B(2) = Cel
                End If
              End If
            Next l

            A0 = -1 : A1 = -1 : B0 = -1 : B1 = -1 : B2 = -1
            For ia As Integer = 0 To 1   ' A(0), A(1)
              For ib As Integer = 0 To 2 ' B(0), B(1), B(2)
                If isRow Then ' Test des Rangées
                  If Is_SameCol(A(ia), B(ib)) And Not Is_SameCol(A(1 - ia), B((ib + 1) Mod 3)) And Not Is_SameCol(A(1 - ia), B((ib + 2) Mod 3)) Then
                    A0 = A(ia) : B0 = B(ib)
                    A1 = A(1 - ia) : B1 = B((ib + 1) Mod 3) : B2 = B((ib + 2) Mod 3)
                    Exit For
                  End If
                Else          ' Test des Colonnes
                  If Is_SameRow(A(ia), B(ib)) And Not Is_SameRow(A(1 - ia), B((ib + 1) Mod 3)) And Not Is_SameRow(A(1 - ia), B((ib + 2) Mod 3)) Then
                    A0 = A(ia) : B0 = B(ib)
                    A1 = A(1 - ia) : B1 = B((ib + 1) Mod 3) : B2 = B((ib + 2) Mod 3)
                    Exit For
                  End If
                End If '/isRow
              Next ib
              If A0 <> -1 Then Exit For ' Sortie anticipée dès qu'une paire est trouvée
            Next ia

            Dim Cel45(0 To 44) As String : For k1 As Integer = 0 To 44 : Cel45(k1) = "__" : Next k1
            Dim Exc45(0 To 44) As String : For k1 As Integer = 0 To 44 : Exc45(k1) = "__" : Next k1
            Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice

            If A0 <> -1 AndAlso B0 <> -1 AndAlso Is_SameReg(B1, B2) Then
              LCR = U_Reg(B1)
              Dim Grp As Integer() = U_9CelReg(LCR)
              For c As Integer = 0 To 8
                Dim Cellule As Integer = Grp(c)
                If Cellule = A0 Or Cellule = A1 Then Continue For
                If Cellule = B0 Or Cellule = B1 Or Cellule = B2 Then Continue For
                If isRow Then
                  If Not Is_SameCol(Cellule, A1) Then Continue For
                Else
                  If Not Is_SameRow(Cellule, A1) Then Continue For
                End If
                If U_temp(Cellule, 3).Contains(Candidat) Then
                  Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
                End If
              Next c

              If B1 <> -1 AndAlso B2 <> -1 AndAlso Exc45_n <> -1 Then
                Sous_Stratégie = If(isRow, "Shm_L", "Shm_C")
                Code_LCR = If(isRow, "L", "C")
                Cel45(0) = CStr(A0)           ' Petite Branche
                Cel45(1) = CStr(A1)
                Cel45(2) = CStr(B0)           ' Grande branche
                Cel45(3) = CStr(B1)
                Cel45(4) = CStr(B2)
                Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, CStr(LCR), Candidat, Cel45, Exc45)
              End If '/B1 <> -1 AndAlso B2 <> -1 AndAlso Exc45_n <> -1 

            End If '/A0 <> -1 AndAlso B0 <> -1 AndAlso Is_SameReg(B1, B2)
          End If '/U_ps(If(isRow, k, 9), If(isRow, 9, k)) = "3"
        Next k
      End If '/U_ps(If(isRow, i, 9), If(isRow, 9, i)) = "2" Then
    Next i
  End Sub

  Private Sub U_ps_Calcul(U_temp(,) As String, Candidat As String, ByRef U_ps(,) As String)
    Dim Grp As Integer()
    Dim n As Integer
    ' 01 - Il faut déterminer les Lignes/Colonnes où le candidat étudié se trouve 
    For LCR As Integer = 0 To 8
      Grp = U_9CelRow(LCR)
      For i As Integer = 0 To 8
        If U_temp(Grp(i), 3).Contains(Candidat) = True Then U_ps(U_Row(Grp(i)), U_Col(Grp(i))) = Candidat
      Next i
    Next LCR
    ' 02 - Total lignes où le candidat étudié se trouve 
    For row As Integer = 0 To 8
      n = 0
      For col As Integer = 0 To 8
        If U_ps(row, col) <> "" Then n += 1
      Next col
      U_ps(row, 9) = CStr(n)
    Next row
    ' 03 - Total Colonnes où le candidat étudié se trouve 
    For col As Integer = 0 To 8
      n = 0
      For row As Integer = 0 To 8
        If U_ps(row, col) <> "" Then n += 1
      Next row
      U_ps(9, col) = CStr(n)
    Next col
    ' 04 - Total des lignes  
    n = 0
    For row As Integer = 0 To 8
      If U_ps(row, 9) <> "0" Then n += CInt(U_ps(row, 9))
    Next row
    U_ps(9, 9) = CStr(n)
  End Sub

  Sub U_ps_Display(Candidat As String, U_ps(,) As String)
    Jrn_Add("SDK_Space")
    Jrn_Add(, {"Candidat : " & Candidat})
    'Le décalage en 9-9 est correct.
    Dim Sb As New System.Text.StringBuilder()
    For row As Integer = 0 To 9
      Sb.Clear()
      For col As Integer = 0 To 9
        Sb.Append(If(U_ps(row, col) = "", " _", " " & U_ps(row, col)))
      Next col
      Dim S As String = Sb.ToString()
      Jrn_Add(, {Sb.ToString()})
    Next row
  End Sub


  Public Sub Strategy_Xwg(U_temp(,) As String,
                          ByRef Strategy_Rslt(,) As String,
                          Code_LCR As String,
                          Candidat As String)

    Dim Grp(0 To 8) As Integer
    Dim n1, n2, n3 As Integer
    Dim U_px(9, 9) As String 'Ce tableau est ensuite transmis
    For row As Integer = 0 To 9
      For col As Integer = 0 To 9
        U_px(row, col) = "__"
      Next col
    Next row

    ' 01 - Il faut déterminer les lignes où le candidat étudié se trouve 
    n2 = 0 'n2 = Nombre d'unités_LCR dans lesquelles on trouve 2 fois un candidat
    For LCR As Integer = 0 To 8
      Select Case Code_LCR
        Case "L" : Grp = U_9CelRow(LCR)
        Case "C" : Grp = U_9CelCol(LCR)
      End Select
      n1 = 0  'n1 = Nombre de fois où un candidat se trouve dans une Unité_LCR
      Dim Cel As Integer = -1
      For i As Integer = 0 To 8
        Cel = Grp(i)
        If U_temp(Cel, 3).Contains(Candidat) = True Then
          n1 += 1
          U_px(U_Row(Cel), U_Col(Cel)) = CStr(Cel)
        End If
      Next i
      ' 02 - Dans ce cas où le candidat ne se trouve que 2 fois le numéro de la cellule est conservé
      For i As Integer = 0 To 8
        Cel = Grp(i)
        If n1 <> 2 Then U_px(U_Row(Cel), U_Col(Cel)) = "__"
      Next i
      'On ajoute en colonne 9 le nombre d'Unité si 2
      Select Case Code_LCR
        Case "L"
          If n1 <> 2 Then U_px(U_Row(Cel), 9) = " " & CStr(n1)
          If n1 = 2 Then U_px(U_Row(Cel), 9) = Code_LCR & CStr(n1)
        Case "C"
          If n1 <> 2 Then U_px(U_Col(Cel), 9) = " " & CStr(n1)
          If n1 = 2 Then U_px(U_Col(Cel), 9) = Code_LCR & CStr(n1)
      End Select
      If n1 = 2 Then
        U_px(U_Row(Cel), 9) = Code_LCR & CStr(n1)
        n2 += 1
      End If
    Next LCR
    'On ajoute en cellule 9,9 le nombre d'Unité si 2 et plus
    U_px(9, 9) = Code_LCR & CStr(n2)

    If n2 < 2 Then Exit Sub
    'Une unité doit comporter au moins 2 unités LCR de 2 candidats pour être intéressante
    'Pour lancer un traitement d'unités, il ne faut que 2 unités.
    'Donc si 2, 1 lancement  : a,b     ab
    '     si 3, 3 lancements : a,b,c   ab, ac, bc
    '     si 4, 6 lancements : a,b,c,d ab, ac, ad, bc, bd, cd
    '     etc
    n3 = 0 'n3 = Nombre de comparaisons des unités_LCR dans lesquelles on trouve 2 fois un candidat
    For a1 As Integer = 0 To 8
      For b1 As Integer = a1 To 8
        If a1 <> b1 Then
          'On compare U_ps(a1, 9) et U_ps(b1, 9)
          'Il faut qu'il y ait L,C en position 1
          If (Mid$(U_px(a1, 9), 1, 1) <> " " And Mid$(U_px(b1, 9), 1, 1) <> " ") Then
            n3 += 1
            Select Case Code_LCR
              Case "L" : Strategy_Xwg_L(U_temp, Strategy_Rslt, U_px, Code_LCR, Candidat, a1, b1)
              Case "C" : Strategy_Xwg_C(U_temp, Strategy_Rslt, U_px, Code_LCR, Candidat, a1, b1)
            End Select
          End If
        End If
      Next b1
    Next a1
    Exit Sub
  End Sub

  Public Sub Strategy_Xwg_L(U_temp(,) As String,
                            ByRef Strategy_Rslt(,) As String,
                            ByRef U_px(,) As String,
                            Code_LCR As String,
                            Candidat As String,
                            LCR1 As Integer,
                            LCR2 As Integer)
    ' Cette méthode s'applique dans les cas suivants, avec 2 candidats:
    '- dans 2 lignes, en supprimant les candidats dans 2 colonnes
    '- dans 2 lignes, en supprimant les candidats dans 2 régions
    Dim Stratégie As String = "Xwg"
    Dim Sous_Stratégie As String = "Xwg_C"
    Dim Pos1g As String = "__"  'Correspond à la première cellule de la première ligne
    '        Dim Pos1g_Row As Integer
    Dim Pos1g_Col As Integer
    Dim Pos1d As String = "__"  'Correspond à la deuxième cellule de la première ligne
    '        Dim Pos1d_Row As Integer
    Dim Pos1d_Col As Integer
    Dim Pos2g As String = "__"  'Correspond à la première cellule de la deuxième ligne
    '        Dim Pos2g_Row As Integer
    Dim Pos2g_Col As Integer
    Dim Pos2d As String = "__"  'Correspond à la deuxième cellule de la deuxième ligne
    '        Dim Pos2d_Row As Integer
    Dim Pos2d_Col As Integer
    Dim n As Integer = 0
    For col As Integer = 0 To 8
      If U_px(LCR1, col) <> "__" Then
        n += 1
        If n = 1 Then Pos1g = U_px(LCR1, col)
        If n = 2 Then Pos1d = U_px(LCR1, col)
      End If
    Next col
    n = 0
    For col As Integer = 0 To 8
      If U_px(LCR2, col) <> "__" Then
        n += 1
        If n = 1 Then Pos2g = U_px(LCR2, col)
        If n = 2 Then Pos2d = U_px(LCR2, col)
      End If
    Next col

    '01 Est-ce que ces 2 lignes ont dans les mêmes colonnes ?
    Pos1g_Col = U_Col(CInt(Pos1g))
    Pos1d_Col = U_Col(CInt(Pos1d))
    Pos2g_Col = U_Col(CInt(Pos2g))
    Pos2d_Col = U_Col(CInt(Pos2d))
    If (Pos1g_Col = Pos2g_Col) And (Pos1d_Col = Pos2d_Col) Then
      Dim Cellule As Integer
      Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
      Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
      Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice

      'On enlève dans les colonnes
      Dim Col_g As Integer = Pos1g_Col  'Colonne de Gauche
      Dim Col_d As Integer = Pos1d_Col  'Colonne de Droite
      ' 1 Traitement de la Colonne de gauche sauf Pos1g et Pos2g
      Dim Grp As Integer() = U_9CelCol(Col_g)
      For i As Integer = 0 To 8
        Cellule = Grp(i)
        If (Cellule <> CInt(Pos1g) And Cellule <> CInt(Pos2g) And U_temp(Cellule, 2) = " " And U_temp(Cellule, 3).Contains(Candidat)) Then
          Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
        End If
      Next i
      ' 2 Traitement de la Colonne de droite sauf Pos1d et Pos2d
      Grp = U_9CelCol(Col_d)
      For i As Integer = 0 To 8
        Cellule = Grp(i)
        If (Cellule <> CInt(Pos1d) And Cellule <> CInt(Pos2d) And U_temp(Cellule, 2) = " " And U_temp(Cellule, 3).Contains(Candidat)) Then
          Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
        End If
      Next i
      Cel45(0) = Pos1g
      Cel45(1) = Pos1d
      Cel45(2) = Pos2g
      Cel45(3) = Pos2d
      If Exc45_n > -1 Then Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "-1", Candidat, Cel45, Exc45)
    End If
  End Sub

  Public Sub Strategy_Xwg_C(U_temp(,) As String,
                            ByRef Strategy_Rslt(,) As String,
                            ByRef U_px(,) As String,
                            Code_LCR As String,
                            Candidat As String,
                            LCR1 As Integer,
                            LCR2 As Integer)

    ' Cette méthode s'applique dans les cas suivants, avec 2 candidats:
    '- dans 2 colonnes, en supprimant les candidats dans 2 lignes
    '- dans 2 colonnes, en supprimant les candidats dans 2 régions
    Dim Stratégie As String = "Xwg"
    Dim Sous_Stratégie As String = "Xwg_L"
    Dim Pos1g_Row As Integer
    Dim Pos1d_Row As Integer
    Dim Pos2g_Row As Integer
    Dim Pos2d_Row As Integer
    Dim Pos1g As String = "__"
    Dim Pos1d As String = "__"
    Dim Pos2g As String = "__"
    Dim Pos2d As String = "__"
    Dim n As Integer = 0
    For row As Integer = 0 To 8
      If U_px(row, LCR1) <> "__" Then
        n += 1
        If n = 1 Then Pos1g = U_px(row, LCR1)
        If n = 2 Then Pos1d = U_px(row, LCR1)
      End If
    Next row
    n = 0
    For row As Integer = 0 To 8
      If U_px(row, LCR2) <> "__" Then
        n += 1
        If n = 1 Then Pos2g = U_px(row, LCR2)
        If n = 2 Then Pos2d = U_px(row, LCR2)
      End If
    Next row

    '01 Est-ce que ces 2 Colonnes sont dans les mêmes lignes ?
    Pos1g_Row = U_Row(CInt(Pos1g))
    Pos1d_Row = U_Row(CInt(Pos1d))
    Pos2g_Row = U_Row(CInt(Pos2g))
    Pos2d_Row = U_Row(CInt(Pos2d))
    If (Pos1g_Row = Pos2g_Row) And (Pos1d_Row = Pos2d_Row) Then

      Dim Cellule As Integer
      Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
      Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
      Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice

      'On enlève dans les Lignes
      Dim Lig_h As Integer = Pos1g_Row  ' Ligne haute
      Dim Lig_b As Integer = Pos1d_Row  ' Ligne basse
      ' 1 Traitement de la ligne du haut sauf Pos1g et Pos2g
      'Grp = U_9CelRow(Lig_h)
      Dim Grp As Integer() = U_9CelRow(Lig_h)
      For i As Integer = 0 To 8
        Cellule = Grp(i)
        If (Cellule <> CInt(Pos1g) And Cellule <> CInt(Pos2g) And U_temp(Cellule, 2) = " " And U_temp(Cellule, 3).Contains(Candidat)) Then
          Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
        End If
      Next i
      ' 2 Traitement de la ligne de bas sauf Pos1d et Pos2d
      Grp = U_9CelRow(Lig_b)
      For i As Integer = 0 To 8
        Cellule = Grp(i)
        If (Cellule <> CInt(Pos1d) And Cellule <> CInt(Pos2d) And U_temp(Cellule, 2) = " " And U_temp(Cellule, 3).Contains(Candidat)) Then
          Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
        End If
      Next i
      Cel45(0) = Pos1g
      Cel45(1) = Pos1d
      Cel45(2) = Pos2g
      Cel45(3) = Pos2d
      'Code_LCR, CStr(LCR) Code_LCR indique L ou C, mais il y en a 2, donc LCR = "-1"
      If Exc45_n > -1 Then Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "-1", Candidat, Cel45, Exc45)
    End If
  End Sub

End Module