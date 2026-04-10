'-------------------------------------------------------------------------------------------
' Stratégie des Cbl Candidats Bloqués
'-------------------------------------------------------------------------------------------

Friend Module P12_Strategy_Cbl

  Function Strategy_Cbl(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "Cbl"
    Dim Sous_Stratégie As String = "Cbl"

    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Cbl
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 
    For LCR As Integer = 0 To 8
      For V As Integer = 1 To 9
        Strategy_Cbl_01(U_temp, Strategy_Rslt, "R", LCR, V)       'Analyse des Cbl pour les Régions de 1 à 9
        Strategy_Cbl_02(U_temp, Strategy_Rslt, "L", LCR, V)       'Analyse des Cbl pour les Lignes  de 1 à 9
        Strategy_Cbl_02(U_temp, Strategy_Rslt, "C", LCR, V)       'Analyse des Cbl pour les Colonnes de 1 à 9
      Next V
    Next LCR
    Return Strategy_Rslt
  End Function

  Sub Strategy_Cbl_01(U_temp(,) As String,
                      ByRef Strategy_Rslt(,) As String,
                      Code_LCR As String,
                      LCR As Integer,
                      Valeur As Integer)
    Dim Stratégie As String = "Cbl"
    Dim Grp() As Integer = U_9CelReg(LCR)
    Dim CB(9) As Integer 'Ne comporte que 2 ou 3 postes dans la Stratégie des Candidats Bloqués
    Dim Nb As Integer = 0
    Dim Cel45(0 To 44) As String
    Dim Exc45(0 To 44) As String
    Dim Exc45_n As Integer       ' Nombre de Cellules Exc45 ou indice

    'La valeur est seulement contenue dans 2 ou 3 cellules d'une ligne ou d'une colonne
    For i As Integer = 0 To 8
      If U_temp(Grp(i), 3).Contains(CStr(Valeur)) = True Then
        CB(Nb) = Grp(i)
        Nb += 1
      End If
    Next i

    Dim Cl1 As Integer = U_Col(CB(0)) : Dim Rw1 As Integer = U_Row(CB(0))
    Dim Cl2 As Integer = U_Col(CB(1)) : Dim Rw2 As Integer = U_Row(CB(1))
    Dim Cl3 As Integer = U_Col(CB(2)) : Dim Rw3 As Integer = U_Row(CB(2))

    If Nb = 2 Then ' Il y a une valeur dans 2 cellules 
      If Cl1 = Cl2 Then ' Il y a une valeur dans 2 cellules de la colonne
        'Suppression dans cette colonne SAUF POUR LA REGION des candidats
        Dim Grp2() As Integer = U_9CelCol(Cl1)
        Dim n As Integer = 0
        For j As Integer = 0 To 8
          If (U_Reg(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
        Next j
        If n > 0 Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1
          Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
          Cel45(1) = CStr(CB(1))
          For k As Integer = 0 To 8
            If (U_Reg(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
            End If
          Next k
          Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "01C2a", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
        End If '/n > 0
      End If '/If Cl1 = Cl2
      If Rw1 = Rw2 Then ' Il y a une valeur dans 2 cellules de la Ligne
        'Suppression dans cette Ligne SAUF POUR LA REGION des candidats
        Dim Grp2() As Integer = U_9CelRow(Rw1)
        Dim n As Integer = 0
        For j As Integer = 0 To 8
          If (U_Reg(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
        Next j
        If n > 0 Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1
          Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
          Cel45(1) = CStr(CB(1))
          For k As Integer = 0 To 8
            If (U_Reg(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
            End If
          Next k
          Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "02L2b", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
        End If '/If n > 0
      End If '/If Rw1 = Rw2
    End If '/If Nb = 2

    If Nb = 3 Then ' Il y a une valeur dans 3 cellules 
      If (Cl1 = Cl2 And Cl2 = Cl3) Then ' Il y a une valeur dans 3 cellules de la colonne
        'Suppression dans cette colonne SAUF POUR LA REGION des candidats
        Dim Grp2() As Integer = U_9CelCol(Cl1)
        Dim n As Integer = 0
        For j As Integer = 0 To 8
          If (U_Reg(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
        Next j
        If n > 0 Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1
          Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
          Cel45(1) = CStr(CB(1))
          Cel45(2) = CStr(CB(2))
          For k As Integer = 0 To 8
            If (U_Reg(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
            End If
          Next k
          Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "01C3c", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
        End If '/n > 0
      End If '/If Cl1 = Cl2 = Cl3

      If (Rw1 = Rw2 And Rw2 = Rw3) Then ' Il y a une valeur dans 3 cellules de la ligne
        'Suppression dans cette ligne SAUF POUR LA REGION des candidats
        Dim Grp2() As Integer = U_9CelRow(Rw1)
        Dim n As Integer = 0
        For j As Integer = 0 To 8
          If (U_Reg(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
        Next j
        If n > 0 Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1
          Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
          Cel45(1) = CStr(CB(1))
          Cel45(2) = CStr(CB(2))
          For k As Integer = 0 To 8
            If (U_Reg(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
            End If
          Next k
          Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "01L3d", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
        End If '/If n > 0
      End If '/If Rw1 = Rw2 = Rw3
    End If '/If Nb = 3
  End Sub

  Sub Strategy_Cbl_02(U_temp(,) As String,
                      ByRef Strategy_Rslt(,) As String,
                      Code_LCR As String,
                      LCR As Integer,
                      Valeur As Integer)
    Dim Stratégie As String = "Cbl"
    Dim Grp() As Integer = {0}
    Select Case Code_LCR
      Case "L" : Grp = U_9CelRow(LCR)
      Case "C" : Grp = U_9CelCol(LCR)
    End Select
    Dim CB(9) As Integer 'Ne comporte que 2 ou 3 postes dans la Stratégie des Candidats Bloqués
    Dim Nb As Integer = 0
    Dim Cel45(0 To 44) As String
    Dim Exc45(0 To 44) As String
    Dim Exc45_n As Integer   ' Nombre de Cellules Exc45 ou indice
    'La valeur est seulement contenue dans 2 ou 3 cellules d'une ligne ou d'une colonne

    For i As Integer = 0 To 8
      If U_temp(Grp(i), 3).Contains(CStr(Valeur)) = True Then
        CB(Nb) = Grp(i)
        Nb += 1
      End If
    Next i

    Dim Rg1 As Integer = U_Reg(CB(0))
    Dim Rg2 As Integer = U_Reg(CB(1))
    Dim Rg3 As Integer = U_Reg(CB(2))

    If Nb = 2 Then ' Il y a une valeur dans 2 cellules 
      If Rg1 = Rg2 Then
        'La valeur est dans la même région
        'la valeur peut être exclue des autres cellules, HORS Ligne ou HORS Colonne
        Dim Grp2() As Integer = U_9CelReg(Rg1)
        Select Case Code_LCR
          Case "L"
            Dim n As Integer = 0
            For j As Integer = 0 To 8
              If (U_Row(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
            Next j
            If n > 0 Then
              For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
              For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
              Exc45_n = -1
              Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
              Cel45(1) = CStr(CB(1))
              For k As Integer = 0 To 8
                If (U_Row(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
                  Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
                End If
              Next k
              Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "02L2e", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
            End If '/If n > 0

          Case "C"
            Dim n As Integer = 0
            For j As Integer = 0 To 8
              If (U_Col(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
            Next j
            If n > 0 Then
              For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
              For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
              Exc45_n = -1
              Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
              Cel45(1) = CStr(CB(1))
              For k As Integer = 0 To 8
                If (U_Col(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
                  Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
                End If
              Next k
              Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "02C2f", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
            End If '/If n > 0
        End Select

      End If '/If Rg1 = Rg2
    End If '/If Nb = 2

    If Nb = 3 Then ' Il y a une valeur dans 3 cellules 
      If (Rg1 = Rg2 And Rg2 = Rg3) Then
        'La valeur est dans la même région
        'la valeur peut être exclue des autres cellules, HORS Ligne ou HORS Colonne
        Dim Grp2() As Integer = U_9CelReg(Rg1)
        Select Case Code_LCR
          Case "L"
            Dim n As Integer = 0
            For j As Integer = 0 To 8
              If (U_Row(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
            Next j
            If n > 0 Then
              For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
              For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
              Exc45_n = -1
              Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
              Cel45(1) = CStr(CB(1))
              Cel45(2) = CStr(CB(2))
              For k As Integer = 0 To 8
                If (U_Row(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
                  Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
                End If
              Next k
              Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "02L3g", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
            End If '/If n > 0

          Case "C"
            Dim n As Integer = 0
            For j As Integer = 0 To 8
              If (U_Col(Grp2(j)) <> LCR And U_temp(Grp2(j), 3).Contains(CStr(Valeur))) = True Then n += 1
            Next j
            If n > 0 Then
              For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
              For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
              Exc45_n = -1
              Cel45(0) = CStr(CB(0))                       'Cellules Concernées par la stratégie
              Cel45(1) = CStr(CB(1))
              Cel45(2) = CStr(CB(2))
              For k As Integer = 0 To 8
                If (U_Col(Grp2(k)) <> LCR And U_temp(Grp2(k), 3).Contains(CStr(Valeur))) = True Then
                  Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp2(k))                 'Cellules Concernées par l’exclusion
                End If
              Next k
              Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Code_LCR & "02C3h", Code_LCR, CStr(LCR), CStr(Valeur), Cel45, Exc45)
            End If '/If n > 0
        End Select
      End If '/If Rg1 = Rg2 = Rg3
    End If '/If Nb = 3
  End Sub
End Module