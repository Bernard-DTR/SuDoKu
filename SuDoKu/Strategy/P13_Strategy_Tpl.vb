'-------------------------------------------------------------------------------
' Stratégie des Tuples: Paires, Triples et Quadruples
'-------------------------------------------------------------------------------
Friend Module P13_Strategy_Tpl
  Function Strategy_Tpl(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "Tpl"
    Dim Sous_Stratégie As String = "Tpl"

    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Cbl
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    Dim T_Code_LCR As String() = {"L", "C", "R"}
    For i1 As Integer = 0 To 2 'L, C, R  _  Lancement : 3LCR
      Dim Code_LCR As String = T_Code_LCR(i1)
      For LCR As Integer = 0 To 8   ' _  Lancement : 3LCR*9
        For Tp_Nb_Cdd As Integer = 2 To 4 'Test des Doubles, Triples et Quadruples pour un traitement Ligne, Colonne et Région  _  Lancement : 3LCR*9*3=54*2(caché/évident)=108
          Dim Grp(0 To 8) As Integer
          Select Case Code_LCR
            Case "L" : Grp = U_9CelRow(LCR) 'Comporte les 9 cellules de la ligne
            Case "C" : Grp = U_9CelCol(LCR) '                              colonne
            Case "R" : Grp = U_9CelReg(LCR) '                              région
          End Select
          Strategy_Tpl_Caché_LCR(U_temp, Strategy_Rslt, Tp_Nb_Cdd, Code_LCR, LCR, Grp)
          Strategy_Tpl_Evident_LCR(U_temp, Strategy_Rslt, Tp_Nb_Cdd, Code_LCR, LCR, Grp)
        Next Tp_Nb_Cdd
      Next LCR
    Next i1
    Return Strategy_Rslt
  End Function

  Sub Strategy_Tpl_Evident_LCR(U_temp(,) As String,
                               ByRef Strategy_Rslt(,) As String,
                               Tp As Integer,
                               Code_LCR As String,
                               LCR As Integer,
                               Grp() As Integer)
    Dim UTpE(9) As Integer
    Dim Tp_Nb As Integer = 0
    Dim GrpUtpE(9) As String
    Dim n As Integer
    For i As Integer = 0 To 8 : UTpE(i) = 0 : Next i
    '1 Le nombre de candidats est 2       pour Tp = 2 
    '                             3, 2    pour Tp = 3
    '                             4, 3, 2 pour Tp = 4

    For i As Integer = 0 To 8      ' Analyse de la ligne, de la colonne ou de la région
      If (Wh_Cell_Nb_Candidats(U_temp, Grp(i)) > 1 And Wh_Cell_Nb_Candidats(U_temp, Grp(i)) <= Tp) Then
        UTpE(Tp_Nb) = Grp(i)
        Tp_Nb += 1
      End If
    Next i

    If Tp_Nb >= Tp Then
      '2 Analyse du groupe
      '~~~~~~~~~~~~~~~~~~~~
      Select Case Tp
        Case 2 'Paires
          n = 0
          For i1 As Integer = 0 To Tp_Nb - 1
            For i2 As Integer = i1 + 1 To Tp_Nb - 1
              If Wh_Cell_Nb_Candidats(U_temp, UTpE(i1)) = Tp Or Wh_Cell_Nb_Candidats(U_temp, UTpE(i2)) = Tp Then
                For i As Integer = 0 To 8 : GrpUtpE(i) = "__" : Next i
                For i As Integer = 0 To 8 : If U_temp(Grp(i), 2) <> " " Then : GrpUtpE(i) = "VV" : End If : Next i
                For i As Integer = 0 To 8
                  If UTpE(i1) = Grp(i) Or UTpE(i2) = Grp(i) Then GrpUtpE(i) = "xx"
                Next i
                n += 1 : Strategy_Tpl_Evident_Analyse(U_temp, Strategy_Rslt, Tp, Code_LCR, LCR, Grp, GrpUtpE, n)
              End If
            Next i2
          Next i1
        Case 3 'Triples
          n = 0
          For i1 As Integer = 0 To Tp_Nb - 1
            For i2 As Integer = i1 + 1 To Tp_Nb - 1
              For i3 As Integer = i2 + 1 To Tp_Nb - 1
                'Au moins 1 à 3
                If Wh_Cell_Nb_Candidats(U_temp, UTpE(i1)) = Tp Or Wh_Cell_Nb_Candidats(U_temp, UTpE(i2)) = Tp Or Wh_Cell_Nb_Candidats(U_temp, UTpE(i2)) = Tp Then
                  For i As Integer = 0 To 8 : GrpUtpE(i) = "__" : Next i
                  For i As Integer = 0 To 8 : If U_temp(Grp(i), 2) <> " " Then : GrpUtpE(i) = "VV" : End If : Next i
                  For i As Integer = 0 To 8
                    If UTpE(i1) = Grp(i) Or UTpE(i2) = Grp(i) Or UTpE(i3) = Grp(i) Then GrpUtpE(i) = "xx"
                  Next i
                  n += 1 : Strategy_Tpl_Evident_Analyse(U_temp, Strategy_Rslt, Tp, Code_LCR, LCR, Grp, GrpUtpE, n)
                End If
              Next i3
            Next i2
          Next i1
        Case 4 'Quadruples
          n = 0
          For i1 As Integer = 0 To Tp_Nb - 1
            For i2 As Integer = i1 + 1 To Tp_Nb - 1
              For i3 As Integer = i2 + 1 To Tp_Nb - 1
                For i4 As Integer = i3 + 1 To Tp_Nb - 1
                  If Wh_Cell_Nb_Candidats(U_temp, UTpE(i1)) = Tp Or Wh_Cell_Nb_Candidats(U_temp, UTpE(i2)) = Tp Or Wh_Cell_Nb_Candidats(U_temp, UTpE(i2)) = Tp Or Wh_Cell_Nb_Candidats(U_temp, UTpE(i4)) = Tp Then
                    For i As Integer = 0 To 8 : GrpUtpE(i) = "__" : Next i
                    For i As Integer = 0 To 8 : If U_temp(Grp(i), 2) <> " " Then : GrpUtpE(i) = "VV" : End If : Next i
                    For i As Integer = 0 To 8
                      If UTpE(i1) = Grp(i) Or UTpE(i2) = Grp(i) Or UTpE(i3) = Grp(i) Or UTpE(i4) = Grp(i) Then GrpUtpE(i) = "xx"
                    Next i
                    n += 1 : Strategy_Tpl_Evident_Analyse(U_temp, Strategy_Rslt, Tp, Code_LCR, LCR, Grp, GrpUtpE, n)
                  End If
                Next i4
              Next i3
            Next i2
          Next i1
      End Select
    End If
  End Sub

  Sub Strategy_Tpl_Caché_LCR(U_temp(,) As String,
                             ByRef Strategy_Rslt(,) As String,
                             Tp As Integer,
                             Code_LCR As String,
                             LCR As Integer,
                             Grp() As Integer)
    Dim GrpUtpC(9) As String
    Dim n As Integer
    'Détermination du Tuple
    Select Case Tp
      Case 2 'Paires
        n = 0
        For i1 As Integer = 0 To 8
          If U_temp(Grp(i1), 2) = " " Then
            For i2 As Integer = i1 + 1 To 8
              If U_temp(Grp(i2), 2) = " " Then
                For i As Integer = 0 To 8 : GrpUtpC(i) = "__" : Next i
                For i As Integer = 0 To 8 : If U_temp(Grp(i), 2) <> " " Then : GrpUtpC(i) = "VV" : End If : Next i
                For i As Integer = 0 To 8
                  If (Grp(i) = Grp(i1) Or Grp(i) = Grp(i2)) Then GrpUtpC(i) = "xx"
                Next i
                n += 1 : Strategy_Tpl_Caché_Analyse(U_temp, Strategy_Rslt, Tp, Code_LCR, LCR, Grp, GrpUtpC, n)
              End If
            Next i2
          End If
        Next i1
      Case 3 'Triples
        n = 0
        For i1 As Integer = 0 To 8
          If U_temp(Grp(i1), 2) = " " Then
            For i2 As Integer = i1 + 1 To 8
              If U_temp(Grp(i2), 2) = " " Then
                For i3 As Integer = i2 + 1 To 8
                  If U_temp(Grp(i3), 2) = " " Then
                    For i As Integer = 0 To 8 : GrpUtpC(i) = "__" : Next i
                    For i As Integer = 0 To 8 : If U_temp(Grp(i), 2) <> " " Then : GrpUtpC(i) = "VV" : End If : Next i
                    For i As Integer = 0 To 8
                      If (Grp(i) = Grp(i1) Or Grp(i) = Grp(i2) Or Grp(i) = Grp(i3)) Then GrpUtpC(i) = "xx"
                    Next i
                    n += 1 : Strategy_Tpl_Caché_Analyse(U_temp, Strategy_Rslt, Tp, Code_LCR, LCR, Grp, GrpUtpC, n)
                  End If
                Next i3
              End If
            Next i2
          End If
        Next i1
      Case 4 'Quadruples
        n = 0
        For i1 As Integer = 0 To 8
          If U_temp(Grp(i1), 2) = " " Then
            For i2 As Integer = i1 + 1 To 8
              If U_temp(Grp(i2), 2) = " " Then
                For i3 As Integer = i2 + 1 To 8
                  If U_temp(Grp(i3), 2) = " " Then
                    For i4 As Integer = i3 + 1 To 8
                      If U_temp(Grp(i4), 2) = " " Then
                        For i As Integer = 0 To 8 : GrpUtpC(i) = "__" : Next i
                        For i As Integer = 0 To 8 : If U_temp(Grp(i), 2) <> " " Then : GrpUtpC(i) = "VV" : End If : Next i
                        For i As Integer = 0 To 8
                          If (Grp(i) = Grp(i1) Or Grp(i) = Grp(i2) Or Grp(i) = Grp(i3) Or Grp(i) = Grp(i4)) Then GrpUtpC(i) = "xx"
                        Next i
                        n += 1 : Strategy_Tpl_Caché_Analyse(U_temp, Strategy_Rslt, Tp, Code_LCR, LCR, Grp, GrpUtpC, n)
                      End If
                    Next i4
                  End If
                Next i3
              End If
            Next i2
          End If
        Next i1
    End Select
  End Sub

  Sub Strategy_Tpl_Evident_Analyse(U_temp(,) As String,
                                   ByRef Strategy_Rslt(,) As String,
                                   Tp As Integer,
                                   Code_LCR As String,
                                   LCR As Integer,
                                   Grp() As Integer,
                                   GrpUtpE() As String,
                                   n As Integer)
    Dim Stratégie As String = "Tpl"
    Dim Sous_Stratégie As String = "Evi_"

    Dim nb As Integer
    Dim Cd_Str As String
    Dim TpE(10) As Integer ' Permet de compter les candidats Evident 0 ne sert pas
    '                                                                1 comporte le nombre de cd 1
    '                                                                7 comporte le nombre de cd 7
    Dim Cel45(0 To 44) As String
    Dim Exc45(0 To 44) As String
    Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice

    'Dénombrement des candidats
    For cd As Integer = 0 To 9 : TpE(cd) = 0 : Next cd

    For i As Integer = 0 To 8
      For Cd As Integer = 1 To 9
        If U_temp(Grp(i), 2) = " " Then
          'Tuple Evident
          If (U_temp(Grp(i), 3).Contains(CStr(Cd)) = True And GrpUtpE(i) = "xx") Then TpE(Cd) += 1
        End If
      Next Cd
    Next i

    'Tuple Evident
    Dim NbE As Integer = 0
    Dim NbEMax As Integer = 0
    For cd As Integer = 1 To 9
      If TpE(cd) <> 0 Then NbE += 1
      If TpE(cd) > NbEMax Then NbEMax = TpE(cd)
    Next cd

    If (Tp = NbE And Tp = NbEMax) Then
      Cd_Str = "........."
      For cd As Integer = 1 To 9
        If TpE(cd) <> 0 Then Mid$(Cd_Str, cd, 1) = CStr(cd)
      Next cd

      For cd As Integer = 1 To 9
        nb = 0
        For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
        For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
        Cel45_n = -1
        Exc45_n = -1
        For i As Integer = 0 To 8
          If (GrpUtpE(i) = "__" And U_temp(Grp(i), 3).Contains(CStr(cd)) = True And Cd_Str.Contains(CStr(cd)) = True) Then
            nb += 1
          End If
        Next i
        If nb > 0 Then
          For k As Integer = 0 To 8
            If GrpUtpE(k) = "xx" Then
              Cel45_n += 1 : Cel45(Cel45_n) = CStr(Grp(k))
            End If
          Next k
          For k As Integer = 0 To 8
            If (GrpUtpE(k) = "__" And U_temp(Grp(k), 3).Contains(CStr(cd)) = True And Cd_Str.Contains(CStr(cd)) = True) Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp(k)) 'Cellules Concernées par l’exclusion
            End If
          Next k
          Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie & CStr(Tp), Code_LCR, CStr(LCR), CStr(cd), Cel45, Exc45)
        End If '/If nb > 0
      Next cd
    End If
  End Sub

  Sub Strategy_Tpl_Caché_Analyse(U_temp(,) As String,
                                 ByRef Strategy_Rslt(,) As String,
                                 Tp As Integer,
                                 Code_LCR As String,
                                 LCR As Integer,
                                 Grp() As Integer,
                                 GrpUtpC() As String,
                                 n As Integer)
    Dim Stratégie As String = "Tpl"
    Dim Sous_Stratégie As String = "Cac_"

    Dim nb As Integer
    Dim Cd_Str As String
    Dim TpC(10) As Integer ' Permet de compter les candidats Caché   0 ne sert pas
    '                                                                1 comporte le nombre de cd 1
    '                                                                7 comporte le nombre de cd 7
    Dim TpA(10) As Integer
    For Cd As Integer = 0 To 9 : TpC(Cd) = 0 : Next Cd
    For Cd As Integer = 0 To 9 : TpA(Cd) = 0 : Next Cd
    Dim Cel45(0 To 44) As String
    Dim Exc45(0 To 44) As String
    Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice

    'Dénombrement des candidats
    For i As Integer = 0 To 8
      For Cd As Integer = 1 To 9
        If U_temp(Grp(i), 2) = " " Then
          If (U_temp(Grp(i), 3).Contains(CStr(Cd)) And GrpUtpC(i) = "xx") = True Then TpC(Cd) += 1
          If (U_temp(Grp(i), 3).Contains(CStr(Cd)) And GrpUtpC(i) = "__") = True Then TpA(Cd) += 1
        End If
      Next Cd
    Next i

    'Tuple Caché
    Dim NbC As Integer = 0
    Dim NbCMax As Integer = 0
    Dim NbA As Integer = 0
    For Cd As Integer = 1 To 9
      If TpC(Cd) > NbCMax Then NbCMax = TpC(Cd)
      If TpA(Cd) = 0 Then NbA += 1
    Next Cd

    For Cd As Integer = 1 To 9
      If (TpC(Cd) <> 0 And TpA(Cd) = 0 And NbCMax = Tp) Then NbC += 1
    Next Cd

    If NbC = Tp Then
      Cd_Str = "........."
      For cd As Integer = 1 To 9
        If (TpC(cd) <> 0 And TpA(cd) = 0) Then Mid$(Cd_Str, cd, 1) = CStr(cd)
      Next cd

      For cd As Integer = 1 To 9
        nb = 0
        For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
        For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
        Cel45_n = -1
        Exc45_n = -1
        For i As Integer = 0 To 8
          If (GrpUtpC(i) = "xx" And U_temp(Grp(i), 3).Contains(CStr(cd)) = True And Cd_Str.Contains(CStr(cd)) = False) Then
            nb += 1
          End If
        Next i
        If nb > 0 Then
          For k As Integer = 0 To 8
            If GrpUtpC(k) = "__" Then
              Cel45_n += 1 : Cel45(Cel45_n) = CStr(Grp(k))
            End If
          Next k
          For k As Integer = 0 To 8
            If (GrpUtpC(k) = "xx" And U_temp(Grp(k), 3).Contains(CStr(cd)) = True And Cd_Str.Contains(CStr(cd)) = False) Then
              Exc45_n += 1 : Exc45(Exc45_n) = CStr(Grp(k)) 'Cellules Concernées par l’exclusion
            End If
          Next k
          Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie & CStr(Tp), Code_LCR, CStr(LCR), CStr(cd), Cel45, Exc45)
        End If '/If nb > 0
      Next cd
    End If
  End Sub
End Module