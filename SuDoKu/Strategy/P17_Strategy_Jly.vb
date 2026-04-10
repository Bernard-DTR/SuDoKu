'-------------------------------------------------------------------------------
'  Ce module comporte la stratégie Jellyfish, identique à swordfish, mais en 4x4
'-------------------------------------------------------------------------------

Friend Module P17_Strategy_Jly
  Function Strategy_Jly(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "Jly"
    Dim Sous_Stratégie As String = "Jly"
    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Xwg

    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 
    For Cdd As Integer = 1 To 9
      Strategy_Jly_Row(U_temp, Strategy_Rslt, Cdd)
      Strategy_Jly_Col(U_temp, Strategy_Rslt, Cdd)
    Next Cdd

    Return Strategy_Rslt
  End Function

  Sub Strategy_Jly_Row(U_temp(,) As String,
                       ByRef Strategy_Rslt(,) As String,
                       Cdd As Integer)
    Dim XY1(9, 9) As String ' col, row, comporte les 9 colonnes et 1 colonne à droite pour les totaux
    '                                   comporte les 9 lignes  et 1 ligne   en bas   pour les totaux
    Dim n As Integer
    'Initialisation de XY1
    For row As Integer = 0 To 9
      For col As Integer = 0 To 9
        XY1(col, row) = "X" 'Le tableau est string et initialisé
      Next col
    Next row

    'Dénombrement des candidats
    For i As Integer = 0 To 80
      If U_temp(i, 3).Contains(CStr(Cdd)) = True Then
        XY1(U_Col(i), U_Row(i)) = CStr(Cdd)
      Else
        XY1(U_Col(i), U_Row(i)) = "_"
      End If
    Next i

    'Comptage par Ligne
    For row As Integer = 0 To 8
      n = 0
      For col As Integer = 0 To 8
        If XY1(col, row) = CStr(Cdd) Then n += 1
      Next col
      XY1(9, row) = CStr(n)
    Next row

    n = 0
    For row As Integer = 0 To 8
      If XY1(9, row) >= "2" Then n += 1
    Next row
    Dim nRow As Integer = 0 ' Il faut au moins 4 lignes de 2, de 3 ou de 4
    For i As Integer = 0 To 8
      If CInt(XY1(9, i)) = 2 Or CInt(XY1(9, i)) = 3 Or CInt(XY1(9, i)) = 4 Then nRow += 1
    Next i
    If nRow < 4 Then GoTo Strategy_Jly_Row_End

    'Combinaison des 4 Lignes
    Dim Row1 As Integer
    Dim Row2 As Integer
    Dim Row3 As Integer
    Dim Row4 As Integer
    For i As Integer = 0 To 8
      If CInt(XY1(9, i)) = 2 Or CInt(XY1(9, i)) = 3 Or CInt(XY1(9, i)) = 4 Then
        Row1 = i
        For j As Integer = i + 1 To 8
          If CInt(XY1(9, j)) = 2 Or CInt(XY1(9, j)) = 3 Or CInt(XY1(9, j)) = 4 Then
            Row2 = j
            For k As Integer = j + 1 To 8
              If CInt(XY1(9, k)) = 2 Or CInt(XY1(9, k)) = 3 Or CInt(XY1(9, k)) = 4 Then
                Row3 = k
                For l As Integer = k + 1 To 8
                  If CInt(XY1(9, l)) = 2 Or CInt(XY1(9, l)) = 3 Or CInt(XY1(9, l)) = 4 Then
                    Row4 = l
                    'Etude du groupe des 4 lignes
                    Strategy_Jly_Row_Etude(Strategy_Rslt, Cdd, XY1, Row1, Row2, Row3, Row4)
                  End If
                Next l
              End If
            Next k
          End If
        Next j
      End If
    Next i

Strategy_Jly_Row_End:
  End Sub

  Sub Strategy_Jly_Col(U_temp(,) As String,
                       ByRef Strategy_Rslt(,) As String,
                       Cdd As Integer)
    Dim XY1(9, 9) As String ' col, row, comporte les 9 colonnes et 1 colonne à droite pour les totaux
    '                                   comporte les 9 lignes   et 1 ligne   en bas   pour les totaux
    Dim n As Integer
    'Initialisation de XY1
    For row As Integer = 0 To 9
      For col As Integer = 0 To 9
        XY1(col, row) = "X" 'Le tableau est string et initialisé
      Next col
    Next row

    'Dénombrement des candidats
    For i As Integer = 0 To 80
      If U_temp(i, 3).Contains(CStr(Cdd)) = True Then
        XY1(U_Col(i), U_Row(i)) = CStr(Cdd)
      Else
        XY1(U_Col(i), U_Row(i)) = "_"
      End If
    Next i

    'Comptage par colonne
    For col As Integer = 0 To 8
      n = 0
      For row As Integer = 0 To 8
        If XY1(col, row) = CStr(Cdd) Then n += 1
      Next row
      XY1(col, 9) = CStr(n)
    Next col

    n = 0
    For col As Integer = 0 To 8
      If XY1(col, 9) = "2" Then n += 1
    Next col
    Dim nCol As Integer = 0 ' Il faut au moins 4 colonnes de 2, de 3 ou de 4
    For i As Integer = 0 To 8
      If CInt(XY1(i, 9)) = 2 Or CInt(XY1(i, 9)) = 3 Or CInt(XY1(i, 9)) = 4 Then nCol += 1
    Next i
    If nCol < 4 Then GoTo Strategy_Jly_Col_End

    'Combinaison des 3 Colonnes
    Dim Col1 As Integer
    Dim Col2 As Integer
    Dim Col3 As Integer
    Dim Col4 As Integer
    For i As Integer = 0 To 8
      If CInt(XY1(i, 9)) = 2 Or CInt(XY1(i, 9)) = 3 Or CInt(XY1(i, 9)) = 4 Then
        Col1 = i
        For j As Integer = i + 1 To 8
          If CInt(XY1(j, 9)) = 2 Or CInt(XY1(j, 9)) = 3 Or CInt(XY1(j, 9)) = 4 Then
            Col2 = j
            For k As Integer = j + 1 To 8
              If CInt(XY1(k, 9)) = 2 Or CInt(XY1(k, 9)) = 3 Or CInt(XY1(k, 9)) = 4 Then
                Col3 = k
                For l As Integer = k + 1 To 8
                  If CInt(XY1(l, 9)) = 2 Or CInt(XY1(l, 9)) = 3 Or CInt(XY1(l, 9)) = 4 Then
                    Col4 = l
                    'Etude des 4 Colonnes
                    Strategy_Jly_Col_Etude(Strategy_Rslt, Cdd, XY1, Col1, Col2, Col3, Col4)
                  End If
                Next l
              End If
            Next k
          End If
        Next j
      End If
    Next i
Strategy_Jly_Col_End:
  End Sub

  Sub Strategy_Jly_Row_Etude(ByRef Strategy_Rslt(,) As String,
                             Cdd As Integer,
                             XY1(,) As String,
                             Row1 As Integer, Row2 As Integer, Row3 As Integer, Row4 As Integer)
    ' XY1(col, row)
    ' Il faut que le candidat Cdd des Lignes Row_1234 occupe 4 colonnes exactement
    Dim Col(0 To 8) As Integer
    For i As Integer = 0 To 8 : Col(i) = 0 : Next i
    For i As Integer = 0 To 8
      If XY1(i, Row1) = CStr(Cdd) Then Col(i) = 1
      If XY1(i, Row2) = CStr(Cdd) Then Col(i) = 1
      If XY1(i, Row3) = CStr(Cdd) Then Col(i) = 1
      If XY1(i, Row4) = CStr(Cdd) Then Col(i) = 1
    Next i
    Dim n As Integer = 0
    For i As Integer = 0 To 8
      If Col(i) = 1 Then n += 1
    Next i
    If n <> 4 Then GoTo Strategy_Jly_Row_Etude_End

    'Quelles sont ces 4 colonnes ?
    Dim Col1 As Integer = -1
    Dim Col2 As Integer = -1
    Dim Col3 As Integer = -1
    Dim Col4 As Integer = -1
    n = 0
    For i As Integer = 0 To 8
      If Col(i) = 1 Then
        n += 1
        If n = 1 Then Col1 = i
        If n = 2 Then Col2 = i
        If n = 3 Then Col3 = i
        If n = 4 Then Col4 = i
      End If
    Next i

    'Cas de Jellyfish
    'Il faut stocker dans Cel45 les cellules des 4 Lignes
    ' et dans Exc45 les cellules des colonnes sauf les cellules des lignes
    ' s'il y a au moins une cellule Exc45, alors c'est un Jellyfish
    Dim Stratégie As String = "Jly"
    Dim Sous_Stratégie As String = "Jly_R"
    Dim Candidat As String = CStr(Cdd)
    Dim Code_LCR As String = "#"
    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice
    Dim Cellule As Integer
    'Documentation des Cel45 avec les 4 lignes
    For i As Integer = 0 To 8 'Ligne 1
      If XY1(i, Row1) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row1)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 2
      If XY1(i, Row2) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row2)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 3
      If XY1(i, Row3) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row3)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 4
      If XY1(i, Row4) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row4)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    'Documentation des Exc45 avec les 4 colonnes qui comportent un Cdd sauf celles des Lignes
    For i As Integer = 0 To 8 'Colonne 1
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If i = Row4 Then Continue For
      If XY1(Col1, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col1, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 2
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If i = Row4 Then Continue For
      If XY1(Col2, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col2, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 3
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If i = Row4 Then Continue For
      If XY1(Col3, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col3, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 4
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If i = Row4 Then Continue For
      If XY1(Col4, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col4, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    If Exc45_n <> -1 Then
      'Il y a un Jellyfish
      Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "-1", Candidat, Cel45, Exc45)
    End If

Strategy_Jly_Row_Etude_End:
  End Sub

  Sub Strategy_Jly_Col_Etude(ByRef Strategy_Rslt(,) As String,
                             Cdd As Integer,
                             XY1(,) As String,
                             Col1 As Integer, Col2 As Integer, Col3 As Integer, Col4 As Integer)
    ' XY1(col, row)
    ' Il faut que le candidat Cdd des Colonnes Col_1234 occupe 4 lignes exactement
    Dim Row(0 To 8) As Integer
    For i As Integer = 0 To 8 : Row(i) = 0 : Next i
    For i As Integer = 0 To 8
      If XY1(Col1, i) = CStr(Cdd) Then Row(i) = 1
      If XY1(Col2, i) = CStr(Cdd) Then Row(i) = 1
      If XY1(Col3, i) = CStr(Cdd) Then Row(i) = 1
      If XY1(Col4, i) = CStr(Cdd) Then Row(i) = 1
    Next i
    Dim n As Integer = 0
    For i As Integer = 0 To 8
      If Row(i) = 1 Then n += 1
    Next i
    If n <> 4 Then GoTo Strategy_Jly_Col_Etude_End

    'Quelles sont ces 4 lignes ?
    Dim Row1 As Integer = -1
    Dim Row2 As Integer = -1
    Dim Row3 As Integer = -1
    Dim Row4 As Integer = -1
    n = 0
    For i As Integer = 0 To 8
      If Row(i) = 1 Then
        n += 1
        If n = 1 Then Row1 = i
        If n = 2 Then Row2 = i
        If n = 3 Then Row3 = i
        If n = 4 Then Row4 = i
      End If
    Next i

    'Cas de Jellyfish
    'Il faut stocker dans Cel45 les cellules des Lignes
    ' et dans Exc45 les cellules des colonnes sauf les cellules des lignes
    ' s'il y a au moins une cellule Exc45, alors c'est un Jellyfish
    Dim Stratégie As String = "Jly"
    Dim Sous_Stratégie As String = "Jly_C"
    Dim Candidat As String = CStr(Cdd)
    Dim Code_LCR As String = "#"
    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Cel45_n As Integer = -1
    Dim Exc45_n As Integer = -1
    Dim Cellule As Integer
    'Documentation des Cel45 avec les 4 colonnes
    For i As Integer = 0 To 8 'Colonne 1
      If XY1(Col1, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col1, i)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 2
      If XY1(Col2, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col2, i)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 3
      If XY1(Col3, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col3, i)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 4
      If XY1(Col4, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col4, i)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    'Documentation des Exc45 avec les 4 lignes qui comportent un Cdd sauf celles des Colonnes
    For i As Integer = 0 To 8 'Ligne 1
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If i = Col4 Then Continue For
      If XY1(i, Row1) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row1)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 2
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If i = Col4 Then Continue For
      If XY1(i, Row2) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row2)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 3
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If i = Col4 Then Continue For
      If XY1(i, Row3) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row3)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 4
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If i = Col4 Then Continue For
      If XY1(i, Row4) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row4)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    If Exc45_n <> -1 Then
      'Il y a un Jellyfish
      Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "-1", Candidat, Cel45, Exc45)
    End If

Strategy_Jly_Col_Etude_End:
  End Sub
End Module