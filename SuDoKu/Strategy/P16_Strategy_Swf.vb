'-------------------------------------------------------------------------------
'
' Stratégie des Swordfish
' Préfixe       Swf
'
' U_temp doit être utilisé en lieu et place de U qui ne peut pas être utilisé
' Méthodologie: Traitement par Colonne et ensuite traitement par Ligne
'               Il faut qu'il y ait pour un candidat cdd 2 présences dans 3 colonnes
'               ou                                       2 présences dans 3 lignes
'
'  La stratégie est complétée pour prendre en compte les groupes de 3 lignes ou de 3 colonnes
'     et résoudre ainsi les puzzles de HODOKU
'-------------------------------------------------------------------------------


Friend Module P16_Strategy_Swf
  Function Strategy_Swf(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "Swf"
    Dim Sous_Stratégie As String = "Swf"
    '
    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Xwg
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    For Cdd As Integer = 1 To 9
      Strategy_Swf_Row(U_temp, Strategy_Rslt, Cdd)
      Strategy_Swf_Col(U_temp, Strategy_Rslt, Cdd)
    Next Cdd

    Return Strategy_Rslt
  End Function

  Sub Strategy_Swf_Row(U_temp(,) As String,
                       ByRef Strategy_Rslt(,) As String,
                       Cdd As Integer)
    Dim XY1(9, 9) As String ' col, row, comporte les 9 colonnes et 1 colonne à droite pour les totaux
    '                                   comporte les 9 lignes   et 1 ligne  en bas   pour les totaux
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
    Dim nRow As Integer = 0 ' Il faut au moins 3 lignes de 2 ou 3
    For i As Integer = 0 To 8
      If CInt(XY1(9, i)) = 2 Or CInt(XY1(9, i)) = 3 Then nRow += 1
    Next i
    If nRow < 3 Then GoTo Strategy_Swf_Row_End


    'Combinaison des 3 lignes
    Dim Row1 As Integer
    Dim Row2 As Integer
    Dim Row3 As Integer
    For i As Integer = 0 To 8
      If CInt(XY1(9, i)) = 2 Or CInt(XY1(9, i)) = 3 Then
        Row1 = i
        For j As Integer = i + 1 To 8
          If CInt(XY1(9, j)) = 2 Or CInt(XY1(9, j)) = 3 Then
            Row2 = j
            For k As Integer = j + 1 To 8
              If CInt(XY1(9, k)) = 2 Or CInt(XY1(9, k)) = 3 Then
                Row3 = k
                'Etude du groupe des 3 lignes
                Strategy_Swf_Row_Etude(Strategy_Rslt, Cdd, XY1, Row1, Row2, Row3)
              End If
            Next k
          End If
        Next j
      End If
    Next i

Strategy_Swf_Row_End:
  End Sub

  Sub Strategy_Swf_Col(U_temp(,) As String,
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
    Dim nCol As Integer = 0 ' Il faut au moins 3 colonnes de 2 ou de 3
    For i As Integer = 0 To 8
      If CInt(XY1(i, 9)) = 2 Or CInt(XY1(i, 9)) = 3 Then nCol += 1
    Next i
    If nCol < 3 Then GoTo Strategy_Swf_Col_End


    'Combinaison des 3 Colonnes
    Dim Col1 As Integer
    Dim Col2 As Integer
    Dim Col3 As Integer
    For i As Integer = 0 To 8
      If CInt(XY1(i, 9)) = 2 Or CInt(XY1(i, 9)) = 3 Then
        Col1 = i
        For j As Integer = i + 1 To 8
          If CInt(XY1(j, 9)) = 2 Or CInt(XY1(j, 9)) = 3 Then
            Col2 = j
            For k As Integer = j + 1 To 8
              If CInt(XY1(k, 9)) = 2 Or CInt(XY1(k, 9)) = 3 Then
                Col3 = k
                'Etude des 3 Colonnes
                Strategy_Swf_Col_Etude(Strategy_Rslt, Cdd, XY1, Col1, Col2, Col3)
              End If
            Next k
          End If
        Next j
      End If
    Next i
Strategy_Swf_Col_End:
  End Sub

  Sub Strategy_Swf_Row_Etude(ByRef Strategy_Rslt(,) As String,
                             Cdd As Integer,
                             XY1(,) As String,
                             Row1 As Integer, Row2 As Integer, Row3 As Integer)
    ' XY1(col, row)
    ' Il faut que le candidat Cdd des lignes Row_123 occupe que 3 colonnes exactement
    Dim Col(0 To 8) As Integer
    For i As Integer = 0 To 8 : Col(i) = 0 : Next i
    For i As Integer = 0 To 8
      If XY1(i, Row1) = CStr(Cdd) Then Col(i) = 1
      If XY1(i, Row2) = CStr(Cdd) Then Col(i) = 1
      If XY1(i, Row3) = CStr(Cdd) Then Col(i) = 1
    Next i
    Dim n As Integer = 0
    For i As Integer = 0 To 8
      If Col(i) = 1 Then n += 1
    Next i
    If n <> 3 Then GoTo Strategy_Swf_Row_Etude_End

    'Quelles sont ces 3 colonnes ?
    Dim Col1 As Integer = -1
    Dim Col2 As Integer = -1
    Dim Col3 As Integer = -1
    n = 0
    For i As Integer = 0 To 8
      If Col(i) = 1 Then
        n += 1
        If n = 1 Then Col1 = i
        If n = 2 Then Col2 = i
        If n = 3 Then Col3 = i
      End If
    Next i

    'Cas de Swordfish
    'Il faut stocker dans Cel10 les cellules des 3 Lignes
    ' et dans Exc20 les cellules des colonnes sauf les cellules des lignes
    ' s'il y a au moins une cellule Exc20, alors c'est un swordfish
    Dim Stratégie As String = "Swf"
    Dim Sous_Stratégie As String = "Swf_R"
    Dim Candidat As String = CStr(Cdd)
    Dim Code_LCR As String = "#"
    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice
    Dim Cellule As Integer
    'Documentation des Cel10 avec les 3 lignes
    For i As Integer = 0 To 8 'Ligne 1
      If XY1(i, Row1) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row1)
        Cel45_n += 1 : Cel45(Cel45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 1
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
    'Documentation des Exc20 avec les 3 colonnes qui comportent un Cdd sauf celles des Lignes
    For i As Integer = 0 To 8 'Colonne 1
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If XY1(Col1, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col1, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 2
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If XY1(Col2, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col2, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Colonne 3
      If i = Row1 Then Continue For
      If i = Row2 Then Continue For
      If i = Row3 Then Continue For
      If XY1(Col3, i) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(Col3, i)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    If Exc45_n <> -1 Then
      'Il y a un swordfish
      Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "-1", Candidat, Cel45, Exc45)
    End If

Strategy_Swf_Row_Etude_End:
  End Sub

  Sub Strategy_Swf_Col_Etude(ByRef Strategy_Rslt(,) As String,
                             Cdd As Integer,
                             XY1(,) As String,
                             Col1 As Integer, Col2 As Integer, Col3 As Integer)
    ' XY1(col, row)
    ' Il faut que le candidat Cdd des Colonnes Col_123 occupe que 3 Lignes exactement
    Dim Row(0 To 8) As Integer
    For i As Integer = 0 To 8 : Row(i) = 0 : Next i
    For i As Integer = 0 To 8
      If XY1(Col1, i) = CStr(Cdd) Then Row(i) = 1
      If XY1(Col2, i) = CStr(Cdd) Then Row(i) = 1
      If XY1(Col3, i) = CStr(Cdd) Then Row(i) = 1
    Next i
    Dim n As Integer = 0
    For i As Integer = 0 To 8
      If Row(i) = 1 Then n += 1
    Next i
    If n <> 3 Then GoTo Strategy_Swf_Col_Etude_End

    'Quelles sont ces 3 lignes ?
    Dim Row1 As Integer = -1
    Dim Row2 As Integer = -1
    Dim Row3 As Integer = -1
    n = 0
    For i As Integer = 0 To 8
      If Row(i) = 1 Then
        n += 1
        If n = 1 Then Row1 = i
        If n = 2 Then Row2 = i
        If n = 3 Then Row3 = i
      End If
    Next i

    'Cas de Swordfish
    'Il faut stocker dans Cel10 les cellules des Lignes
    ' et dans Exc20 les cellules des colonnes sauf les cellules des lignes
    ' s'il y a au moins une cellule Exc20, alors c'est un swordfish
    Dim Stratégie As String = "Swf"
    Dim Sous_Stratégie As String = "Swf_C"
    Dim Candidat As String = CStr(Cdd)
    Dim Code_LCR As String = "#"
    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice
    Dim Cellule As Integer
    'Documentation des Cel10 avec les 3 colonnes
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
    'Documentation des Exc20 avec les 3 lignes qui comportent un Cdd sauf celles des Colonnes
    For i As Integer = 0 To 8 'Ligne 1
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If XY1(i, Row1) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row1)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 2
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If XY1(i, Row2) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row2)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i
    For i As Integer = 0 To 8 'Ligne 3
      If i = Col1 Then Continue For
      If i = Col2 Then Continue For
      If i = Col3 Then Continue For
      If XY1(i, Row3) = CStr(Cdd) Then
        Cellule = Wh_Cellule_ColRow(i, Row3)
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cellule)
      End If
    Next i

    If Exc45_n <> -1 Then
      'Il y a un swordfish
      Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, Code_LCR, "-1", Candidat, Cel45, Exc45)
    End If

Strategy_Swf_Col_Etude_End:
  End Sub
End Module