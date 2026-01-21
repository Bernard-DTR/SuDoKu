Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------
'
' Stratégie des XYZ_Wing
' Préfixe       XYZ
'
'-------------------------------------------------------------------------------

Friend Module P18_Strategy_XYZ
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Function Strategy_XYZ(U_temp(,) As String) As String(,)
    Dim Cdd_X As String = ""
    Dim Cdd_Y As String = ""
    Dim Cdd_Z As String = ""
    Dim n As Integer = -1
    Dim Stratégie As String = "XYZ"
    Dim Sous_Stratégie As String = "XYZ"

    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Xwg
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    Dim XYZ_t(6, 81) As String      'Taille maximale 81,  XYZ_t_Nb représente le nombre de lignes remplies de 0 à x
    '         0  Numéro du poste    'Comporte les cellules à 3 candidats
    '         1  Cellule
    '         2  Cellule en coord
    '         3  Candidats
    '         4  Candidat 1
    '         5  Candidat 2
    '         6  Candidat 3
    Dim XYZ_t_Nb As Integer = -1
    For i As Integer = 0 To 6
      For j As Integer = 0 To 80
        XYZ_t(i, j) = "" 'Le tableau est string et initialisé
      Next j
    Next i

    ' 01 Analyse des Cellules à 3 candidats
    '    Les cellules à 3 candidats sont stockées dans XYZ_t
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = 3 Then
        n += 1
        XYZ_t(0, n) = CStr(n)
        XYZ_t(1, n) = CStr(i) : XYZ_t(2, n) = U_cr(i)
        Dim n1 As Integer = 0 'Premier, deuxième et troisième candidat
        For j As Integer = 0 To 8
          If Mid$(U_temp(i, 3), j + 1, 1) <> " " Then
            n1 += 1
            If n1 = 1 Then Cdd_X = Mid$(U_temp(i, 3), j + 1, 1)
            If n1 = 2 Then Cdd_Y = Mid$(U_temp(i, 3), j + 1, 1)
            If n1 = 3 Then Cdd_Z = Mid$(U_temp(i, 3), j + 1, 1)
          End If
        Next j
        XYZ_t(3, n) = U_temp(i, 3) : XYZ_t(4, n) = Cdd_X : XYZ_t(5, n) = Cdd_Y : XYZ_t(6, n) = Cdd_Z
        XYZ_t_Nb = n  'Détermination de XYZ_t_Ubound, dernier poste rempli
      End If
    Next i
    If XYZ_t_Nb = -1 Then GoTo Strategy_XYZ_End

    ' 02 Analyse des Cellules à 3 candidats
    '    Combinaison des 3 candidats 
    'X   XY et XZ
    'Y   XY et YZ
    'Z   XZ et YZ
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim Row_xyz As Integer = -1
    Dim Col_xyz As Integer = -1
    Dim Cel_xyz As Integer = -1
    Dim Cdd_XY As String = Cnddts_Blancs
    Dim Cdd_XZ As String = Cnddts_Blancs
    Dim Cdd_YZ As String = Cnddts_Blancs
    For i As Integer = 0 To XYZ_t_Nb
      Cel_xyz = CInt(XYZ_t(1, i))
      Row_xyz = U_Row(Cel_xyz)
      Col_xyz = U_Col(Cel_xyz)
      Cdd_X = XYZ_t(4, i) : Cdd_Y = XYZ_t(5, i) : Cdd_Z = XYZ_t(6, i)

      'X   XY et XZ
      Cdd_XY = Cnddts_Blancs : Cdd_XZ = Cnddts_Blancs : Cdd_YZ = Cnddts_Blancs
      Mid$(Cdd_XY, CInt(Cdd_X), 1) = Cdd_X : Mid$(Cdd_XY, CInt(Cdd_Y), 1) = Cdd_Y
      Mid$(Cdd_XZ, CInt(Cdd_X), 1) = Cdd_X : Mid$(Cdd_XZ, CInt(Cdd_Z), 1) = Cdd_Z
      Strategy_XYZ_A1(U_temp, Strategy_Rslt, "XY-XZ", CInt(XYZ_t(1, i)), Cdd_XY, Cdd_XZ, Cdd_X)
      Strategy_XYZ_A2(U_temp, Strategy_Rslt, "XY-XZ", CInt(XYZ_t(1, i)), Cdd_XY, Cdd_XZ, Cdd_X)

      'Y   XY et YZ
      Cdd_XY = Cnddts_Blancs : Cdd_XZ = Cnddts_Blancs : Cdd_YZ = Cnddts_Blancs
      Mid$(Cdd_XY, CInt(Cdd_X), 1) = Cdd_X : Mid$(Cdd_XY, CInt(Cdd_Y), 1) = Cdd_Y
      Mid$(Cdd_YZ, CInt(Cdd_Y), 1) = Cdd_Y : Mid$(Cdd_YZ, CInt(Cdd_Z), 1) = Cdd_Z
      Strategy_XYZ_A1(U_temp, Strategy_Rslt, "XY-YZ", CInt(XYZ_t(1, i)), Cdd_XY, Cdd_YZ, Cdd_Y)
      Strategy_XYZ_A2(U_temp, Strategy_Rslt, "XY-YZ", CInt(XYZ_t(1, i)), Cdd_XY, Cdd_YZ, Cdd_Y)

      'Z   XZ et YZ
      Cdd_XY = Cnddts_Blancs : Cdd_XZ = Cnddts_Blancs : Cdd_YZ = Cnddts_Blancs
      Mid$(Cdd_XZ, CInt(Cdd_X), 1) = Cdd_X : Mid$(Cdd_XZ, CInt(Cdd_Z), 1) = Cdd_Z
      Mid$(Cdd_YZ, CInt(Cdd_Y), 1) = Cdd_Y : Mid$(Cdd_YZ, CInt(Cdd_Z), 1) = Cdd_Z
      Strategy_XYZ_A1(U_temp, Strategy_Rslt, "XZ-YZ", CInt(XYZ_t(1, i)), Cdd_XZ, Cdd_YZ, Cdd_Z)
      Strategy_XYZ_A2(U_temp, Strategy_Rslt, "XZ-YZ", CInt(XYZ_t(1, i)), Cdd_XZ, Cdd_YZ, Cdd_Z)

    Next i

Strategy_XYZ_End:
    Return Strategy_Rslt
  End Function

  Sub Strategy_XYZ_A1(U_temp(,) As String,
                      ByRef Strategy_Rslt(,) As String,
                      Combinaison_XYZ As String,
                      Cel_XYZ As Integer,
                      Cdd_XZ As String,
                      Cdd_YZ As String,
                      Cdd_Z As String)
    ' 03 Analyse des Cellules à 3 candidats
    Dim XZ_Row As Boolean
    Dim YZ_Reg As Boolean
    Dim Cel45(0 To 44) As String
    Dim Exc45(0 To 44) As String
    Dim Cel_XZ As Integer = -1
    Dim Cel_YZ As Integer = -1
    Dim Cel As Integer

    'A1
    'XYZ est dans une Région      XZ sur la ligne, mais pas dans la région
    '                             YZ dans la région, mais pas sur la ligne
    '  Z ne peut se trouver sur la ligne de la région

    XZ_Row = False
    Dim Grp As Integer() = U_9CelRow(U_Row(Cel_XYZ))                     'Ligne de XYZ
    For i As Integer = 0 To 8
      Cel = Grp(i)
      If U_Reg(Cel) = U_Reg(Cel_XYZ) Then Continue For 'ne doivent pas être dans la même région
      If U_temp(Cel, 3) = Cdd_XZ Then
        XZ_Row = True
        Cel_XZ = Cel
        Exit For
      End If
    Next i

    YZ_Reg = False
    Grp = U_9CelReg(U_Reg(Cel_XYZ))                      'Région de XYZ
    For i As Integer = 0 To 8
      Cel = Grp(i)
      If U_Row(Cel) = U_Row(Cel_XYZ) Then Continue For  'ne doivent pas être sur la même ligne
      If U_temp(Cel, 3) = Cdd_YZ Then
        YZ_Reg = True
        Cel_YZ = Cel
        Exit For
      End If
    Next i

    If XZ_Row = True And YZ_Reg = True Then
      For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
      For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
      Grp = U_9CelReg(U_Reg(Cel_XYZ))
      Dim CelExc1 As Integer = -1
      Dim CelExc2 As Integer = -1
      Dim n As Integer = 0
      For i As Integer = 0 To 8
        Cel = Grp(i)
        If U_Row(Cel) = U_Row(Cel_XYZ) Then
          If Cel = Cel_XYZ Then Continue For        '  Pas de traitement pour la cellule XYZ
          If U_temp(Cel, 3).Contains(Cdd_Z) = True Then
            n += 1
            If n = 1 Then CelExc1 = Cel
            If n = 2 Then CelExc2 = Cel
          End If
        End If
      Next i
      Cel45(0) = CStr(Cel_XYZ) : Cel45(1) = CStr(Cel_XZ) : Cel45(2) = CStr(Cel_YZ)
      If CelExc1 <> -1 Then Exc45(0) = CStr(CelExc1)
      If CelExc2 <> -1 Then Exc45(1) = CStr(CelExc2)
      If n > 0 Then Strategy_Rslt_Add(Strategy_Rslt, "XYZ", Combinaison_XYZ & "_A1", "X", "-1", Cdd_Z, Cel45, Exc45) 'Sous-Stratégie A1, Code_LCR="X"
    End If
  End Sub

  Sub Strategy_XYZ_A2(U_temp(,) As String,
                      ByRef Strategy_Rslt(,) As String,
                      Combinaison_XYZ As String,
                      Cel_XYZ As Integer,
                      Cdd_XZ As String,
                      Cdd_YZ As String,
                      Cdd_Z As String)
    ' 03 Analyse des Cellules à 3 candidats
    Dim XZ_Col As Boolean
    Dim YZ_Reg As Boolean
    Dim Cel45(0 To 44) As String
    Dim Exc45(0 To 44) As String
    Dim Cel_XZ As Integer = -1
    Dim Cel_YZ As Integer = -1
    Dim Cel As Integer

    'A2
    'XYZ est dans une Région	  XZ dans la colonne, mais pas dans la région
    '                             YZ dans la région, mais pas dans la colonne 
    '  Z ne peut se trouver sur la colonne de la région

    XZ_Col = False
    Dim Grp As Integer() = U_9CelCol(U_Col(Cel_XYZ))                     'Colonne de XYZ
    For i As Integer = 0 To 8
      Cel = Grp(i)
      If U_Reg(Cel) = U_Reg(Cel_XYZ) Then Continue For 'ne doivent pas être dans la même région
      If U_temp(Cel, 3) = Cdd_XZ Then
        XZ_Col = True
        Cel_XZ = Cel
        Exit For
      End If
    Next i

    YZ_Reg = False
    Grp = U_9CelReg(U_Reg(Cel_XYZ))                      'Région de XYZ
    For i As Integer = 0 To 8
      Cel = Grp(i)
      If U_Col(Cel) = U_Col(Cel_XYZ) Then Continue For  'ne doivent pas être sur la même colonne
      If U_temp(Cel, 3) = Cdd_YZ Then
        YZ_Reg = True
        Cel_YZ = Cel
        Exit For
      End If
    Next i

    If XZ_Col = True And YZ_Reg = True Then
      For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
      For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
      Grp = U_9CelReg(U_Reg(Cel_XYZ))
      Dim CelExc1 As Integer = -1
      Dim CelExc2 As Integer = -1
      Dim n As Integer = 0
      For i As Integer = 0 To 8
        Cel = Grp(i)
        If U_Col(Cel) = U_Col(Cel_XYZ) Then
          If Cel = Cel_XYZ Then Continue For        '  Pas de traitement pour la cellule XYZ
          If U_temp(Cel, 3).Contains(Cdd_Z) = True Then
            n += 1
            If n = 1 Then CelExc1 = Cel
            If n = 2 Then CelExc2 = Cel
          End If
        End If
      Next i
      Cel45(0) = CStr(Cel_XYZ) : Cel45(1) = CStr(Cel_XZ) : Cel45(2) = CStr(Cel_YZ)
      If CelExc1 <> -1 Then Exc45(0) = CStr(CelExc1)
      If CelExc2 <> -1 Then Exc45(1) = CStr(CelExc2)
      If n > 0 Then Strategy_Rslt_Add(Strategy_Rslt, "XYZ", Combinaison_XYZ & "_A2", "X", "-1", Cdd_Z, Cel45, Exc45) 'Sous-Stratégie A2, Code_LCR="X"
    End If
  End Sub

  Sub Display_XYZ_Display(XYZ_t(,) As String, XYZ_t_Nb As Integer)
    'Contrôle : Liste de XYZ_t
    Dim S As String
    Jrn_Add(, {"Liste de XYZ_t : " & CStr(XYZ_t_Nb)})
    For i As Integer = 0 To XYZ_t_Nb
      S = XYZ_t(0, i).PadLeft(2) & " " & XYZ_t(1, i).PadLeft(2) & " " & XYZ_t(2, i) & " _ " & XYZ_t(3, i) & " _ " & XYZ_t(4, i) & " " & XYZ_t(5, i) & " " & XYZ_t(6, i)
      Jrn_Add(, {S})
    Next i
    Jrn_Add(, {"/Liste de XYZ_t"})
  End Sub
End Module