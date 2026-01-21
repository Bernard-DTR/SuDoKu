Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------
' Stratégie des XYwing
' Préfixe       XYw
' Stratégie     XYw
' S/Stratégie   XYw
'
' Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Aa1",                "R", "-1", Candidat, Cel45, Exc45)
' Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Ba1",                "C", "-1", Candidat, Cel45, Exc45)
' Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Cb" & CStr(Exc45_n), "L", "-1", Candidat, Cel45, Exc45)
' Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Db" & CStr(Exc45_n), "L", "-1", Candidat, Cel45, Exc45)
' Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Eb" & CStr(Exc45_n), "L", "-1", Candidat, Cel45, Exc45)
' Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Fb" & CStr(Exc45_n), "C", "-1", Candidat, Cel45, Exc45)
'
' S/ Stratégie, premier caractère  
' Select Case Mid$(Strategy_Rslt(2, Ligne), 1, 1)
'   Case "C"
'     G4_MdC_Row_Col_Box("Row", U_Row(CInt(Strategy_Rslt(10, Ligne))))
'     G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
'   Case "F"
'     G4_MdC_Row_Col_Box("Col", U_Col(CInt(Strategy_Rslt(10, Ligne))))
'     G4_MdC_Row_Col_Box("Box", U_Reg(CInt(Strategy_Rslt(10, Ligne))))
'   Case Else
'     G4_MdC_Row_Col_Box("Row", U_Row(CInt(Strategy_Rslt(55, Ligne))))
'     G4_MdC_Row_Col_Box("Col", U_Col(CInt(Strategy_Rslt(55, Ligne))))
' End Select
'-------------------------------------------------------------------------------

Friend Module P15_Strategy_XYw
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Function Strategy_XYw(U_temp(,) As String) As String(,)
    Dim Cdd_X As String = ""
    Dim Cdd_Y As String = ""
    Dim n As Integer = -1
    Dim Stratégie As String = "XYw"
    Dim Sous_Stratégie As String = "XYw"

    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des Cbl
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Cel45_n As Integer = -1 ' Nombre de Cellules Cel45 ou indice
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice


    Dim XY1_t(5, 81) As String      'Taille maximale 81,  XY1_t_Nb représente le nombre de lignes remplies de 0 à x
    '         0  Numéro du poste    'Comporte les cellules à 2 candidats
    '         1  Cellule
    '         2  Cellule en coord
    '         3  Candidats
    '         4  Candidat 1
    '         5  Candidat 2
    'Le nombre de postes 
    Dim XY1_t_Nb As Integer = -1
    For i As Integer = 0 To 5
      For j As Integer = 0 To 80
        XY1_t(i, j) = "" 'Le tableau est string  et initialisé
      Next j
    Next i

    ' 01 Analyse des Cellules à 2 candidats
    '    Les cellules à 2 candidats sont stockées dans XY1_t
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    For i As Integer = 0 To 80
      If Wh_Cell_Nb_Candidats(U_temp, i) = 2 Then
        n += 1
        XY1_t(0, n) = CStr(n)
        XY1_t(1, n) = CStr(i) : XY1_t(2, n) = U_cr(i)
        Dim n1 As Integer = 0 'Premier et deuxième candidat
        For j As Integer = 0 To 8
          If Mid$(U_temp(i, 3), j + 1, 1) <> " " Then
            n1 += 1
            If n1 = 1 Then Cdd_X = Mid$(U_temp(i, 3), j + 1, 1)
            If n1 = 2 Then Cdd_Y = Mid$(U_temp(i, 3), j + 1, 1)
          End If
        Next j
        XY1_t(3, n) = U_temp(i, 3) : XY1_t(4, n) = Cdd_X : XY1_t(5, n) = Cdd_Y
        XY1_t_Nb = n  'Détermination de XY_t_Ubound, dernier poste rempli
      End If
    Next i
    If XY1_t_Nb = -1 Then GoTo Strategy_XYw_End

    ' 03 on a XY, on recherche XZ et YZ
    '    Les combinaisons XY, XZ et YZ sont stockées dans XY2_t
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Dim XY2_t(15, 0) As String      'Taille inconnue, agrandie à chaque usage
    '          0  Numéro du poste   'Comporte les associations XY,XZ et YZ
    '      1 à 3  XY
    '      4 à 6  XZ
    '      7 à 9  YZ
    'Le nombre de postes 
    Dim XY2_t_Nb As Integer = -1

    For i As Integer = 0 To XY1_t_Nb
      Cdd_X = XY1_t(4, i)
      Cdd_Y = XY1_t(5, i)
      For cdd As Integer = 1 To 9
        If cdd = CInt(Cdd_X) Or cdd = CInt(Cdd_Y) Then Continue For
        Dim Cdd_XZ As String = Cnddts_Blancs
        Dim Cdd_YZ As String = Cnddts_Blancs
        Mid$(Cdd_XZ, CInt(Cdd_X), 1) = Cdd_X
        Mid$(Cdd_YZ, CInt(Cdd_Y), 1) = Cdd_Y
        Mid$(Cdd_XZ, cdd, 1) = CStr(cdd)
        Mid$(Cdd_YZ, cdd, 1) = CStr(cdd)

        For i_xz As Integer = 0 To XY1_t_Nb
          If Cdd_XZ = XY1_t(3, i_xz) Then
            For i_yz As Integer = 0 To XY1_t_Nb
              If Cdd_YZ = XY1_t(3, i_yz) Then
                ' Combinaison XY, XZ et YZ
                XY2_t_Nb += 1
                If XY2_t_Nb > UBound(XY2_t, 2) Then ReDim Preserve XY2_t(15, XY2_t_Nb + 1)
                XY2_t(0, XY2_t_Nb) = CStr(XY2_t_Nb)
                XY2_t(1, XY2_t_Nb) = XY1_t(1, i)      'XY
                XY2_t(2, XY2_t_Nb) = XY1_t(2, i)
                XY2_t(3, XY2_t_Nb) = XY1_t(3, i)
                XY2_t(4, XY2_t_Nb) = XY1_t(1, i_xz)   'XZ
                XY2_t(5, XY2_t_Nb) = XY1_t(2, i_xz)
                XY2_t(6, XY2_t_Nb) = XY1_t(3, i_xz)
                XY2_t(7, XY2_t_Nb) = XY1_t(1, i_yz)   'YZ
                XY2_t(8, XY2_t_Nb) = XY1_t(2, i_yz)
                XY2_t(9, XY2_t_Nb) = XY1_t(3, i_yz)
                XY2_t(10, XY2_t_Nb) = CStr(cdd)       'Le candidat
              End If
            Next i_yz
          End If
        Next i_xz
      Next cdd
    Next i
    If XY2_t_Nb = -1 Then GoTo Strategy_XYw_End

    Dim Cel_XY As Integer = -1
    Dim Cel_XZ As Integer = -1
    Dim Cel_YZ As Integer = -1

    ' 05 XY_Wing N° 1 en carré
    Dim Cellule_BD As Integer = -1
    Dim Candidat As String = ""
    For i As Integer = 0 To XY2_t_Nb
      Cel_XY = CInt(XY2_t(1, i))
      Candidat = XY2_t(10, i)
      ' XY    YZ  
      '
      ' XZ    BD Candidat ?
      '
      Cel_XZ = CInt(XY2_t(4, i))
      Cel_YZ = CInt(XY2_t(7, i))
      If (U_Col(Cel_XY) = U_Col(Cel_XZ) And U_Row(Cel_XY) = U_Row(Cel_YZ)) Then
        Cellule_BD = Wh_Cellule_ColRow(U_Col(Cel_YZ), U_Row(Cel_XZ))
        If U_temp(Cellule_BD, 3).Contains(Candidat) Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Cel45(0) = CStr(Cel_XY)
          Cel45(1) = CStr(Cel_XZ)
          Cel45(2) = CStr(Cel_YZ)
          Exc45(0) = CStr(Cellule_BD)
          Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Aa1", "R", "-1", Candidat, Cel45, Exc45)
        End If
      End If


      ' XY    XZ  
      '
      ' YZ    BD Candidat ?
      '
      Cel_XZ = CInt(XY2_t(4, i))
      Cel_YZ = CInt(XY2_t(7, i))
      If (U_Col(Cel_XY) = U_Col(Cel_YZ) And U_Row(Cel_XY) = U_Row(Cel_XZ)) Then
        Cellule_BD = Wh_Cellule_ColRow(U_Col(Cel_XZ), U_Row(Cel_YZ))
        If U_temp(Cellule_BD, 3).Contains(Candidat) Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Cel45(0) = CStr(Cel_XY)
          Cel45(1) = CStr(Cel_YZ)
          Cel45(2) = CStr(Cel_XZ)
          Exc45(0) = CStr(Cellule_BD)
          Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Ba1", "C", "-1", Candidat, Cel45, Exc45)
        End If
      End If
    Next i

    ' 05 XY_Wing N° 2 dans un Rectangle 

    For i As Integer = 0 To XY2_t_Nb
      Cel_XY = CInt(XY2_t(1, i))
      Candidat = XY2_t(10, i)
      Cel_XZ = CInt(XY2_t(4, i))
      Cel_YZ = CInt(XY2_t(7, i))
      '
      '  ? XY  ?  -- -- --  -- YZ --  
      ' -- -- --  -- -- --  -- -- --
      ' YZ -- --  -- -- --   ?  ?  ?
      '
      Dim R As Integer = -1
      Dim Grp(9) As Integer

      Dim Rh_XY As Integer = U_Bh(Cel_XY)
      Dim Rh_XZ As Integer = U_Bh(Cel_XZ)
      Dim Rh_YZ As Integer = U_Bh(Cel_YZ)
      Dim Rv_XY As Integer = U_Bv(Cel_XY)
      Dim Rv_XZ As Integer = U_Bv(Cel_XZ)
      Dim Rv_YZ As Integer = U_Bv(Cel_YZ)
      ' XY, XZ et YZ doivent être dans le même rectangle
      If (Rh_XY = Rh_XZ And Rh_XY = Rh_YZ) Or (Rv_XY = Rv_XZ And Rv_XY = Rv_YZ) Then

        'Cas 1 Même ligne et même région
        If (U_Row(Cel_XY) = U_Row(Cel_XZ)) And
           (U_Reg(Cel_XY) <> U_Reg(Cel_XZ)) And
           (U_Reg(Cel_XY) = U_Reg(Cel_YZ)) And
           (U_Row(Cel_XY) <> U_Row(Cel_YZ)) Then
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1 ' Nombre de Cellules Exc45 ou indice


          'Périmètre: Région de XY, Ligne XY, sauf XY
          '         : Région de XZ, Ligne YZ
          R = U_Reg(Cel_XY) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If ((U_Row(Cel_R) = U_Row(Cel_XY)) And (Cel_R <> Cel_XY)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          R = U_Reg(Cel_XZ) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If (U_Row(Cel_R) = U_Row(Cel_YZ)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          If Exc45_n >= 0 Then
            Cel45(0) = CStr(Cel_XY)
            Cel45(1) = CStr(Cel_XZ)
            Cel45(2) = CStr(Cel_YZ)
            Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Cb" & CStr(Exc45_n), "L", "-1", Candidat, Cel45, Exc45)
          End If
        End If

        'Cas 2 Même ligne et même région
        If (U_Row(Cel_XY) = U_Row(Cel_YZ)) And
           (U_Reg(Cel_XY) <> U_Reg(Cel_YZ)) And
           (U_Reg(Cel_XY) = U_Reg(Cel_XZ)) And
           (U_Row(Cel_XY) <> U_Row(Cel_XZ)) Then

          'Périmètre: Région de XY, Ligne XY, sauf XY
          '         : Région de YZ, Ligne XZ
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1 ' Nombre de Cellules Exc45 ou indice
          R = U_Reg(Cel_XY) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If ((U_Row(Cel_R) = U_Row(Cel_XY)) And (Cel_R <> Cel_XY)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          R = U_Reg(Cel_YZ) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If (U_Row(Cel_R) = U_Row(Cel_XZ)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          If Exc45_n >= 0 Then
            Cel45(0) = CStr(Cel_XY)
            Cel45(1) = CStr(Cel_YZ)
            Cel45(2) = CStr(Cel_XZ)
            Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Db" & CStr(Exc45_n), "L", "-1", Candidat, Cel45, Exc45)
          End If

        End If

        'Cas 3 Même Colonne et même région
        If (U_Col(Cel_XY) = U_Col(Cel_XZ)) And
           (U_Reg(Cel_XY) <> U_Reg(Cel_XZ)) And
           (U_Reg(Cel_XY) = U_Reg(Cel_YZ)) And
           (U_Col(Cel_XY) <> U_Col(Cel_YZ)) Then

          'Périmètre: Région de XY, Colonne XY, sauf XY
          '         : Région de XZ, Colonne YZ
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1 ' Nombre de Cellules Exc45 ou indice
          R = U_Reg(Cel_XY) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If ((U_Col(Cel_R) = U_Col(Cel_XY)) And (Cel_R <> Cel_XY)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          R = U_Reg(Cel_XZ) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If (U_Col(Cel_R) = U_Col(Cel_YZ)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          If Exc45_n >= 0 Then
            Cel45(0) = CStr(Cel_XY)
            Cel45(1) = CStr(Cel_XZ)
            Cel45(2) = CStr(Cel_YZ)
            Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Eb" & CStr(Exc45_n), "L", "-1", Candidat, Cel45, Exc45)
          End If

        End If

        'Cas 4 Même Colonne et même région
        If (U_Col(Cel_XY) = U_Col(Cel_YZ)) And
           (U_Reg(Cel_XY) <> U_Reg(Cel_YZ)) And
           (U_Reg(Cel_XY) = U_Reg(Cel_XZ)) And
           (U_Col(Cel_XY) <> U_Col(Cel_XZ)) Then

          'Périmètre: Région de XY, Colonne XY, sauf XY
          '         : Région de YZ, Colonne XZ
          For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Exc45_n = -1 ' Nombre de Cellules Exc45 ou indice
          R = U_Reg(Cel_XY) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If ((U_Col(Cel_R) = U_Col(Cel_XY)) And (Cel_R <> Cel_XY)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          R = U_Reg(Cel_YZ) : Grp = U_9CelReg(R)
          For k As Integer = 0 To 8
            Dim Cel_R As Integer = Grp(k)
            If (U_Col(Cel_R) = U_Col(Cel_XZ)) Then
              If U_temp(Cel_R, 3).Contains(Candidat) Then
                Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_R)
              End If
            End If
          Next k
          If Exc45_n >= 0 Then
            Cel45(0) = CStr(Cel_XY)
            Cel45(1) = CStr(Cel_YZ)
            Cel45(2) = CStr(Cel_XZ)
            Strategy_Rslt_Add(Strategy_Rslt, "XYw", "Fb" & CStr(Exc45_n), "C", "-1", Candidat, Cel45, Exc45)
          End If
        End If
      End If  '/ même rectangle horizontal ou vertical
    Next i

Strategy_XYw_End:
    Return Strategy_Rslt
  End Function
End Module