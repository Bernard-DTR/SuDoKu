'-------------------------------------------------------------------------------
' Création  Mardi 14/05/2024  
'
' La mise en place d'une nouvelle stratégie entraîne la mise place de différents éléments:
' Création d'un module P19_Strategy_SKy.vb qui comporte les calculs de la stratégie
'                                          les calculs sont stockés dans Strategy_Rslt(,) As String
' Un bouton K dans la barre d'outils qui lance G4_Grid_Stratégie_SKy() pour afficher Cellules, Candidats et Aide Graphique
' Un complément des Plcy comportant les stratégies 
'    (SKy devrait précéder Unq)
'
' Exemple proposé:
' 697.....2..1972.63..3..679.912...6.737426.95.8657.9.241486932757.9.24..6..68.7..9
' ..1.28759.879.5132952173486.2.7..34....5..27.714832695....9.817.78.5196319..87524
'-------------------------------------------------------------------------------
' Stratégie des Skyscraper                         
' Préfixe       SKy
' Lettre        K
' Documentation issue de http://sudopedia.enjoysudoku.com 
'  
'-------------------------------------------------------------------------------   
' Le module comporte également la stratégie Kite Cerf-Volant
' S/Stratégie   SKyH1 H2 H3 H4
'               SKyV1 V2 V3 V4
' 
'               Kyt   Mise en place le 24/05/2024
'               EyR   Mise en place ds Empty Rectangle le 05/06/2024
'-------------------------------------------------------------------------------
Friend Module P19_Strategy_SKy
  Public Class SKy_Cls 'Classe structurant les résultats de la stratégie Skyscraper
    Public Property Typologie As String              ' La Typologie H/V 
    Public Property Candidat As String               ' Le Candidat
    Public Property Cellule_1 As Integer             ' Première Cellule
    Public Property Cellule_2 As Integer             ' Deuxième Cellule

    Sub New(New_Typologie As String,
            New_Candidat As String,
            New_Cellule_1 As Integer,
            New_cellule_2 As Integer)
      Typologie = New_Typologie
      Candidat = New_Candidat
      Cellule_1 = New_Cellule_1
      Cellule_2 = New_cellule_2
    End Sub
  End Class

  Sub Strategy_SKy_Display(SKy_List As List(Of SKy_Cls))
    Jrn_Add(, {"Liste de Sky_List : Nombre de Postes " & CStr(SKy_List.Count)})
    Dim S As String
    For i As Integer = 0 To SKy_List.Count - 1
      S = CStr(i).PadRight(5) & "  " & SKy_List.Item(i).Typologie & " " &
                                       SKy_List.Item(i).Candidat & " " &
                                       U_Coord(SKy_List.Item(i).Cellule_1) & " " &
                                       U_Coord(SKy_List.Item(i).Cellule_2) & " "
      Jrn_Add(, {S})
    Next i
  End Sub
  '
  '////// Skyscraper Gratte-Ciel////////////////////////////////////////////////////////////////////////////////////////////////
  '
  '        Extension / Résolution
  '        Mnu08_Résolution_Click(sender As Object, e As EventArgs) Handles Mnu08_Résolution.Click
  '        -->   Pzzl_Slv_Interactif()
  '              --> Pzzl_Slv("*All", Prd, Strategy_Rslt)

  Function Strategy_SKy(U_temp(,) As String) As String(,)
    Dim Stratégie As String = "SKy"
    Dim Sous_Stratégie As String = "SKy"
    Dim Strategy_Rslt(99, 0) As String 'initialisé  
    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des SKy
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 
    'Recherche d'une typologie Horizontale et ensuite Verticale

    Strategy_SKy_Gratte_Ciel(U_temp, Sous_Stratégie & "H", Strategy_Rslt)
    Strategy_SKy_Gratte_Ciel(U_temp, Sous_Stratégie & "V", Strategy_Rslt)
    Strategy_SKy_Kyte_Cerf_Volant(U_temp, "Kyt", Strategy_Rslt)
    Strategy_SKy_EyR_Rectangle_Vide(U_temp, "EyR", Strategy_Rslt)
    Return Strategy_Rslt
  End Function

  Sub Strategy_SKy_Gratte_Ciel(U_temp(,) As String, ByRef Sous_stratégie As String, ByRef Strategy_Rslt(,) As String)
    'Skyscraper _ Gratte-Ciel
    Dim Stratégie As String = "SKy"
    Dim Typologie As String = Sous_stratégie.Substring(3, 1)
    Dim Typologie_Type As String
    Dim SKy1_List As New List(Of SKy_Cls)
    Dim Grp(0 To 8) As Integer
    Dim Cellule, Cellule_1, Cellule_2 As Integer
    Dim n As Integer
    For cd As Integer = 1 To 9
      'Recherche H/V de 2 Cellules comportant le même candidat
      For rc As Integer = 0 To 8
        Select Case Typologie
          Case "H" : Grp = U_9CelRow(rc)
          Case "V" : Grp = U_9CelCol(rc)
        End Select
        n = 0
        For j As Integer = 0 To 8
          Cellule = Grp(j)
          If U_temp(Cellule, 3).Contains(CStr(cd)) Then
            n += 1
            If n = 1 Then Cellule_1 = Cellule
            If n = 2 Then Cellule_2 = Cellule
          End If
        Next j
        'Seul le cas ou le candidat cd est présent 2 fois est enregistré
        If n = 2 Then SKy1_List.Add(New SKy_Cls(Typologie, CStr(cd), Cellule_1, Cellule_2))
      Next rc
    Next cd '/For cd = 1 To 9
    'Strategy_SKy_Display(SKy1_List)

    'Il faut à présent 2 lignes ou 2 colonnes pour le même candidat
    Dim Candidat_1, Candidat_2 As String
    Dim Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2 As Integer
    For i As Integer = 0 To SKy1_List.Count - 1
      Candidat_1 = SKy1_List.Item(i).Candidat
      For j As Integer = i + 1 To SKy1_List.Count - 1
        Candidat_2 = SKy1_List.Item(j).Candidat
        If Candidat_1 <> Candidat_2 Then Continue For
        Select Case Typologie
          Case "H" 'Il faut une colonne identique
            If U_Col(SKy1_List.Item(i).Cellule_1) = U_Col(SKy1_List.Item(j).Cellule_1) Then
              Typologie_Type = "1"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_1
              cel_Base_2 = SKy1_List.Item(j).Cellule_1
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_2
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_2
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
            If U_Col(SKy1_List.Item(i).Cellule_1) = U_Col(SKy1_List.Item(j).Cellule_2) Then
              Typologie_Type = "2"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_1
              cel_Base_2 = SKy1_List.Item(j).Cellule_2
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_2
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_1
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
            If U_Col(SKy1_List.Item(i).Cellule_2) = U_Col(SKy1_List.Item(j).Cellule_1) Then
              Typologie_Type = "3"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_2
              cel_Base_2 = SKy1_List.Item(j).Cellule_1
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_1
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_2
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
            If U_Col(SKy1_List.Item(i).Cellule_2) = U_Col(SKy1_List.Item(j).Cellule_2) Then
              Typologie_Type = "4"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_2
              cel_Base_2 = SKy1_List.Item(j).Cellule_2
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_1
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_1
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If

          Case "V" 'Il faut une rangée identique
            If U_Row(SKy1_List.Item(i).Cellule_1) = U_Row(SKy1_List.Item(j).Cellule_1) Then
              Typologie_Type = "1"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_1
              cel_Base_2 = SKy1_List.Item(j).Cellule_1
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_2
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_2
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
            If U_Row(SKy1_List.Item(i).Cellule_1) = U_Row(SKy1_List.Item(j).Cellule_2) Then
              Typologie_Type = "2"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_1
              cel_Base_2 = SKy1_List.Item(j).Cellule_2
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_2
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_1
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
            If U_Row(SKy1_List.Item(i).Cellule_2) = U_Row(SKy1_List.Item(j).Cellule_1) Then
              Typologie_Type = "3"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_2
              cel_Base_2 = SKy1_List.Item(j).Cellule_1
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_1
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_2
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
            If U_Row(SKy1_List.Item(i).Cellule_2) = U_Row(SKy1_List.Item(j).Cellule_2) Then
              Typologie_Type = "4"
              Cel_Base_1 = SKy1_List.Item(i).Cellule_2
              cel_Base_2 = SKy1_List.Item(j).Cellule_2
              Cel_Fins_1 = SKy1_List.Item(i).Cellule_1
              Cel_Fins_2 = SKy1_List.Item(j).Cellule_1
              Strategy_SKy_Gratte_Ciel_Add(U_temp, Stratégie, Sous_stratégie & Typologie_Type, Strategy_Rslt, Candidat_1, Cel_Base_1, cel_Base_2, Cel_Fins_1, Cel_Fins_2)
            End If
        End Select
      Next j 'For j = Cel_i + 1 To SKy1_List.Count - 1
    Next i   '/For Cel_i = 0 To SKy1_List.Count - 1
  End Sub

  Sub Strategy_SKy_Gratte_Ciel_Add(U_temp(,) As String, Stratégie As String, ByRef Sous_stratégie As String,
                          ByRef Strategy_Rslt(,) As String,
                          Candidat As String,
                          Cel_Base_1 As Integer, Cel_Base_2 As Integer, Cel_Fins_1 As Integer, Cel_Fins_2 As Integer)
    'Recherche des candidats à exclure
    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
    Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice

    For i As Integer = 0 To 80
      If U_temp(i, 2) <> " " Then Continue For
      If i = Cel_Base_1 Or i = Cel_Base_2 Or i = Cel_Fins_1 Or i = Cel_Fins_2 Then Continue For
      If U_temp(i, 3).Contains(Candidat) = False Then Continue For
      If (U_Reg(i) = U_Reg(Cel_Fins_1) Or U_Row(i) = U_Row(Cel_Fins_1) Or U_Col(i) = U_Col(Cel_Fins_1)) _
      And (U_Reg(i) = U_Reg(Cel_Fins_2) Or U_Row(i) = U_Row(Cel_Fins_2) Or U_Col(i) = U_Col(Cel_Fins_2)) Then
        Exc45_n += 1 : Exc45(Exc45_n) = CStr(i)
      End If
    Next i
    If Exc45_n <> -1 Then
      Cel45(0) = CStr(Cel_Base_1)
      Cel45(1) = CStr(Cel_Base_2)
      Cel45(2) = CStr(Cel_Fins_1)
      Cel45(3) = CStr(Cel_Fins_2)
      Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_stratégie, "X", "-1", Candidat, Cel45, Exc45)
    End If
  End Sub

  '
  '////// Kyte Cerf-Volant /////////////////////////////////////////////////////////////////////////////////////////////////////
  '
  ' Exemple proposé:
  ' .81.2.6...42.6..89.568..24.69314275842835791617568932451..3689223...846.86.2.....
  ' 3617..295842395671.5.2614831.8526.34625....18.341..5264..61.85258...2167216857349

  Public Class Kyt_Cls 'Classe structurant les résultats de la stratégie Kyte
    Public Property Candidat As String                ' Le Candidat
    Public Property Typologie As String               ' H1V1, H1V2, H2V1, H2V2, (V1H1, V1H2, V2H1, V2H2)
    Public Property Cellule_H1 As Integer             ' Première Cellule Horizontale
    Public Property Cellule_H2 As Integer             ' Deuxième Cellule Horizontale
    Public Property Cellule_V1 As Integer             ' Première Cellule Verticale
    Public Property Cellule_V2 As Integer             ' Deuxième Cellule Verticale
    Sub New(New_Candidat As String,
            New_typologie As String,
            New_Cellule_H1 As Integer,
            New_cellule_H2 As Integer,
            New_Cellule_V1 As Integer,
            New_cellule_V2 As Integer)
      Candidat = New_Candidat
      Typologie = New_typologie
      Cellule_H1 = New_Cellule_H1
      Cellule_H2 = New_cellule_H2
      Cellule_V1 = New_Cellule_V1
      Cellule_V2 = New_cellule_V2
    End Sub
  End Class
  Sub Strategy_Kyt_Display(Kyt_list As List(Of Kyt_Cls))
    Jrn_Add(, {"Liste de Kyt_list : Nombre de Postes " & CStr(Kyt_list.Count)})
    Dim S As String
    S = "i    Cdd " & "Typo " & "H1    " & "H2    " & "V1    " & "V2    "
    Jrn_Add(, {S})
    For i As Integer = 0 To Kyt_list.Count - 1
      S = CStr(i).PadRight(5) & "  " & Kyt_list.Item(i).Candidat & " " &
                                       Kyt_list.Item(i).Typologie & " " &
                                       U_Coord(Kyt_list.Item(i).Cellule_H1) & " " &
                                       U_Coord(Kyt_list.Item(i).Cellule_H2) & " " &
                                       U_Coord(Kyt_list.Item(i).Cellule_V1) & " " &
                                       U_Coord(Kyt_list.Item(i).Cellule_V2) & " "
      Jrn_Add(, {S})
    Next i
  End Sub
  Sub Strategy_SKy_Kyte_Cerf_Volant(U_temp(,) As String, ByRef Sous_stratégie As String, ByRef Strategy_Rslt(,) As String)
    Dim Typologie As String
    Dim Kyt_List As New List(Of Kyt_Cls)
    'La Stratégie est SKy, La S/Stratégie est Kyt Kyte _ Cerf-Volant
    Dim Grp() As Integer
    Dim Cellule, Cellule_H1, Cellule_H2, Cellule_V1, Cellule_V2 As Integer
    Dim n As Integer
    For cd As Integer = 1 To 9
      'Recherche Horizontale de 2 Cellules comportant le même candidat
      For h As Integer = 0 To 8
        Grp = U_9CelRow(h)
        n = 0
        Cellule_H1 = -1 : Cellule_H2 = -1
        For j As Integer = 0 To 8
          Cellule = Grp(j)
          If U_temp(Cellule, 3).Contains(CStr(cd)) Then
            n += 1
            If n = 1 Then Cellule_H1 = Cellule
            If n = 2 Then Cellule_H2 = Cellule
          End If
        Next j
        'Seul le cas ou le candidat cd est présent 2 fois est enregistré
        If n <> 2 Then Continue For
        'Cellule_H1 et Cellule_H2 occupent 2 régions différentes
        If U_Reg(Cellule_H1) = U_Reg(Cellule_H2) Then Continue For
        'Recherche Verticale de 2 cellules comportant le même candidat
        For v As Integer = 0 To 8
          Grp = U_9CelCol(v)
          n = 0
          Cellule_V1 = -1 : Cellule_V2 = -1
          For j As Integer = 0 To 8
            Cellule = Grp(j)
            If U_temp(Cellule, 3).Contains(CStr(cd)) Then
              n += 1
              If n = 1 Then Cellule_V1 = Cellule
              If n = 2 Then Cellule_V2 = Cellule
            End If
          Next j
          'Seul le cas ou le candidat cd est présent 2 fois est enregistré
          If n <> 2 Then Continue For
          'Il faut 4 cellules différentes
          If Cellule_H1 = Cellule_V1 Or Cellule_H1 = Cellule_V2 Then Continue For
          If Cellule_H2 = Cellule_V1 Or Cellule_H2 = Cellule_V2 Then Continue For

          'Cellule_V1 et Cellule_V2 occupent 2 régions différentes
          If U_Reg(Cellule_V1) = U_Reg(Cellule_V2) Then Continue For

          ' Il faut que H1 ou H2 soit dans la même région que V1 ou V2
          '      ou que V1 ou V2 soit dans la même région que H1 ou H2
          Typologie = "#"
          If U_Reg(Cellule_H1) = U_Reg(Cellule_V1) Then Typologie = "H1V1"
          If U_Reg(Cellule_H1) = U_Reg(Cellule_V2) Then Typologie = "H1V2"
          If U_Reg(Cellule_H2) = U_Reg(Cellule_V1) Then Typologie = "H2V1"
          If U_Reg(Cellule_H2) = U_Reg(Cellule_V2) Then Typologie = "H2V2"

          If Typologie <> "#" Then
            Kyt_List.Add(New Kyt_Cls(CStr(cd), Typologie, Cellule_H1, Cellule_H2, Cellule_V1, Cellule_V2))
          End If
        Next v '/For v = 0 To 8
      Next h   '/For h = 0 To 8
    Next cd    '/For cd = 1 To 9
    'Strategy_Kyt_Display(Kyt_List)

    'Traitement des lignes de la List et ajout dans Strategy_Rslt
    Strategy_SKy_Kyte_Cerf_Volant_Add(U_temp, Sous_stratégie, Strategy_Rslt, Kyt_List)

  End Sub

  Sub Strategy_SKy_Kyte_Cerf_Volant_Add(U_temp(,) As String, ByRef Sous_stratégie As String, ByRef Strategy_Rslt(,) As String, Kyt_list As List(Of Kyt_Cls))
    'Traitement des lignes
    Dim Stratégie As String = "SKy"
    For l As Integer = 0 To Kyt_list.Count - 1
      With Kyt_list.Item(l)
        'Recherche des candidats à exclure
        Dim Candidat As String = .Candidat
        Dim Cellule_1, Cellule_2 As Integer
        'Pour les 81 cellules 
        For Cel_i As Integer = 0 To 80
          Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
          Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k
          Dim Exc45_n As Integer = -1 ' Nombre de Cellules Exc45 ou indice
          'Cel_i ne doit pa être une cellule du cerf-volant, ni être remplie et doit comporter le candidat
          If Cel_i = .Cellule_H1 Or Cel_i = .Cellule_H2 Or Cel_i = .Cellule_V1 Or Cel_i = .Cellule_V2 Then Continue For
          If U_temp(Cel_i, 2) <> " " Then Continue For
          If U_temp(Cel_i, 3).Contains(Candidat) = False Then Continue For
          Cellule_1 = -1
          Cellule_2 = -1

          Select Case .Typologie
            Case "H1V1" : Cellule_1 = .Cellule_H2 : Cellule_2 = .Cellule_V2
              '3617..295842395671.5.2614831.8526.34625....18.341..5264..61.85258...2167216857349
            Case "H1V2" : Cellule_1 = .Cellule_H2 : Cellule_2 = .Cellule_V1
              '513486279978312564264..5813.516...27.962..1.57325.16.812.854..6645.23.8138.16.452
            Case "H2V1" : Cellule_1 = .Cellule_H1 : Cellule_2 = .Cellule_V2
              '254.61.8318.32.5466..458.218.61.52375.1..269.72...615.3185..462465213879972684315
            Case "H2V2" : Cellule_1 = .Cellule_H1 : Cellule_2 = .Cellule_V1
              '.81.2.6...42.6..89.568..24.69314275842835791617568932451..3689223...846.86.2.....
          End Select
          If Cellule_1 = -1 Or Cellule_2 = -1 Then Continue For
          ' Cellule_1 et Cellule_2 sont dans la même ligne et dans la même colonne
          If U_Col(Cel_i) = U_Col(Cellule_1) And U_Row(Cel_i) = U_Row(Cellule_2) Then
            Exc45_n += 1 : Exc45(Exc45_n) = CStr(Cel_i)
          End If
          If Exc45_n <> -1 Then
            Cel45(0) = CStr(.Cellule_H1)
            Cel45(1) = CStr(.Cellule_H2)
            Cel45(2) = CStr(.Cellule_V1)
            Cel45(3) = CStr(.Cellule_V2)
            Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_stratégie & .Typologie, "X", "-1", Candidat, Cel45, Exc45)
          End If
        Next Cel_i '/For Cel_i = 0 To 80
      End With
    Next l
  End Sub



  '
  '////// EyR Empty Rectangle ////////////////////////////////////////////////////////////////////////////////////////////////////
  '
  ' Exemple proposé:
  ' 7...56.381.842.............5..3......4..8.7....1...24...3.........1....5.5...7.6.
  'Après résolution
  '7249561381684235979357186245..3..81..4..8175..81.7.24..13....72...1...85.5...7.61

  ' ..86........759...6..1...9.......83.9.6.....5..24.5.....5..438.3.18.2..6......2..
  'Après résolution
  '598643..2..37596486741285934572..83.9.63.7425.324.5.6...59.438.341872956..953.2.4

  Sub Strategy_SKy_EyR_Rectangle_Vide(U_temp(,) As String, Sous_stratégie As String, ByRef Strategy_Rslt(,) As String)
    'La Stratégie est SKy, La S/Stratégie est EyR Empty Rectangle
    For cd As Integer = 1 To 9
      For Région As Integer = 0 To 8
        Strategy_SKy_EyR_Rectangle_Vide_Calcul(U_temp, Sous_stratégie, Strategy_Rslt, cd, Région)
      Next Région
    Next cd
  End Sub

  Sub Strategy_SKy_EyR_Rectangle_Vide_Calcul(U_temp(,) As String, Sous_stratégie As String, ByRef Strategy_Rslt(,) As String, Cd As Integer, Région As Integer)
    Dim Grp() As Integer
    Dim Cellule As Integer
    Dim Cellule_ER As Integer
    Dim Row1, Row2, Row3 As Integer
    Dim Col1, Col2, Col3 As Integer
    Dim Cell_Rect_ER As Integer = -1
    Dim Typologie As String
    Dim n As Integer
    Dim Cellule_A, Cellule_B As Integer
    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k

    ' 1 On décompte le candidat dans les 3 Rows et les 3 Cols de la région
    Row1 = 0 : Row2 = 0 : Row3 = 0 : Col1 = 0 : Col2 = 0 : Col3 = 0
    Grp = U_9CelReg(Région)
    For j As Integer = 0 To 8
      Cellule = Grp(j)
      If U_temp(Cellule, 3).Contains(CStr(Cd)) = True Then
        Select Case j
          Case 0 : Row1 += 1 : Col1 += 1
          Case 1 : Row1 += 1 : Col2 += 1
          Case 2 : Row1 += 1 : Col3 += 1
          Case 3 : Row2 += 1 : Col1 += 1
          Case 4 : Row2 += 1 : Col2 += 1
          Case 5 : Row2 += 1 : Col3 += 1
          Case 6 : Row3 += 1 : Col1 += 1
          Case 7 : Row3 += 1 : Col2 += 1
          Case 8 : Row3 += 1 : Col3 += 1
        End Select
      End If
    Next j
    ' Parce qu'un candidat qui se trouve dans une ligne se trouve de facto dans la colonne et vice-versa
    If Row1 = 1 Then Row1 = 0
    If Row2 = 1 Then Row2 = 0
    If Row3 = 1 Then Row3 = 0
    If Col1 = 1 Then Col1 = 0
    If Col2 = 1 Then Col2 = 0
    If Col3 = 1 Then Col3 = 0

    ' 2 Le candidat doit se trouver dans une seule ligne ET une seule colonne
    '   La typologie RxCy signifie le numéro de la Rangée et le numéro de la colonne de la Région concernée
    '   R1C1 : Rangée 1 et Colonne 1 de la région (c'est la cellule 0 en haut à droite)
    Typologie = "#"
    If Row1 <> 0 And Row2 = 0 And Row3 = 0 And Col1 <> 0 And Col2 = 0 And Col3 = 0 Then Typologie = "R1C1"
    If Row1 = 0 And Row2 <> 0 And Row3 = 0 And Col1 <> 0 And Col2 = 0 And Col3 = 0 Then Typologie = "R2C1"
    If Row1 = 0 And Row2 = 0 And Row3 <> 0 And Col1 <> 0 And Col2 = 0 And Col3 = 0 Then Typologie = "R3C1"
    If Row1 <> 0 And Row2 = 0 And Row3 = 0 And Col1 = 0 And Col2 <> 0 And Col3 = 0 Then Typologie = "R1C2"
    If Row1 = 0 And Row2 <> 0 And Row3 = 0 And Col1 = 0 And Col2 <> 0 And Col3 = 0 Then Typologie = "R2C2"
    If Row1 = 0 And Row2 = 0 And Row3 <> 0 And Col1 = 0 And Col2 <> 0 And Col3 = 0 Then Typologie = "R3C2"
    If Row1 <> 0 And Row2 = 0 And Row3 = 0 And Col1 = 0 And Col2 = 0 And Col3 <> 0 Then Typologie = "R1C3"
    If Row1 = 0 And Row2 <> 0 And Row3 = 0 And Col1 = 0 And Col2 = 0 And Col3 <> 0 Then Typologie = "R2C3"
    If Row1 = 0 And Row2 = 0 And Row3 <> 0 And Col1 = 0 And Col2 = 0 And Col3 <> 0 Then Typologie = "R3C3"

    ' 3 Un ER Empty Rectangle est trouvé
    If Typologie = "#" Then Exit Sub
    Select Case Typologie
      Case "R1C1" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 0)
      Case "R2C1" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 3)
      Case "R3C1" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 6)
      Case "R1C2" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 1)
      Case "R2C2" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 4)
      Case "R3C2" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 7)
      Case "R1C3" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 2)
      Case "R2C3" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 5)
      Case "R3C3" : Cell_Rect_ER = Wh_NumérodelaCellule_RegNreg(Région, 8)
    End Select

    ' 4 On recherche une paire conjuguée :
    '   une ligne ou une colonne où 2 cellules comporte 2 candidats dont cd

    '   Recherche d'une ligne avec 2 candidats dont LE CANDIDAT
    For i As Integer = 0 To 8
      n = 0 : Cellule_A = -1 : Cellule_B = -1
      Grp = Wh_9CellulesRow_Row(i)
      For j As Integer = 0 To 8
        Cellule = Grp(j)
        If U_Reg(Cellule) <> Région AndAlso Wh_Cell_Nb_Candidats(U_temp, Cellule) >= 2 AndAlso U_temp(Cellule, 3).Contains(CStr(Cd)) Then
          n += 1
          If n = 1 Then Cellule_A = Cellule
          If n = 2 Then Cellule_B = Cellule
        End If
      Next j
      If n <> 2 Then Continue For

      ' Il faut qu'une des 2 cellules se trouve sur la colonne de Cell_Rect_ER
      If U_Row(Cell_Rect_ER) = U_Row(Cellule_A) Or U_Col(Cell_Rect_ER) = U_Col(Cellule_A) Then
        Cellule_ER = Wh_Cellule_RowCol(U_Row(Cell_Rect_ER), U_Col(Cellule_B))
        ' Cellule_ER ne doit pas être dans la Région 
        If Cellule_ER <> Cellule_B And U_Reg(Cellule_ER) <> U_Reg(Cell_Rect_ER) And U_temp(Cellule_ER, 3).Contains(CStr(Cd)) Then
          Cel45(0) = CStr(Cell_Rect_ER)
          Cel45(1) = CStr(Cellule_A)
          Cel45(2) = CStr(Cellule_B)
          Exc45(0) = CStr(Cellule_ER)
          Strategy_Rslt_Add(Strategy_Rslt, "SKy", Sous_stratégie, "X", "-1", CStr(Cd), Cel45, Exc45)
        End If
      End If

      If U_Row(Cell_Rect_ER) = U_Row(Cellule_B) Or U_Col(Cell_Rect_ER) = U_Col(Cellule_B) Then
        Cellule_ER = Wh_Cellule_RowCol(U_Row(Cell_Rect_ER), U_Col(Cellule_A))
        If Cellule_ER <> Cellule_A And U_Reg(Cellule_ER) <> U_Reg(Cell_Rect_ER) And U_temp(Cellule_ER, 3).Contains(CStr(Cd)) Then
          Cel45(0) = CStr(Cell_Rect_ER)
          Cel45(1) = CStr(Cellule_B)
          Cel45(2) = CStr(Cellule_A)
          Exc45(0) = CStr(Cellule_ER)
          Strategy_Rslt_Add(Strategy_Rslt, "SKy", Sous_stratégie, "X", "-1", CStr(Cd), Cel45, Exc45)
        End If
      End If
    Next i

    '   Recherche d'une colonne avec 2 candidats dont LE CANDIDAT
    For i As Integer = 0 To 8
      n = 0 : Cellule_A = -1 : Cellule_B = -1
      Grp = Wh_9CellulesColumn_Col(i)
      For j As Integer = 0 To 8
        Cellule = Grp(j)
        If U_Reg(Cellule) <> Région AndAlso Wh_Cell_Nb_Candidats(U_temp, Cellule) >= 2 AndAlso U_temp(Cellule, 3).Contains(CStr(Cd)) Then
          n += 1
          If n = 1 Then Cellule_A = Cellule
          If n = 2 Then Cellule_B = Cellule
        End If
      Next j
      If n <> 2 Then Continue For

      ' Il faut qu'une des 2 cellules se trouve sur la colonne de Cell_Rect_ER
      If U_Row(Cell_Rect_ER) = U_Row(Cellule_A) Or U_Col(Cell_Rect_ER) = U_Col(Cellule_A) Then
        Cellule_ER = Wh_Cellule_RowCol(U_Row(Cellule_B), U_Col(Cell_Rect_ER))
        If Cellule_ER <> Cellule_B And U_Reg(Cellule_ER) <> U_Reg(Cell_Rect_ER) And U_temp(Cellule_ER, 3).Contains(CStr(Cd)) Then
          Cel45(0) = CStr(Cell_Rect_ER)
          Cel45(1) = CStr(Cellule_A)
          Cel45(2) = CStr(Cellule_B)
          Exc45(0) = CStr(Cellule_ER)
          Strategy_Rslt_Add(Strategy_Rslt, "SKy", Sous_stratégie, "X", "-1", CStr(Cd), Cel45, Exc45)
        End If
      End If
      If U_Row(Cell_Rect_ER) = U_Row(Cellule_B) Or U_Col(Cell_Rect_ER) = U_Col(Cellule_B) Then
        Cellule_ER = Wh_Cellule_RowCol(U_Row(Cellule_A), U_Col(Cell_Rect_ER))
        If Cellule_ER <> Cellule_A And U_Reg(Cellule_ER) <> U_Reg(Cell_Rect_ER) And U_temp(Cellule_ER, 3).Contains(CStr(Cd)) Then
          Cel45(0) = CStr(Cell_Rect_ER)
          Cel45(1) = CStr(Cellule_B)
          Cel45(2) = CStr(Cellule_A)
          Exc45(0) = CStr(Cellule_ER)
          Strategy_Rslt_Add(Strategy_Rslt, "SKy", Sous_stratégie, "X", "-1", CStr(Cd), Cel45, Exc45)
        End If
      End If
    Next i
  End Sub


End Module