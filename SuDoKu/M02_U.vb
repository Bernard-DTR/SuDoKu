Option Strict On
Option Explicit On

Module M02_U
  '-------------------------------------------------------------------------------
  ' Regroupe l'ensemble des traitements concernant le tableau U
  ' 
  '-------------------------------------------------------------------------------

  '-------------------------------------------------------------------------------
  ' Gestion de U
  ' Cellules                           Régions
  ' 
  '  0  1  2| 3  4  5| 6  7  8       '         |        |  
  '  9 10 11|12 13 14|15 16 17       '     0   |    1   |    2
  ' 18 19 20|21 22 23|24 25 26       '         |        |  
  ' --------+--------+--------       ' --------+--------+--------
  ' 27 28 29|30 31 32|33 34 35       '         |        |  
  ' 36 37 38|39 40 41|42 43 44       '     3   |    4   |    5
  ' 45 46 47|48 49 50|51 52 53       '         |        |  
  ' --------+--------+--------       ' --------+--------+--------
  ' 54 55 56|57 58 59|60 61 62       '         |        |  
  ' 63 64 65|66 67 68|69 70 71       '     6   |    7   |    8
  ' 72 73 74|75 76 77|78 79 80       '         |        |   
  ' 
  '-------------------------------------------------------------------------------

  Function U_Coord(Cellule As Integer) As String
    'U_Coord et .Coordonnées ont le même format
    Select Case Cellule
      Case 0 To 80 : U_Coord = "L" & U_Row(Cellule) + 1 & "_" & "C" & U_Col(Cellule) + 1     'format Lx_Cy
      Case Else : U_Coord = "#" & CStr(Cellule).PadLeft(3) & "#"                             'format #-1 ou #82 #
    End Select
    Return U_Coord
  End Function
  Function U_Coord_DB(Cellule As Integer) As String
    'U_Coord_DB (Denis Berthier) propose le Blk et le Numéro du Carré dans le Blk
    Dim Row As Integer = U_Row(Cellule)
    Dim Col As Integer = U_Col(Cellule)
    Dim Blk As Integer = CInt(1 + (3 * Int(Row / 3)) + Int(Col / 3))
    Dim Nuc As Integer = 1 + 3 * ((Row + 3) Mod 3) + (Col + 3) Mod 3
    Select Case Cellule
      Case 0 To 80 : U_Coord_DB = "B" & Blk & "_" & "n" & Nuc                                   'format Bx_ny
      Case Else : U_Coord_DB = "#" & CStr(Cellule).PadLeft(3) & "#"                             'format #-1 ou #82 #
    End Select
    Return U_Coord_DB
  End Function

  'For i As Integer = 0 To 80
  '    Jrn_Add_Orange(CStr(i).PadLeft(3) & " " & U_Coord(i) & " " & U_Coord_DB(i))
  'Next
  '  0 L1_C1 B1_n1
  '  1 L1_C2 B1_n2
  '  2 L1_C3 B1_n3
  '  9 L2_C1 B1_n4
  ' 10 L2_C2 B1_n5
  ' 11 L2_C3 B1_n6
  ' 18 L3_C1 B1_n7
  ' 19 L3_C2 B1_n8
  ' 20 L3_C3 B1_n9
  ' ...
  ' 60 L7_C7 B9_n1
  ' 61 L7_C8 B9_n2
  ' 62 L7_C9 B9_n3
  ' 69 L8_C7 B9_n4
  ' 70 L8_C8 B9_n5
  ' 71 L8_C9 B9_n6
  ' 78 L9_C7 B9_n7
  ' 79 L9_C8 B9_n8
  ' 80 L9_C9 B9_n9

  Public Sub U_Display()
    Jrn_Add("SDK_30010")
    Jrn_Add("SDK_30011")
    Jrn_Add("SDK_30012")
    Jrn_Add("SDK_30013")
    Jrn_Add("SDK_30014")
    Jrn_Add("SDK_30015")
    For i As Integer = 0 To 80
      Dim sc As New Cellule_Cls With {.Numéro = i}

      Dim S As String = "| "
      S &= CStr(i).PadLeft(2)
      S = S & " | " & U_Coord(i)
      S = S & " | " & U(i, 2)
      S = S & " | " & U(i, 3)
      S = S & " | " & U_CddExc(i)
      S = S & " | " & U_Sol(i)
      Select Case U(i, 1)
        Case " " : S &= " "
        Case Else : S &= "i"
      End Select
      Select Case sc.Candidat_Unique
        Case True : S = S & " | " & "T "
        Case False : S = S & " |  " & " "
      End Select
      S = S & "| " & " " & "  | " & "  " & " |" & " " & " " & " |"
      Jrn_Add(, {S})
    Next i
    Jrn_Add("SDK_30015")
    Jrn_Add("SDK_Space")
  End Sub

  Public Sub U_Display2()
    ' à utiliser pour transmettre le jeu à une IA
    Jrn_Add(, {Procédure_Name_Get()})
    Jrn_Add(, {"Liste des Valeurs et des Candidats"})
    Dim Val As String = ""
    Dim Cdd As String = ""
    Dim Numéro As Integer = -1
    For i As Integer = 0 To 80
      Val &= U(i, 2).Replace(" ", "") & ","
      Cdd &= U(i, 3).Replace(" ", "") & ","
    Next i
    Jrn_Add(, {"les 81 Valeurs"})
    Jrn_Add(, {Mid$(Val, 1, Val.Length - 1)})
    Jrn_Add(, {"Les 81 Candidats"})
    Jrn_Add(, {Mid$(Cdd, 1, Cdd.Length - 1)})

    Jrn_Add(, {"Les Valeurs en 9x9"})
    For row As Integer = 0 To 8
      Val = ""
      For col As Integer = 0 To 8
        Numéro = Wh_Cellule_ColRow(col, row)
        Val &= U(Numéro, 2).PadLeft(8) & ","
      Next col
      Jrn_Add(, {Mid$(Val, 1, Val.Length - 1)})
    Next row

    Jrn_Add(, {"Les Candidats en 9x9"})
    For row As Integer = 0 To 8
      Cdd = ""
      For col As Integer = 0 To 8
        Numéro = Wh_Cellule_ColRow(col, row)
        Cdd &= U(Numéro, 3).Replace(" ", "").PadLeft(8) & ","
      Next col
      Jrn_Add(, {Mid$(Cdd, 1, Cdd.Length - 1)})
    Next row

  End Sub
End Module