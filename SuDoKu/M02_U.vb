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
    'U_Coord devient un wrapper
    If Cellule < 1 OrElse Cellule > 81 Then
      Return $"[ERR:{Cellule}]"
    End If
    Return U_cr(Cellule)
  End Function

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
      S = S & " | " & U_cr(i)
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