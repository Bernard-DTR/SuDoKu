Option Strict On
Option Explicit On

Friend Module P30_Strategy_ForceBrute

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'La force brute utilise U, elle est utilisée ESSENTIELLEMENT interactivement ET NON en arrière-plan
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

  '-------------------------------------------------------------------------------
  ' Mise en place le 18/06/2023 Guise-Abbaye de St-Michel en Thiérache
  ' Origine: \SuDoKu_2023\S95_AutresJeux\Planete\VB_Net_Sud1916447222005\Sudoku\Sudoku.sln
  '          Le code Force Brute est issu d'un autre site
  '          L'ensemble des URL est 404!
  '-------------------------------------------------------------------------------
  Dim Profondeur_Rsl_Max As Integer = 0
  Dim Profondeur_Rls As Integer = 0
  ReadOnly Cube_Sdk(9, 9, 11) As Integer
  'la couche 10 de Cube_SDK est la couche des valeurs réelles
  'Pour Al Escargot on a :
  'Couche 10
  '     100007090 162857493
  '     030020008 534129678
  '     009600500 789643521
  '     005300900 475312986
  '     010080002 913586742
  '     600004000 628794135
  '     300000010 356478219
  '     040000007 241935867
  '     007000300 897261354
  '

  Public Sub Cube_Sdk_Lst()
    Jrn_Add(, {"Cube_Sdk"})
    Dim Lng As String
    For k As Integer = 0 To 10
      Jrn_Add("SDK_Space")
      Jrn_Add(, {"Couche " & CStr(k)})
      For i As Integer = 0 To 8
        Lng = ""
        For j As Integer = 0 To 8
          Lng &= CStr(Cube_Sdk(i, j, k))
        Next j
        Jrn_Add(, {Lng})
      Next i
    Next k
  End Sub
  Public Sub Strategy_Force_Brute()
    Dim Rslt As Integer

    Dim i, j, k, x As Integer
    Dim n As Integer
    'Initialisation des couches  0 à 8 = 1
    '               de la couche 9     = 9
    'Calcul de x soit 0, soit la valeur CInt(U(n, 2))
    'S'il y a une valeur alors:
    '  les couches 0 à 8       = 0
    '  la couche   0 si x = 1
    '              1 si x = 2
    '        etc   8 si x = 9  = 1
    '  la couche   9           = 1
    '  la couche  10           = la valeur
    For i = 0 To (8)
      For j = 0 To (8)
        For k = 0 To (8)
          Cube_Sdk(i, j, k) = 1
        Next k
        Cube_Sdk(i, j, 9) = 9

        n = (i * 9) + j
        Select Case U(n, 2)
          Case " " : x = 0
          Case Else : x = CInt(U(n, 2))
        End Select

        If x > 0 Then
          For k = 0 To (8)
            Cube_Sdk(i, j, k) = 0
          Next k
          Cube_Sdk(i, j, x - 1) = 1
          Cube_Sdk(i, j, 9) = 1
          Cube_Sdk(i, j, 10) = x
        End If
      Next j
    Next i
    'Cube_Sdk_Lst()

    Rslt = Calcul_ForceBrute()
    'Cube_Sdk_Lst()
    If Rslt > 0 Then
      Jrn_Add(, {"Stratégie de Force Brute : maximum recursion profondeur: " & Rslt.ToString()})
    Else
      Jrn_Add(, {"Stratégie de Force Brute : ECHEC!"})
    End If
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub

  Public Function Calcul_ForceBrute() As Integer
    Dim Rslt As Boolean
    Rslt = Tentative_de_Résolution()
    'Récupération de la couche 10 dans U(n,2)
    If Rslt = True Then
      Dim n As Integer
      For i As Integer = 0 To 8
        For j As Integer = 0 To 8
          n = (i * 9) + j
          U(n, 2) = CStr(Cube_Sdk(i, j, 10))
        Next j
      Next i
    Else
      Profondeur_Rsl_Max = 0   'indicate it failed
    End If
    Return Profondeur_Rsl_Max
  End Function

  Private Function Tentative_de_Résolution() As Boolean
    Dim Nb_Val_Connues As Integer
    Dim Prv_Nb_Val_Connues As Integer
    Dim i, j As Integer

    Dim Save_Cube_Sdk(9, 9, 11) As Integer
    Dim Abandon As Integer
    Tentative_de_Résolution = False
    Profondeur_Rls += 1
    If (Profondeur_Rsl_Max < Profondeur_Rls) Then
      Profondeur_Rsl_Max = Profondeur_Rls
    End If
    Nb_Val_Connues = Calcul_du_Nombre_Valeurs_Connues()
    While (Nb_Val_Connues > Prv_Nb_Val_Connues) And (Nb_Val_Connues < 81)
      Prv_Nb_Val_Connues = Nb_Val_Connues
      For i = 0 To 8
        For j = 0 To 8
          If Cube_Sdk(i, j, 9) > 1 Then
            Résolution_Row(i, j)
            Résolution_Col(i, j)
            Résolution_Région(i, j)
          End If
        Next j
      Next i
      Nb_Val_Connues = Calcul_du_Nombre_Valeurs_Connues()
    End While

    If Contrôle() = False Then
      Profondeur_Rls -= 1
      Return False
    End If

    If (Nb_Val_Connues = 81) Then
      Profondeur_Rls -= 1
      Return True
    End If

    Abandon = 0
    For i = 0 To 8
      For j = 0 To 8
        If Cube_Sdk(i, j, 9) > 1 Then
          Abandon = 1
          Exit For
        End If
      Next j
      If Abandon = 1 Then
        Exit For
      End If
    Next i

    Dim CS_Getlength As Integer = Cube_Sdk.GetLength(2)

    Dim Save_Entrée As Integer()
    ReDim Save_Entrée(CS_Getlength)
    For u As Integer = 0 To CS_Getlength - 1
      Save_Entrée(u) = Cube_Sdk(i, j, u)
    Next u

    For l As Integer = 0 To 8
      If Save_Entrée(l) = 1 Then
        For a As Integer = 0 To Cube_Sdk.GetLength(0) - 1
          For b As Integer = 0 To Cube_Sdk.GetLength(1) - 1
            For c As Integer = 0 To Cube_Sdk.GetLength(2) - 1
              Save_Cube_Sdk(a, b, c) = Cube_Sdk(a, b, c)
            Next c
          Next b
        Next a
        For m As Integer = 0 To 8
          Cube_Sdk(i, j, m) = 0
        Next m
        Cube_Sdk(i, j, l) = 1
        Cube_Sdk(i, j, 9) = 1
        Cube_Sdk(i, j, 10) = l + 1
        Dim Succès As Boolean
        Succès = Tentative_de_Résolution()
        If Succès = True Then
          Profondeur_Rls -= 1
          Return Succès
        Else
          For a As Integer = 0 To Cube_Sdk.GetLength(0) - 1
            For b As Integer = 0 To Cube_Sdk.GetLength(1) - 1
              For c As Integer = 0 To Cube_Sdk.GetLength(2) - 1
                Cube_Sdk(a, b, c) = Save_Cube_Sdk(a, b, c)
              Next c
            Next b
          Next a
        End If
      End If
    Next l
  End Function

  Private Function Calcul_du_Nombre_Valeurs_Connues() As Integer
    ' Peut atteindre 81
    Dim n As Integer
    For i As Integer = 0 To 8
      For j As Integer = 0 To 8
        If Cube_Sdk(i, j, 9) = 1 Then n += 1
      Next j
    Next i
    Return n
  End Function

  Private Sub Résolution_Row(i As Integer, j As Integer)
    For k As Integer = 0 To (8)
      If (Not (j = k)) AndAlso (Cube_Sdk(i, k, 9) = 1) AndAlso (Cube_Sdk(i, j, 9) > 1) Then
        If ((Cube_Sdk(i, j, Cube_Sdk(i, k, 10) - 1)) = 1) Then
          Cube_Sdk(i, j, Cube_Sdk(i, k, 10) - 1) = 0
          Dim c As Integer = Cube_Sdk(i, j, 9)
          Cube_Sdk(i, j, 9) = c - 1
          If Cube_Sdk(i, j, 9) = 1 Then
            Cube_Sdk(i, j, 10) = Valeur_Unique(i, j)
          End If
        End If
      End If
    Next k
  End Sub

  Private Sub Résolution_Col(i As Integer, j As Integer)
    For k As Integer = 0 To 8
      If (Not (i = k)) AndAlso (Cube_Sdk(k, j, 9) = 1) AndAlso (Cube_Sdk(i, j, 9) > 1) Then
        If Cube_Sdk(i, j, Cube_Sdk(k, j, 10) - 1) = 1 Then
          Cube_Sdk(i, j, Cube_Sdk(k, j, 10) - 1) = 0
          Dim c As Integer = Cube_Sdk(i, j, 9)
          Cube_Sdk(i, j, 9) = c - 1
          If Cube_Sdk(i, j, 9) = 1 Then
            Cube_Sdk(i, j, 10) = Valeur_Unique(i, j)
          End If
        End If
      End If
    Next k
  End Sub

  Private Sub Résolution_Région(i As Integer, j As Integer)
    Dim x, y As Integer
    x = (i \ 3) * 3
    y = (j \ 3) * 3
    For k As Integer = x To x + 2
      For l As Integer = y To y + 2
        If ((Not (k = i)) Or (Not (l = j))) AndAlso (Cube_Sdk(k, l, 9) = 1) Then
          If (Cube_Sdk(i, j, Cube_Sdk(k, l, 10) - 1) = 1) Then
            Cube_Sdk(i, j, Cube_Sdk(k, l, 10) - 1) = 0
            Dim c As Integer = Cube_Sdk(i, j, 9)
            Cube_Sdk(i, j, 9) = c - 1
            If Cube_Sdk(i, j, 9) = 1 Then
              Cube_Sdk(i, j, 10) = Valeur_Unique(i, j)
            End If
          End If
        End If
      Next l
    Next k
  End Sub

  Private Function Valeur_Unique(x As Integer, y As Integer) As Integer
    Dim Val_Unique As Integer = 10
    For i As Integer = 0 To 8
      If Cube_Sdk(x, y, i) = 1 Then
        Val_Unique = i + 1
        Exit For
      End If
    Next i
    Return Val_Unique
  End Function

  Private Function Contrôle() As Boolean
    Dim x, y As Integer
    For i As Integer = 0 To (8)
      For j As Integer = 0 To (8)
        If Cube_Sdk(i, j, 9) = 0 Then
          Return False
        End If
        For k As Integer = 0 To (8)
          If (Cube_Sdk(i, k, 9) = 1) AndAlso (Cube_Sdk(i, j, 9) = 1) _
               AndAlso (Not (j = k)) AndAlso (Cube_Sdk(i, k, 10) = Cube_Sdk(i, j, 10)) Then
            Return False
          End If
        Next k
        For k As Integer = 0 To (8)
          If (Cube_Sdk(k, j, 9) = 1) AndAlso (Cube_Sdk(i, j, 9) = 1) _
               AndAlso (Not (i = k)) AndAlso (Cube_Sdk(k, j, 10) = Cube_Sdk(i, j, 10)) Then
            Return False
          End If
        Next k
        x = (i \ 3) * 3
        y = (j \ 3) * 3
        For k As Integer = x To x + 2
          For l As Integer = y To y + 2
            If (Cube_Sdk(k, l, 9) = 1) AndAlso (Cube_Sdk(i, j, 9) = 1) _
                AndAlso ((Not (k = i)) Or (Not (l = j))) AndAlso
               (Cube_Sdk(k, l, 10) = Cube_Sdk(i, j, 10)) Then
              Return False
            End If
          Next l
        Next k
      Next j
    Next i
    Return True
  End Function
End Module