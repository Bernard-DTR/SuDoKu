Namespace DancingLink
  Module A_Copyright

#Region "Copyright" ' Copyright (c) 2006 Miran Uhan
    ' - Donald E. Knuth "Dancing Links" (see
    '   http://www-cs-faculty.stanford.edu/~knuth/preprints.html).
#End Region

    ' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Public Structure DL_Solve_Struct
      ' Nb_Solution = -1 U_temp(,) est contrôlé Not U_Check.Check  
      '                0 La grille n'a pas de solutions
      '                1 La GRILLE EST un PUZZLE
      '                n La grille a plusieurs solutions
      ' Nb_Solution = -1 Solution(0) = StrDup(81, "0")
      '                0 Solution(0) = StrDup(81, "0")
      '                1 Solution(0) comporte la solution
      '                n Solution(0) = StrDup(81, "0") et Solution(de 1 à n) comporte les solutions trouvées
      Public Nb_Solution As Integer  ' Le  Nombre de Solution (-1, 0, 1 ou n)
      Public DLCode As String        ' Dl# Erreur
      '                              ' Dlz Pas de Solution
      '                              ' Dlu Une solution, donc LA SOLUTION
      '                              ' Dln Le nombre de solutions
      Public Solution() As String    ' Les Solutions
    End Structure
    Public Function DL_Solve(U_temp(,) As String) As DL_Solve_Struct
      ' Private _store As SudokuStore est modifié, il passe de = SudokuStore.First
      '                                                     à  = SudokuStore.All 
      ' seules les 9 premières solutions sont enregistrées et listées If _solutionList.Count < 9  Then _solutionList.Add(_solution)
      ' Nécessaire U_Check.Check = True
      ' En entrée U_temp       Standard  ( U_temp(i, 1) sera utilisée, alors que U_Checking utilise 1, 2 et 3) 
      '                                    U_Temp(i, 1) est utilisée pour une grille vide sans VI
      '                                    U_Temp(i, 2) est utilisée pour une grille remplie
      ' En sortie la structure DL_Solve_Struct comportant 
      '                        Le nombre de solutions: -1 u_temp est en erreur avec U_Check
      '                                                 0 Aucune solution
      '                                                 1 LA SOLUTION (la grille est un Puzzle)
      '                                                 x Nombre de solutions trouvées. 
      '                        et LA solution sous la forme d'une chaîne de 81 valeurs (L1C1, L1C2, ... L9C9) ou 81 valeurs à 0

      ' Initialisation de la structure DL_Solve_Struct
      Dim DL As New DL_Solve_Struct With {
        .Nb_Solution = -1,
        .DLCode = "DL#",
        .Solution = {StrDup(81, "0")}
    }

      Try
        ' 1 Contrôle de U_temp
        Dim U_Chk(80, 3) As String
        Array.Copy(U_temp, U_Chk, UNbCopy)
        Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
        If Not U_Check.Check Then Return DL

        ' Dancing Link n'est pas exécuté si la grille est vide 
        Dim Nb_Cellules_Initiales As Integer = 0
        For i As Integer = 0 To 80
          If U_temp(i, 1) <> " " Then Nb_Cellules_Initiales += 1
        Next i
        If Nb_Cellules_Initiales = 0 OrElse Nb_Cellules_Initiales <= Dl_Nb_VI_minimal Then Return DL

        ' 2 Chargement de U_temp dans un tableau iRet(8, 8) en r,c
        Dim iRet(8, 8) As Integer
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Dim i As Integer = c + (r * 9)
            iRet(c, r) = If(U_temp(i, 1) <> " ", CInt(U_temp(i, 1)), 0)
          Next c
        Next r

        Dim _arena As New C_SudokuArena(iRet, 3, 3)
        Dim _sol As Array
        ' 3 Traitement DL
        _arena.Solve()
        DL.Nb_Solution = _arena.Solutions

        Select Case _arena.Solutions
          Case 0 ' Pas de solution
            DL.Nb_Solution = 0
            DL.DLCode = "DLz"
            DL.Solution(0) = StrDup(81, "0")
            Return DL

          Case 1 ' Solution unique
            DL.DLCode = "Dlu"
            _sol = _arena.Solution()
            ' 4 Dans le cas d'une solution unique, elle sera affichée en tant que solution 
            Dim U_DL(0 To 80) As String
            DL.Solution(0) = ""
            Dim n As Integer = -1
            ' La composition de U_DL se fait colonne par colonne
            For Each v As Object In _sol
              n += 1
              Dim i As Integer = Wh_Cellule_ColRow(n \ 9, n Mod 9)
              U_DL(i) = CStr(v)
            Next v
            DL.Solution(0) = String.Join("", U_DL)
            Return DL

          Case Else ' Plusieurs solutions (Maximum enregistrées : 10)
            DL.DLCode = "DL" & _arena.Solutions.ToString()
            ReDim DL.Solution(9)
            For j As Integer = 1 To _arena.Solutions
              If j > 9 Then Exit For
              _sol = _arena.Solution(j)
              Dim U_DL(0 To 80) As String
              Dim S As String = ""
              Dim n As Integer = -1
              ' La composition de U_DL se fait colonne par colonne
              For Each v As Object In _sol
                n += 1
                Dim i As Integer = Wh_Cellule_ColRow(n \ 9, n Mod 9)
                U_DL(i) = CStr(v)
              Next v
              DL.Solution(j) = String.Join("", U_DL)
            Next j
            Return DL
        End Select
      Catch ex As Exception
        '28/05/2024 le message permet de comprendre l'arrêt anormal du traitement
        Dim MsgTit As String = $"{Proc_Name_Get()} {Application.ProductName} {SDK_Version}"
        MsgBox(ex.ToString(),, MsgTit)
      End Try

      Return DL
    End Function

  End Module
End Namespace