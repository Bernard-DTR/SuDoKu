Option Strict On
Option Explicit On

Module M10_Transformation
  '-------------------------------------------------------------------------------
  'Procédure de Transformation d'une grille
  '04/03/2024 Permutation de lignes par bande
  '           Permutation de colonnes par pile
  'https://les-mathematiques.net/vanilla/discussion/1778116/bi-cycles-dans-un-sudoku
  'https://fr.wikipedia.org/wiki/Math%C3%A9matiques_du_sudoku
  'https://perso.univ-rennes1.fr/matthieu.romagny/agreg/exo/sudoku.pdf
  'Il y a permutation des valeurs et persistance des positions des VI
  'Uu(80) booleen reprend U(i,1) pour préciser si la cellule est ou n'est pas Valeur initiale
  'La grille doit être complète.
  'Une Bande est un groupe de 3 Régions Horizontales
  'Une Pile                   3 Régions Verticales
  'Une permutation est Horizontale ou Verticale
  '                concerne les Bandes ou les Piles
  '                1, 2, 3, 1-2, 1-3, 2-3 ou 1-2-3 
  'Une fois la Bande ou la Pile choisie,
  '                concerne les lignes ou les colonnes
  '                1-3-2, 2-1-3, 2-3-1, 3-1-2 et 3-2-1  
  '-------------------------------------------------------------------------------
  ' Denis Berthier
  '  Symétries « géométriques » de la grille : 
  'permutations des rangées individuelles 1, 2, 3 
  'permutations des rangées individuelles 4, 5, 6 ; 
  'permutations des rangées individuelles 7, 8, 9 ; 
  'permutations de triolets des rangées 1-3, 4-6 et 7-9 ; 
  'symétrie par rapport à la première diagonale (symétrie ligne-colonne). 
  'De ces symétries géométriques primaires, d'autres peuvent être déduites : 
  'permutations de colonnes individuelles 1, 2, 3 ; 
  'permutations de colonnes individuelles 4, 5, 6 ; - permutations de colonnes individuelles 7, 8, 9 ; 
  'permutations de triolets des colonnes 1-3, 4-6 et 7-9 ; 
  'réflexion (symétrie gauche-droite) ; 
  'symétrie haut-bas ; 
  'symétrie par rapport à la deuxième diagonale ; 
  'rotation 
  '-------------------------------------------------------------------------------

  Public Sub Transf_Permutation()
    Dim U_temp(80, 3) As String
    Dim Uu(80) As Boolean
    'Programme de transformation par permutation des Bandes ou des Piles

    '01 Vérification de grille correcte et complète 
    Dim U_Chk(80, 3) As String
    Array.Copy(U_temp, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    If Not U_Check.Check Then
      Jrn_Add(, {Proc_Name_Get() & " La grille est incorrecte."})
      Exit Sub
    End If
    If U_Check.Nb_Remplies <> 81 Then
      Jrn_Add(, {Proc_Name_Get() & " La grille n'est pas complète."})
      Exit Sub
    End If

    '02 Permutation par Bande ou par Pile
    Dim Rand As Integer
    '1 Incrémentation
    Rand = Rda.Next(1, 3) 'inclu min, mais pas Max 
    Select Case Rand
      Case 1 ' Permutation par Bande (Horizontale)
        Array.Copy(U, U_temp, UNbCopy)
      Case 2 ' Permutation par Pile  (Verticale)
        Dim Cellule_O As Integer '  Cellule Origine
        Dim Cellule_D As Integer '  Cellule Destination
        For c As Integer = 0 To 8
          For r As Integer = 0 To 8
            Cellule_O = Wh_Cellule_ColRow(r, c)
            Cellule_D = Wh_Cellule_ColRow(c, r)
            U_temp(Cellule_D, 1) = U(Cellule_O, 1)
            U_temp(Cellule_D, 2) = U(Cellule_O, 2)
            U_temp(Cellule_D, 3) = U(Cellule_O, 3)
          Next r
        Next c
        Dim Nom As String = "Permutation V"
        Dim Prb As String = ""
        Dim Jeu As String = ""
        Dim Frc As String = "0"
        For i As Integer = 0 To 80
          Prb &= U_temp(i, 1)
          Jeu &= U_temp(i, 2)
        Next i

        Game_New_Game(Gnrl:="Nrm", "   ", Nom:=Nom, Prb:=Prb, Jeu:=Prb, Sol:=Prb, Cdd729:=StrDup(729, " "), Frc:=Frc, Proc_Name_Get())

    End Select
    For i As Integer = 0 To 80
      Uu(i) = False
      If U_temp(i, 1) <> " " Then Uu(i) = True
    Next i
  End Sub
  Public Source(9, 9, 2) As String 'Source
  Public Cible(9, 9, 2) As String 'Cible

  ' 0 --- 1 Valeur initiale
  ' 1 --- 6 Solution

  Sub Transf_Grid_Tfr_Display(Poste As Integer)
    'Liste U avec For r et For c et i = (c + (r * 9))
    Select Case Poste
      Case 1 : Jrn_Add(, {"Liste de la Grille et de la Solution Poste 1"})
      Case 2 : Jrn_Add(, {"Liste de la Grille et de la Solution Poste 2"})
    End Select
    Dim Slg As String      'Source_Ligne
    Dim Clg As String      'Cible_Ligne 
    Dim cr As String = ""  'col-row
    Dim Esp As String = "          "
    Dim i As Integer
    For r As Integer = 0 To 8
      Slg = ""
      Clg = ""
      For c As Integer = 0 To 8
        i = (c + (r * 9))
        Select Case Poste
          Case 1 : If U(i, 1) = " " Then cr = "." Else cr = U(i, 1)
          Case 2 : If U(i, 2) = " " Then cr = "." Else cr = U(i, 2)
        End Select
        Slg &= cr
        If (c = 2 Or c = 5) Then Slg &= " | "
        If U_Sol(i) = " " Then cr = "." Else cr = U_Sol(i)
        Clg &= cr
        If (c = 2 Or c = 5) Then Clg &= " | "
      Next c
      Jrn_Add(, {Slg & Esp & Clg})
      If (r = 2 Or r = 5) Then Jrn_Add(, {"--- + --- + ---" & Esp & "--- + --- + ---"})
    Next r
  End Sub

  Sub U_Copy_to_Source()
    'Transfert de U avec For r et For c et i = (c + (r * 9)) dans Source(c,r,01)
    Dim i As Integer
    For r As Integer = 0 To 8
      For c As Integer = 0 To 8
        i = (c + (r * 9))
        Source(c, r, 0) = U(i, 1)
        Source(c, r, 1) = U_Sol(i)
      Next c
    Next r
  End Sub
  Sub Cible_Copy_to_Source()
    'Transfert de Cible(c,r,01) vers Source(c,r,01)
    For r As Integer = 0 To 8
      For c As Integer = 0 To 8
        Source(c, r, 0) = Cible(c, r, 0)
        Source(c, r, 1) = Cible(c, r, 1)
      Next c
    Next r
  End Sub
  Sub Transf_Incrémentation(Inc As Integer)
    'Transfert Incrémentation
    U_Copy_to_Source()
    Dim Origine As String = "Transfert Incrémentation " & CStr(Inc)
    For r As Integer = 0 To 8
      For c As Integer = 0 To 8
        If Source(c, r, 0) <> " " Then Cible(c, r, 0) = CStr(Transf_Add(CInt(Source(c, r, 0)), Inc)) Else Cible(c, r, 0) = " "
        If Source(c, r, 1) <> " " Then Cible(c, r, 1) = CStr(Transf_Add(CInt(Source(c, r, 1)), Inc)) Else Cible(c, r, 1) = " "
      Next c
    Next r
    Transf_Display(Origine, 0)
    Transf_FindeTraitement(Origine)
  End Sub
  Sub Transf_Rotation()
    'Transfert Rotation 90°
    U_Copy_to_Source()
    Dim Origine As String = "Transfert Rotation 90° "
    For r As Integer = 0 To 8
      For c As Integer = 0 To 8
        Cible(c, r, 0) = Source(r, Transf_Sym(c), 0)
        Cible(c, r, 1) = Source(r, Transf_Sym(c), 1)
      Next c
    Next r
    Transf_Display(Origine, 0)
    Transf_FindeTraitement(Origine)
  End Sub
  Sub Transf_Symétrie(Sym As Integer)
    Dim Origine As String = ""
    U_Copy_to_Source()
    'Transformation
    Select Case Sym
      Case 1 : Origine = "Médiane Horizontale"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(c, Transf_Sym(r), 0)
            Cible(c, r, 1) = Source(c, Transf_Sym(r), 1)
          Next c
        Next r
      Case 2 : Origine = "Médiane Verticale"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(Transf_Sym(c), r, 0)
            Cible(c, r, 1) = Source(Transf_Sym(c), r, 1)
          Next c
        Next r
      Case 3 : Origine = "Diagonale Droite"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(Transf_Sym(c), Transf_Sym(r), 0)
            Cible(c, r, 1) = Source(Transf_Sym(c), Transf_Sym(r), 1)
          Next c
        Next r
      Case 4 : Origine = "Diagonale Gauche"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(Transf_Sym(r), Transf_Sym(c), 0)
            Cible(c, r, 1) = Source(Transf_Sym(r), Transf_Sym(c), 1)
          Next c
        Next r
    End Select
    Transf_Display(Origine, 0)
    Transf_FindeTraitement(Origine)
  End Sub
  Sub Transf_Région_H()
    Dim Origine As String = "Transf_Région_H"
    U_Copy_to_Source()
    Dim nRc As Integer
    For R As Integer = 0 To 8 'Région N° 0 à 8
      For nRs As Integer = 0 To 8
        Select Case nRs
          Case 0 : nRc = 2
          Case 1 : nRc = 1
          Case 2 : nRc = 0
          Case 3 : nRc = 5
          Case 4 : nRc = 4
          Case 5 : nRc = 3
          Case 6 : nRc = 8
          Case 7 : nRc = 7
          Case 8 : nRc = 6
        End Select
        Dim Cel_Source As Integer = Wh_NumérodelaCellule_RegNreg(R, nRs)
        Dim Cel_Cible As Integer = Wh_NumérodelaCellule_RegNreg(R, nRc)
        Cible(U_Col(Cel_Cible), U_Row(Cel_Cible), 0) = Source(U_Col(Cel_Source), U_Row(Cel_Source), 0)
        Cible(U_Col(Cel_Cible), U_Row(Cel_Cible), 1) = Source(U_Col(Cel_Source), U_Row(Cel_Source), 1)
      Next nRs
    Next R
    Transf_Display(Origine, 0)
    Transf_FindeTraitement(Origine)
  End Sub
  Sub Transf_Région_V()
    Dim Origine As String = "Transf_Région_V"
    U_Copy_to_Source()
    Dim nRc As Integer
    For R As Integer = 0 To 8 'Région N° 0 à 8
      For nRs As Integer = 0 To 8
        Select Case nRs
          Case 0 : nRc = 6
          Case 1 : nRc = 7
          Case 2 : nRc = 8
          Case 3 : nRc = 3
          Case 4 : nRc = 4
          Case 5 : nRc = 5
          Case 6 : nRc = 0
          Case 7 : nRc = 1
          Case 8 : nRc = 2
        End Select
        Dim Cel_Source As Integer = Wh_NumérodelaCellule_RegNreg(R, nRs)
        Dim Cel_Cible As Integer = Wh_NumérodelaCellule_RegNreg(R, nRc)
        Cible(U_Col(Cel_Cible), U_Row(Cel_Cible), 0) = Source(U_Col(Cel_Source), U_Row(Cel_Source), 0)
        Cible(U_Col(Cel_Cible), U_Row(Cel_Cible), 1) = Source(U_Col(Cel_Source), U_Row(Cel_Source), 1)
      Next nRs
    Next R
    Transf_Display(Origine, 0)
    Transf_FindeTraitement(Origine)
  End Sub
  Sub Transf_Aléatoire()
    U_Copy_to_Source()
    Dim Origine As String = "Transformation aléatoire"
    Dim Rand As Integer
    '1 Incrémentation
    Rand = Rd6.Next(0, 9) 'inclu min, mais pas Max si 0, il n'y a pas d'incrémentation
    For r As Integer = 0 To 8
      For c As Integer = 0 To 8
        If Source(c, r, 0) <> " " Then Cible(c, r, 0) = CStr(Transf_Add(CInt(Source(c, r, 0)), Rand)) Else Cible(c, r, 0) = " "
        If Source(c, r, 1) <> " " Then Cible(c, r, 1) = CStr(Transf_Add(CInt(Source(c, r, 1)), Rand)) Else Cible(c, r, 1) = " "
      Next c
    Next r
    Transf_Display("Incrémentation " & CStr(Rand), 0)
    Cible_Copy_to_Source()

    '2 Symétrie
    Rand = Rd6.Next(1, 5)
    Select Case Rand
      Case 1 : Origine = "Médiane Horizontale"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(c, Transf_Sym(r), 0)
            Cible(c, r, 1) = Source(c, Transf_Sym(r), 1)
          Next c
        Next r
      Case 2 : Origine = "Médiane Verticale"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(Transf_Sym(c), r, 0)
            Cible(c, r, 1) = Source(Transf_Sym(c), r, 1)
          Next c
        Next r
      Case 3 : Origine = "Diagonale Droite"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(Transf_Sym(c), Transf_Sym(r), 0)
            Cible(c, r, 1) = Source(Transf_Sym(c), Transf_Sym(r), 1)
          Next c
        Next r
      Case 4 : Origine = "Diagonale Gauche"
        For r As Integer = 0 To 8
          For c As Integer = 0 To 8
            Cible(c, r, 0) = Source(Transf_Sym(r), Transf_Sym(c), 0)
            Cible(c, r, 1) = Source(Transf_Sym(r), Transf_Sym(c), 1)
          Next c
        Next r
      Case 5 : Origine = "Aucune symétrie"
    End Select
    Transf_Display(Origine, 0)
    Cible_Copy_to_Source()

    '3 Rotation
    Rand = Rd6.Next(1, 4)
    Origine = "Transfert Rotation 90° x " & CStr(Rand)
    For i As Integer = 1 To Rand
      For r As Integer = 0 To 8
        For c As Integer = 0 To 8
          Cible(c, r, 0) = Source(r, Transf_Sym(c), 0)
          Cible(c, r, 1) = Source(r, Transf_Sym(c), 1)
        Next c
      Next r
      Transf_Display(Origine, 0)
      Cible_Copy_to_Source()
    Next i

    '4 Région
    Rand = Rd6.Next(1, 3)
    Select Case Rand
      Case 1 : Origine = "Région Horizontale"
        Transf_Région_H()
      Case 2 : Origine = "Région Verticale"
        Transf_Région_V()
      Case 3 : Origine = "Pas de région"
    End Select
    Transf_Display(Origine, 0)
    Cible_Copy_to_Source()

    Origine = "Transformation aléatoire"
    Transf_FindeTraitement(Origine)
  End Sub
  Sub Transf_FindeTraitement(Origine As String)
    'Fin de la transformation
    Dim Procédure As String = Proc_Name_Get() & " " & Origine
    Dim Nom As String = Procédure
    Dim Prb As String = ""
    Dim Jeu As String = StrDup(81, " ")
    Dim Sol As String = ""
    Dim Frc As String = "0"
    For r As Integer = 0 To 8
      For c As Integer = 0 To 8
        Prb &= Cible(c, r, 0)
        Sol &= Cible(c, r, 1)
      Next c
    Next r
    Game_New_Game(Gnrl:="Nrm", "   ", Nom:=Nom, Prb:=Prb, Jeu:=Prb, Sol:=Sol, Cdd729:=StrDup(729, " "), Frc:=Frc, Proc_Name_Get())
  End Sub
  Sub Transf_Display(Origine As String, i As Integer)
    'Liste Source et Csr Valeur Initiale et Solution
    Jrn_Add("SDK_Space")
    Select Case i
      Case 0 : Jrn_Add(, {"Liste " & Origine & " _ Problème "})
      Case 1 : Jrn_Add(, {"Liste " & Origine & " _ Solution "})
    End Select
    Dim Slg As String
    Dim Clg As String
    Dim Cr As String
    Dim Esp As String = "           "
    For r As Integer = 0 To 8
      Slg = ""
      Clg = ""
      For c As Integer = 0 To 8
        Cr = " "
        If CStr(Source(c, r, i)) = " " Then Cr = "."
        If CStr(Source(c, r, i)) <> " " Then Cr = CStr(Source(c, r, i))
        Select Case c
          Case 2, 5 : Slg &= Cr & " | "
          Case Else : Slg &= Cr
        End Select
        Cr = " "
        If CStr(Cible(c, r, i)) = " " Then Cr = "."
        If CStr(Cible(c, r, i)) <> " " Then Cr = CStr(Cible(c, r, i))
        Select Case c
          Case 2, 5 : Clg &= Cr & " | "
          Case Else : Clg &= Cr
        End Select
      Next c
      Select Case r
        Case 2, 5
          Jrn_Add(, {Slg & Esp & Clg})
          Jrn_Add(, {"--- + --- + ---" & Esp & "--- + --- + ---"})
        Case Else
          Jrn_Add(, {Slg & Esp & Clg})
      End Select
    Next r
    Jrn_Add("SDK_Space")
  End Sub
  Function Transf_Add(a As Integer, c As Integer) As Integer
    Dim Add As Integer = a + c
    If Add > 9 Then Add -= 9
    Return Add
  End Function
  Function Transf_Sym(c As Integer) As Integer
    Dim Sym As Integer = -1
    Select Case c
      Case 0 : Sym = 8
      Case 1 : Sym = 7
      Case 2 : Sym = 6
      Case 3 : Sym = 5
      Case 4 : Sym = 4
      Case 5 : Sym = 3
      Case 6 : Sym = 2
      Case 7 : Sym = 1
      Case 8 : Sym = 0
    End Select
    Return Sym
  End Function
End Module