Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
' Stratégie des CdO
'
' La stratégie CdO consiste à vérifier qu'un des candidats de la cellule ne peut aller qu'à cet endroit
'    de la ligne, de la colonne ou de la région. Il faut donc vérifier l'unicité de ce candidat dans les 3 unités
'-------------------------------------------------------------------------------------------

Friend Module P11_Strategy_CdO
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  Function Strategy_CdO(U_temp(,) As String) As String(,)
    '17/02/2023 Déport de l'initialisation de Cel45 et Exc45
    Dim Stratégie As String = "CdO"
    Dim Sous_Stratégie As String = "CdO"
    'Jrn_Add_Yellow(Proc_Name_Get() & " " & Origine)
    Dim Strategy_Rslt(99, 0) As String 'initialisé  

    'Le tableau Strategy_Rslt est passé en Byref pour pouvoir être augmenté et documenté des CdO
    Strategy_Rslt_Init(Strategy_Rslt, Stratégie, Sous_Stratégie) 'Initialisation 

    Dim Cel45(0 To 44) As String : For k As Integer = 0 To 44 : Cel45(k) = "__" : Next k
    Dim Exc45(0 To 44) As String : For k As Integer = 0 To 44 : Exc45(k) = "__" : Next k

    For Cellule As Integer = 0 To 80
      'La cellule doit être vide et comporter plusieurs candidats
      If (U_temp(Cellule, 2) = " " And Wh_Cell_Nb_Candidats(U_temp, Cellule) > 1) Then
        Dim Grp(0 To 8) As Integer
        'Est-ce que la valeur est trouvée dans l'unité LCR autre que la cellule étudiée ?
        Dim Candidat As String

        For v As Integer = 1 To 9 'On parcourt toutes les valeurs candidates
          Candidat = CStr(v)
          If U_temp(Cellule, 3).Substring(v - 1, 1) = Candidat Then

            Dim LCR_Clct As New Collection      ' Collection des 3 Valeurs 1=L,2=C et 3=R
            Clct_Init(LCR_Clct, 1, 3)
            ' Pour ne pas vérifier toujours dans le même ordre L, C et R, on choisit au hasard l'ordonnancement
            For i As Integer = 0 To 3
              Dim Valeur_LCR As Integer = Clct_Random(LCR_Clct)
              Select Case Valeur_LCR
                Case 1
                  If Not Strategy_CdO_Unité(U_temp, U_9CelRow(U_Row(Cellule)), Cellule, Candidat) Then
                    'Seule la première cellule est utilisée
                    Cel45(0) = CStr(Cellule)                         '  Cellule Concernée par la stratégie
                    Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, "L", CStr(U_Row(Cellule)), Candidat, Cel45, Exc45)
                    Exit For
                  End If
                Case 2
                  If Not Strategy_CdO_Unité(U_temp, U_9CelCol(U_Col(Cellule)), Cellule, Candidat) Then
                    Cel45(0) = CStr(Cellule)                         '  Cellule Concernée par la stratégie
                    Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, "C", CStr(U_Col(Cellule)), Candidat, Cel45, Exc45)
                    Exit For
                  End If
                Case 3
                  If Not Strategy_CdO_Unité(U_temp, U_9CelReg(U_Reg(Cellule)), Cellule, Candidat) Then
                    Cel45(0) = CStr(Cellule)                         '  Cellule Concernée par la stratégie
                    Strategy_Rslt_Add(Strategy_Rslt, Stratégie, Sous_Stratégie, "R", CStr(U_Reg(Cellule)), Candidat, Cel45, Exc45)
                    Exit For
                  End If
              End Select
            Next i
          End If '/Valeur candidate existante
        Next v
      End If
    Next Cellule
    'Jrn_Add_Yellow("/" & Proc_Name_Get() & " " & Origine & " " & Strategy_Rslt.GetLength(1))
    Return Strategy_Rslt
  End Function

  Function Strategy_CdO_Unité(U_temp(,) As String, Grp() As Integer, Cellule As Integer, Candidat As String) As Boolean
    Strategy_CdO_Unité = False
    ' Etude de l' Unité L, C ou R
    For g As Integer = 0 To 8 '/Analyse de l'Unité
      Dim Grp_Cel As Integer = Grp(g)
      ' Pas de traitement pour les cellules remplies, ni pour la cellule étudiée
      If U_temp(Grp_Cel, 2) <> " " Or Grp_Cel = Cellule Then Continue For
      Strategy_CdO_Unité = False
      If U_temp(Grp_Cel, 3).Contains(Candidat) = True Then
        Strategy_CdO_Unité = True : Exit For
      End If
    Next g         '/Analyse de l'Unité
    Return Strategy_CdO_Unité
  End Function
End Module