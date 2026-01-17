Option Strict On
Option Explicit On

' Mise en place le 24/01/2025
' Stratégie des Derniers Candidats dans les unités LCR et les bandes
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Il est impératif d'utiliser U_temp, copie de U, pour utiliser la stratégie interactivement ET en arrière-plan
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' 17/11/2025
' La stratégie DCdd utilise désormais DCdd_List en lieu et place de Strategy_Rslt
' elle est optimisée

Friend Module P09_Strategy_DCdd

  Public Function DCdd_List_Exists(Cellule As Integer) As Boolean
    ' Retourne True si la cellule est dans la liste
    ' Utilisation de Linq avec IA .Find peut aussi être utilisé
    Return DCdd_List.Any(Function(item) item.Cellule = Cellule)
  End Function
  Public Function DCdd_Get(Cellule As Integer) As DCdd_Cls
    For Each DCdd As DCdd_Cls In DCdd_List
      If DCdd.Cellule = Cellule Then Return DCdd
    Next DCdd
    Return New DCdd_Cls(Cellule, " ", "#", "#")
  End Function
  Public Function DCdd_List_Add(Cellule As Integer, Candidat As String, Stratégie As String, Sous_Stratégie As String) As Integer
    ' Ajoute une entrée inexistante dans DCdd_List et retourne le nombre d'éléments 
    If Not DCdd_List_Exists(Cellule) Then DCdd_List.Add(New DCdd_Cls(Cellule, Candidat, Stratégie, Sous_Stratégie))
    Return DCdd_List.Count
  End Function
  Public Sub DCdd_List_Display()
    ' Affichage de la liste  DCdd_List  
    Jrn_Add(, {"Affichage de DCdd_List :  " & DCdd_List.Count & " Lignes."})
    If DCdd_List.Count = 0 Then Exit Sub
    Dim Nb As Integer = 0
    For Each DCdd As DCdd_Cls In DCdd_List
      Nb += 1
      Jrn_Add(, {CStr(Nb).PadLeft(2) & " " & " Stratégie " & DCdd.Stratégie & " " & U_Coord(DCdd.Cellule) & " Candidat : " & DCdd.Candidat & " Sous_Stratégie " & DCdd.Sous_Stratégie})
    Next DCdd
  End Sub

  Public Sub Strategy_DCd(U_temp(,) As String)
    ' La stratégie du dernier candidat concerne uniquement une et une seule cellule  
    Plcy_Strg = "DCd"
    Dim Stratégie As String = "DCd"
    'Jrn_Add_Orange(Procédure_Name_Get())
    DCdd_List.Clear()

    If Plcy_AfficherDCdd_Bande Then
      ' Analyse des Bandes pour le Dernier Candidat  
      Dim Bandes As New Dictionary(Of Integer, String) From {{0, "Bh0"}, {1, "Bh1"}, {2, "Bh2"}, {3, "Bv0"}, {4, "Bv1"}, {5, "Bv2"}}
      For Each Bande As KeyValuePair(Of Integer, String) In Bandes
        ' Le traitement est effectué pour les 6 bandes
        ' Pour chaque candidat, 
        '      si la valeur est présente 2 fois et si le candidat est présent une seule fois
        '      alors il est dernier candidat de la bande. 

        Dim Cellule As Integer
        For Cdd As Integer = 1 To 9
          Dim Grp27(0 To 26) As Integer
          For j As Integer = 0 To 26
            Grp27(j) = U_Bandes(Bande.Key, j)
          Next j
          Dim Val_Présente As Integer = 0
          Dim Cdd_Cellule As Integer = -1
          Dim Cdd_Présent As Integer = 0
          For Vs As Integer = 0 To 26
            Cellule = Grp27(Vs)
            If U_temp(Cellule, 2) = CStr(Cdd) Then Val_Présente += 1
            If U_temp(Cellule, 3).Contains(CStr(Cdd)) Then
              Cdd_Cellule = Cellule
              Cdd_Présent += 1
            End If
          Next Vs
          If Val_Présente = 2 And Cdd_Présent = 1 Then
            If DCdd_List_Add(Cdd_Cellule, CStr(Cdd), Stratégie, Bande.Value) >= DCdd_Max Then Exit Sub
          End If
        Next Cdd
      Next Bande
    End If '/Plcy_AfficherDCdd_Bande

    For i As Integer = 0 To Cell_FY_List.Count - 1
      Dim Cellule As Integer = Cell_FY_List(i)
      If U_temp(Cellule, 2) <> " " Then Continue For         ' La cellule doit être vide   
      ' Phase 11 les CdU
      If Trim(U_temp(Cellule, 3)).Length = 1 Then            ' Il n'y a qu'un seul candidat 
        If DCdd_List_Add(Cellule, Trim(U_temp(Cellule, 3)), Stratégie, "CdU") >= DCdd_Max Then Exit Sub
        Continue For ' On passe à la cellule suivante
      End If

      ' Phase 12 les CdO
      If Trim(U_temp(Cellule, 3)).Length > 1 Then            ' Il y a plusieurs candidats 

        Dim Grp() As Integer
        Dim Strategy_CdO As Boolean

        For v As Integer = 1 To 9 'On parcourt toutes les valeurs candidates
          If Not U_temp(Cellule, 3).Contains(CStr(v)) Then Continue For
          Grp = U_9CelRow(U_Row(Cellule)) ' Traitement des Lignes
          For g As Integer = 0 To 8
            If U_temp(Grp(g), 2) <> " " Or Grp(g) = Cellule Then Continue For
            Strategy_CdO = False
            If U_temp(Grp(g), 3).Contains(CStr(v)) = True Then
              Strategy_CdO = True : Exit For
            End If
          Next g
          If Not Strategy_CdO Then
            If DCdd_List_Add(Cellule, CStr(v), Stratégie, "CdO_L") >= DCdd_Max Then Exit Sub
            Exit For
          End If

          Grp = U_9CelCol(U_Col(Cellule)) ' Traitement des Colonnes
          For g As Integer = 0 To 8
            If U_temp(Grp(g), 2) <> " " Or Grp(g) = Cellule Then Continue For
            Strategy_CdO = False
            If U_temp(Grp(g), 3).Contains(CStr(v)) = True Then
              Strategy_CdO = True : Exit For
            End If
          Next g
          If Not Strategy_CdO Then
            If DCdd_List_Add(Cellule, CStr(v), Stratégie, "CdO_C") >= DCdd_Max Then Exit Sub
            Exit For
          End If

          Grp = U_9CelReg(U_Reg(Cellule)) ' Traitement des Régions
          For g As Integer = 0 To 8
            If U_temp(Grp(g), 2) <> " " Or Grp(g) = Cellule Then Continue For
            Strategy_CdO = False
            If U_temp(Grp(g), 3).Contains(CStr(v)) = True Then
              Strategy_CdO = True : Exit For
            End If
          Next g
          If Not Strategy_CdO Then
            If DCdd_List_Add(Cellule, CStr(v), Stratégie, "CdO_R") >= DCdd_Max Then Exit Sub

            Exit For
          End If
        Next v
      End If
    Next i
  End Sub
End Module