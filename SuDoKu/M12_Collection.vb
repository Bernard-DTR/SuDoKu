Module M12_Collection
  '-------------------------------------------------------------------------------
  ' Traitement des Collections et accès aléatoire
  ' Une Collection est un ensemble d'éléments ordonnés qui peut être désigné sous le nom d'unité,
  '-------------------------------------------------------------------------------
  Sub Clct_Init(Coll As Collection, first As Integer, last As Integer)
    'Initialisation d'une Collection INTEGER
    If first >= last Then Exit Sub
    For i As Integer = first To last
      Dim Valeur As Integer = i
      Dim Key As Integer = i
      Coll.Add(Valeur, CStr(Key))
    Next i
  End Sub
  Sub Clct_Add(Coll As Collection, Valeur As Integer)
    'Ajout d'une valeur INTEGER à une Collection INTEGER
    Dim Key As Integer = Coll.Count + 1
    Coll.Add(Valeur, CStr(Key))
  End Sub
  '17/03/2024 Pourra être utilisé dans Pzzl_Crt_Triplet et Pzzl_Crt_XWing
  Public Function Clct_Remove(Coll As Collection, Valeur As Object) As Integer
    For i As Integer = 1 To Coll.Count
      If Coll.Item(i).ToString() = Valeur.ToString() Then
        Coll.Remove(i)
        Jrn_Add(, {"La valeur " & CStr(Valeur) & " est supprimée; poste " & CStr(i) & "."})
        Return i
      End If
    Next
    Return -1
  End Function

  Public Function Clct_Random(Coll As Collection) As Integer
    ' Retourne une valeur INTEGER d'une Collection INTEGER
    ' Clct_Random permet de déterminer une valeur dans une collection de 1 à Coll.Count randomisée
    ' Si la collection est vide, -1 est retourné
    ' Si la collection comporte un élément, c'est celui-ci qui est retourné, et la collection est vidée
    ' Si la collection comporte n éléments, un élément au hasard est retourné et ENLEVÉ
    '   dans ce cas, la collection compte un élément en moins ensuite
    If Coll Is Nothing OrElse Coll.Count = 0 Then
      Return -1
    End If

    Dim index As Integer =
        If(Coll.Count = 1,
           1,
           Rd9.Next(1, Coll.Count + 1))

    Dim valeur As Integer = CInt(Coll.Item(index))
    Coll.Remove(index)
    Return valeur
  End Function

  Sub Clct_Display(Coll As Collection)
    'Affiche les éléments d'une Collection INTEGER
    Jrn_Add(, {"Liste des " & CStr(Coll.Count) & " poste(s) de la Collection."})
    For i As Integer = 1 To Coll.Count
      Dim valeur As Integer = CInt(Coll.Item(i))
      Jrn_Add(, {"Key " + CStr(i) + " Valeur " & CStr(Coll.Item(i))})
    Next i
  End Sub
  Sub List_Random()
    Dim n As Integer
    Dim S As String
    Dim Cellules As List(Of Integer) = Enumerable.Range(0, 81).OrderBy(Function(f) Rde.Next()).ToList()
    For Each Cellule As Integer In Cellules
      Select Case Cellule
        Case 0 : S = "/"
        Case 80 : S = "\"
        Case Else : S = " "
      End Select
      n += 1
      Jrn_Add("SDK_00000", {CStr(n).PadLeft(2) & "/" & CStr(Cellules.Count) & " " & CStr(Cellule).PadLeft(2) & " " & U_cr(Cellule) & " " & S})
    Next Cellule
  End Sub
End Module
