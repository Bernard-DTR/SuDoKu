'-------------------------------------------------------------------------------------------
' Jeudi 14/08/2025, Lundi 25/08/2025
' Stratégie Wing_XYZ
' Préfixe WgW, 
' Les stratégies sont ou seront X, XY, XYZ et W, soit WgX, WgY, WgZ et WgW
'
' 6......7.....2.....58..1.....4.....2...1..9.6597.4...892.4..6.....5.......68.3...
'
'-------------------------------------------------------------------------------------------

Friend Module Q46_Strategy_W_Wing

  Public Sub Strategy_WgW(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "WgW" 
    XRslt_Init()

    ' 2  Inventaire des Cellules Bivalues 
    GCels.Clear()
    XCels_List_Get(U_temp, 2)
    If Xap Then XCels_List_Display()
    GCels = GCels.OrderBy(Function(s) s.Cdd(0)).ThenBy(Function(s) s.Cdd(1)).ToList()

    ' 21 Inventaire des Liens forts
    XLinks_List.Clear()
    XLinks_Cdd_Get_WgW(U_temp, "0")
    XLinks_List = XLinks_List.OrderBy(Function(s) s.Cdd(4)).ThenBy(Function(s) s.Cel(0)).ThenBy(Function(s) s.Cel(1)).ToList() 'par ordre croissant

    ' 3  Parcourir toutes les paires de bivalues différentes
    For i As Integer = 0 To GCels.Count - 2
      Dim Celi As Integer = GCels(i).Cel
      Dim Cddi() As String = {GCels(i).Cdd(0), GCels(i).Cdd(1)}
      XRslt.RoadRight.Clear()
      For j As Integer = i + 1 To GCels.Count - 1
        Dim Celj As Integer = GCels(j).Cel
        Dim Cddj() As String = {GCels(j).Cdd(0), GCels(j).Cdd(1)}
        XRslt.Cellule = {Celi, Celj}

        ' Condition:
        ' 1 les 2 cellules bivalues ne sont pas dans la même unité
        ' 2 Les 2 candidats sont identiques Cddi(0) = Cddj(0) AndAlso Cddi(1) = Cddj(1)
        If Is_SameUnité(Celi, Celj)(0) <> "#" Then Continue For
        If Cddi(0) <> Cddj(0) OrElse Cddi(1) <> Cddj(1) Then Continue For
        ' Traitement avec Candidat0 et Candidat1
        XRslt.RoadRight.Clear()
        XRslt.Candidat = {Cddi(0), Cddi(1)}
        Is_XLink_Exist_WgW(Cddi(0), Celi, Celj)
        If XRslt.RoadRight.Count > 0 Then
          If Candidats_Exclure_WgW(U_temp, Cddi(1), Celi, Celj) > 0 Then XRslt.Productivité = True
        End If
        If XRslt.Productivité Then Exit For

        ' Traitement avec Candidat1 et Candidat0
        XRslt.RoadRight.Clear()
        XRslt.Candidat = {Cddi(1), Cddi(0)}
        Is_XLink_Exist_WgW(Cddi(1), Celi, Celj)
        If XRslt.RoadRight.Count > 0 Then
          If Candidats_Exclure_WgW(U_temp, Cddi(0), Celi, Celj) > 0 Then XRslt.Productivité = True
        End If
        If XRslt.Productivité Then Exit For
      Next j
      If XRslt.Productivité Then Exit For
    Next i
    Stratégies_G_End()

  End Sub

  Function Candidats_Exclure_WgW(U_temp(,) As String, candidat As String, cel0 As Integer, cel1 As Integer) As Integer
    XRslt.CelExcl.Clear()
    For i As Integer = 0 To 80
      If i = cel0 Or i = cel1 Then Continue For
      If i = XRslt.Cellule(0) Or i = XRslt.Cellule(1) Then Continue For
      If Not U_temp(i, 3).Contains(candidat) Then Continue For
      If Is_Vu(i, cel0) AndAlso Is_Vu(i, cel1) Then
        XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = i, .Cdd = candidat, .Exc = {cel0, cel1}})
      End If
    Next
    Return XRslt.CelExcl.Count
  End Function

  Public Sub XLinks_Cdd_Get_WgW(U_temp(,) As String, Candidat As String)
    ' Ajout des liens forts sur 1 ou tous les candidats
    ' La valeur de Candidat permet de filtrer le candidat à traiter
    ' La valeur 0 ne filtre aucun candidat
    Dim Code_LCRs As String() = {"Row", "Col", "Reg"}
    For Each Code_LCR As String In Code_LCRs
      For LCR As Integer = 0 To 8
        Dim Grp As Integer() = Grp_Get(Code_LCR, LCR)
        For Cdd As Integer = 1 To 9
          If CInt(Candidat) = 0 OrElse Cdd = CInt(Candidat) Then
            Dim MatchingCells As List(Of Integer) = CddCel_Find(Grp, Cdd, U_temp)
            If MatchingCells.Count = 2 Then
              Dim Cel1 As Integer = MatchingCells(0)
              Dim Cel2 As Integer = MatchingCells(1)
              If Code_LCR <> "Reg" OrElse (U_Row(Cel1) <> U_Row(Cel2) AndAlso U_Col(Cel1) <> U_Col(Cel2)) Then
                Dim Link As New XLink_Cls With {.Cel = {Cel1, Cel2}, .Cdd = {CStr(Cdd), "0", CStr(Cdd), "0", CStr(Cdd)}, .Type = "S", .Unité = Code_LCR & (LCR + 1).ToString()}
                XLinks_List.Add(Link)
              End If
            End If
          End If
        Next
      Next
    Next
  End Sub

  Public Sub Is_XLink_Exist_WgW(Candidat As String, Cel0 As Integer, Cel1 As Integer)
    Dim Ajout As Boolean = False
    For Each link As XLink_Cls In XLinks_List
      If link.Cdd(4) = Candidat Then
        Dim L0 As Integer = link.Cel(0)
        Dim L1 As Integer = link.Cel(1)
        If L0 = Cel0 Or L1 = Cel1 Or L0 = Cel1 Or L1 = Cel0 Then Continue For
        If (Is_Vu(L0, Cel0) AndAlso Is_Vu(L1, Cel1)) OrElse (Is_Vu(L0, Cel1) AndAlso Is_Vu(L1, Cel0)) Then

          If Is_Vu(L0, Cel0) Then
            Dim Unitéw As String() = Is_SameUnité(L0, Cel0)
            Dim Linkw As New XLink_Cls With {.Cel = {Cel0, L0}, .Cdd = {Candidat, "0", Candidat, "0", Candidat}, .Type = "W", .Unité = Unitéw(0) & Unitéw(1)}
            XRslt.RoadRight.Add(Linkw)
          End If
          If Is_Vu(L1, Cel1) Then
            Dim Unitéw As String() = Is_SameUnité(L1, Cel1)
            Dim Linkw As New XLink_Cls With {.Cel = {Cel1, L1}, .Cdd = {Candidat, "0", Candidat, "0", Candidat}, .Type = "W", .Unité = Unitéw(0) & Unitéw(1)}
            XRslt.RoadRight.Add(Linkw)
          End If

          XRslt.RoadRight.Add(link)
          Ajout = True

          If Is_Vu(L0, Cel1) Then
            Dim Unitéw As String() = Is_SameUnité(L0, Cel1)
            Dim Linkw As New XLink_Cls With {.Cel = {Cel1, L0}, .Cdd = {Candidat, "0", Candidat, "0", Candidat}, .Type = "W", .Unité = Unitéw(0) & Unitéw(1)}
            XRslt.RoadRight.Add(Linkw)
          End If

          If Is_Vu(L1, Cel0) Then
            Dim Unitéw As String() = Is_SameUnité(L1, Cel0)
            Dim Linkw As New XLink_Cls With {.Cel = {Cel0, L1}, .Cdd = {Candidat, "0", Candidat, "0", Candidat}, .Type = "W", .Unité = Unitéw(0) & Unitéw(1)}
            XRslt.RoadRight.Add(Linkw)
          End If

        End If
      End If
      If Ajout Then Exit For
    Next
  End Sub

  Public Function Grp_Get(Code_LCR As String, LCR As Integer) As Integer()
    Select Case Code_LCR
      Case "Row" : Return U_9CelRow(LCR)
      Case "Col" : Return U_9CelCol(LCR)
      Case "Reg" : Return U_9CelReg(LCR)
      Case Else : Throw New ArgumentException("Code LCR invalide")
    End Select
  End Function

  Public Function CddCel_Find(Grp As Integer(), Cdd As Integer, U_temp(,) As String) As List(Of Integer)
    Dim Result As New List(Of Integer)
    For Each Cel As Integer In Grp
      If U_temp(Cel, 3).Contains(Cdd.ToString()) Then
        Result.Add(Cel)
        If Result.Count > 2 Then Exit For
      End If
    Next
    Return Result
  End Function

End Module