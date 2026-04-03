'-------------------------------------------------------------------------------------------
' Mardi 19/08/2025
' Stratégie Wing_XY
'  
' Préfixe WgY, 
' Les stratégies sont ou seront X, XY, XYZ et W, soit WgX, WgY, WgZ et WgW
'
'.54.173..2.84.6.............9.....45..6.4.81.........2...2.9.........76..3.7.....
'-------------------------------------------------------------------------------------------

Friend Module Q080_Strategy_Wing_XY

  Public Sub Strategy_WgY(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})
    Plcy_Strg = "WgY"
    ' 1 Initialisation de XRslt avec Plcy_Strg = "WgY" 
    XRslt_Init()

    ' 2 Inventaire des cellules Bivalues, présentées par ordre des cellules
    GCels.Clear()
    XCels_List_Get(U_temp, 2)
    If Xap Then XCels_List_Display()

    ' 3 Pour chaque cellule XY, on recherche tous les liens XZ et YZ avec chacun des 2 candidats
    For Each Cell_BV As GCel_Cls In GCels
      Dim CddX As String = Cell_BV.Cdd(0)
      Dim CddY As String = Cell_BV.Cdd(1)
      XLinks_List.Clear()
      XLinks_List_Generate_WgY(U_temp, Cell_BV.Cel, CddX, "Row") ' Analyse des rangées
      XLinks_List_Generate_WgY(U_temp, Cell_BV.Cel, CddY, "Row") ' Analyse des rangées
      XLinks_List_Generate_WgY(U_temp, Cell_BV.Cel, CddX, "Col") ' Analyse des colonnes
      XLinks_List_Generate_WgY(U_temp, Cell_BV.Cel, CddY, "Col") ' Analyse des régions 
      XLinks_List_Generate_WgY(U_temp, Cell_BV.Cel, CddX, "Reg") ' Analyse des colonnes
      XLinks_List_Generate_WgY(U_temp, Cell_BV.Cel, CddY, "Reg") ' Analyse des régions 
      If Xap Then XLinks_List_Display("1")

      ' 4 Dans cette liste on regarde si 2 liens XZ et YZ ont le même candidat Z
      If XLinks_List.Count < 2 Then Continue For
      For i As Integer = 0 To XLinks_List.Count - 2
        Dim LinkXZ As XLink_Cls = XLinks_List(i)
        XRslt.RoadRight.Clear()

        For j As Integer = i + 1 To XLinks_List.Count - 1
          Dim LinkYZ As XLink_Cls = XLinks_List(j)
          XRslt.RoadRight.Clear()
          If LinkXZ.Unité = LinkYZ.Unité Then Continue For
          If LinkXZ.Cdd(2) = LinkYZ.Cdd(2) AndAlso LinkXZ.Cdd(3) = LinkYZ.Cdd(3) Then Continue For
          If LinkXZ.Cdd(3) = LinkYZ.Cdd(3) Then
            Dim CddZ As String = LinkXZ.Cdd(3)
            XRslt.RoadRight.Add(LinkXZ)
            XRslt.RoadRight.Add(LinkYZ)

            ' Existe-t'il des cellules vues de ces 2 extrémités comportant le candidat Z ?
            For k As Integer = 0 To 80
              If Not U_temp(k, 3).Contains(CddZ) Then Continue For
              If k = LinkXZ.Cel(1) Or k = LinkYZ.Cel(1) Then Continue For
              If Is_Vu(k, LinkXZ.Cel(1)) AndAlso Is_Vu(k, LinkYZ.Cel(1)) Then
                XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = k, .Cdd = CddZ, .Exc = {LinkXZ.Cel(1), LinkYZ.Cel(1)}})
              End If
            Next k

            If XRslt.CelExcl.Count > 0 Then
              XRslt.Candidat(0) = CddZ
              XRslt.Productivité = True
            End If
          End If
          If XRslt.Productivité Then Exit For
        Next j
        If XRslt.Productivité Then Exit For
      Next i
      If XRslt.Productivité Then Exit For
    Next Cell_BV

    Stratégies_G_End()
  End Sub

  Public Sub XLinks_List_Generate_WgY(U_temp(,) As String, Cellule As Integer, Candidat As String, Code_LCR As String)
    ' On recherche pour une Cellule et un Candidat un lien dans la Row, Col, Reg sur une cellule à 2 candidats

    Dim Grp(0 To 8) As Integer
    Dim Unité As String = ""
    Select Case Code_LCR
      Case "Row" : Grp = U_9CelRow(U_Row(Cellule)) : Unité = "Row" & U_Row(Cellule) + 1
      Case "Col" : Grp = U_9CelCol(U_Col(Cellule)) : Unité = "Col" & U_Col(Cellule) + 1
      Case "Reg" : Grp = U_9CelReg(U_Reg(Cellule)) : Unité = "Reg" & U_Reg(Cellule) + 1
    End Select

    Dim CddA1 As String
    Dim CddA2 As String
    Dim Candidats As String
    Candidats = U_temp(Cellule, 3).Replace(" ", "")
    CddA1 = Candidats(0) : CddA2 = Candidats(1)
    Dim CddB1 As String
    Dim CddB2 As String

    For Each Cel As Integer In Grp
      If Cel = Cellule Then Continue For
      If Wh_Cell_Nb_Candidats(U_temp, Cel) <> 2 Then Continue For
      If U_temp(Cellule, 3) = U_temp(Cel, 3) Then Continue For
      If Not U_temp(Cel, 3).Contains(Candidat) Then Continue For
      Candidats = U_temp(Cel, 3).Replace(" ", "")

      If Candidats(0) = Candidat Then
        CddB1 = Candidats(0) : CddB2 = Candidats(1)
      Else
        CddB1 = Candidats(1) : CddB2 = Candidats(0)
      End If

      Dim Link As New XLink_Cls With {.Cel = {Cellule, Cel}, .Cdd = {CddA1, CddA2, CddB1, CddB2, Candidat}, .Type = "W", .Unité = Unité}
      Select Case Code_LCR
        Case "Row", "Col" : XLinks_List.Add(Link)
        Case "Reg" ' Il ne faut pas que les 2 cellules soient dans la même ligne / Colonne
          If U_Row(Cellule) <> U_Row(Cel) And U_Col(Cellule) <> U_Col(Cel) Then XLinks_List.Add(Link)
      End Select

    Next Cel
  End Sub

End Module