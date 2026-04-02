'-------------------------------------------------------------------------------------------
' Gbl G bi-locaux
' Mise en place 20/12/2025 
'-------------------------------------------------------------------------------------------
Module G050_Strategy_bi_locaux
  Public Sub Strategy_Gbl(U_temp(,) As String)
    ' Cette procédure affiche les liens forts pour supprimer Les Candidats Bloqués
    Plcy_Strg = "Gbl"
    If Xap Then Jrn_Add(, {Proc_Name_Get() & " Xap = True "})

    GRslt_Init()

    ' 21 Inventaire des Liens forts
    GLinks_Build(U_temp, "0")
    GRslt.Nb_Liens = GLinks.Count
    GLinks_OrderBy()
    If Xap Then GLinks_Display()

    ' 22 Analyse des liens
    For Each gLink As GLink_Cls In GLinks
      ' Les 2 candidats doivent être dans la même région et appartenir à la même ligne ou à la même colonne
      If U_Reg(gLink.Cel(0)) <> U_Reg(gLink.Cel(1)) Then Continue For
      If (U_Row(gLink.Cel(0)) = U_Row(gLink.Cel(1)) Or U_Col(gLink.Cel(0)) = U_Col(gLink.Cel(1))) Then
        Dim Reg As Integer = U_Reg(gLink.Cel(0))
        Strategy_Gbl_Productivité(U_temp, gLink, U_9CelReg(Reg))
      End If

      If U_Row(gLink.Cel(0)) = U_Row(gLink.Cel(1)) Then
        Dim Row As Integer = U_Row(gLink.Cel(0))
        Strategy_Gbl_Productivité(U_temp, gLink, U_9CelRow(Row))
      End If

      If U_Col(gLink.Cel(0)) = U_Col(gLink.Cel(1)) Then
        Dim Col As Integer = U_Col(gLink.Cel(0))
        Strategy_Gbl_Productivité(U_temp, gLink, U_9CelCol(Col))
      End If
    Next gLink

    If GRslt.CelExcl.Count > 0 Then GRslt.Productivité = True
    Stratégies_G_End()
  End Sub
  Public Function Strategy_Gbl_Productivité(U_temp(,) As String, gLink As GLink_Cls, Grp() As Integer) As Boolean
    Dim gLink_Productif As Boolean = False
    For r As Integer = 0 To 8
      If Grp(r) = gLink.Cel(0) Or Grp(r) = gLink.Cel(1) Then Continue For
      If U_temp(Grp(r), 2) <> " " Then Continue For
      If U_temp(Grp(r), 3).Contains(gLink.Cdd(0)) Then
        gLink_Productif = True
        GRslt.CelExcl.Add(New GCel_Excl_Cls With {.Cel = Grp(r), .Cdd = gLink.Cdd(0), .Exc = {gLink.Cel(0), gLink.Cel(1)}})
        Dim CelExcl_key As Tuple(Of Integer, String) = Tuple.Create(Grp(r), gLink.Cdd(0))
        GRslt.CelExcl_hs.Add(CelExcl_key)
      End If
    Next r
    If gLink_Productif = True Then GRslt.RoadRight.Add(gLink)
    Return gLink_Productif
  End Function
End Module