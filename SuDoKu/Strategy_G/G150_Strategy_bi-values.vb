Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
'
' Gbv G bi-values
' Mise en place 24/12/2025 
'-------------------------------------------------------------------------------------------

Module G100_Strategy_bi_values

  Public Sub Strategy_Gbv(U_temp(,) As String)
    ' Cette procédure expérimentale affiche les liens forts bi-values
    ' 2 cellules de 2 candidats identiques dans la même unité
    '8..21...7.4..891......4........98..5.9..3......7..14.9...........3...2542.......8
    Plcy_Strg = "Gbv"
    If Xap Then Jrn_Add(, {Proc_Name_Get() & " Xap = True "})

    GRslt_Init()

    ' 2 Inventaire des Cellules Bi-values triées sur les candidats 
    GCels.Clear()
    GCels_bv_Build(U_temp)
    GCels_bv_OrderBy()
    If Xap Then GCels_Display()

    ' 3 Création des liens forts, il faut au moins 2 cellules bivalues pour faire un lien
    '   Les 2 boucles ne sont pas redondantes

    GRslt.RoadRight.Clear()
    GRslt.CelExcl.Clear()
    GRslt.Productivité = False

    For i As Integer = 0 To GCels.Count - 2
      Dim Celi As Integer = GCels(i).Cel
      Dim Cddi() As String = {GCels(i).Cdd(0), GCels(i).Cdd(1)}
      For j As Integer = i + 1 To GCels.Count - 1
        Dim Celj As Integer = GCels(j).Cel
        Dim Cddj() As String = {GCels(j).Cdd(0), GCels(j).Cdd(1)}
        ' Condition:
        ' 1 Les 2 candidats sont identiques Cddi(0) = Cddj(0) AndAlso Cddi(1) = Cddj(1)
        If Cddi(0) <> Cddj(0) OrElse Cddi(1) <> Cddj(1) Then Continue For

        ' 2 les 2 cellules bivalues     sont dans la même Unité
        If U_Row(Celi) = U_Row(Celj) Then
          ' 21 On teste si les 2 cellules sont dans la même Ligne   et il y a des candidats à enlever
          Dim row As Integer = U_Row(Celi)
          Strategy_Gbv_Productif(U_temp, Celi, Celj, Cddi, Cddj,
                "Row" & (row + 1),
                U_9CelRow(row),
                Function(c) True,   ' aucun filtre
                GRslt)
        End If

        If U_Col(Celi) = U_Col(Celj) Then
          ' 22 On teste si les 2 cellules sont dans la même Colonne et il y a des candidats à enlever
          Dim col As Integer = U_Col(Celi)
          Strategy_Gbv_Productif(U_temp, Celi, Celj, Cddi, Cddj,
                "Col" & (col + 1),
                U_9CelCol(col),
                Function(c) True,
                GRslt)
        End If

        If U_Reg(Celi) = U_Reg(Celj) Then
          ' 23 On teste si les 2 cellules sont dans la même Région  et il y a des candidats à enlever
          Dim reg As Integer = U_Reg(Celi)

          Strategy_Gbv_Productif(U_temp, Celi, Celj, Cddi, Cddj,
                "Reg" & (reg + 1),
                U_9CelReg(reg),
                Function(c)
                  ' Exclusions spécifiques à la région
                  If U_Row(c) = U_Row(Celi) AndAlso U_Row(c) = U_Row(Celj) Then Return False
                  If U_Col(c) = U_Col(Celi) AndAlso U_Col(c) = U_Col(Celj) Then Return False
                  Return True
                End Function,
                GRslt)
        End If
      Next j
    Next i
    If GRslt.CelExcl.Count > 0 Then GRslt.Productivité = True
    GRslt.Nb_Liens = GRslt.RoadRight.Count
    Stratégies_G_End()
  End Sub
  Private Sub Strategy_Gbv_Productif(ByVal U_temp(,) As String,
          ByVal Celi As Integer,
          ByVal Celj As Integer,
          ByVal Cddi() As String,
          ByVal Cddj() As String,
          ByVal UnitéNom As String,
          ByVal Grp As Integer(),
          ByVal Filtre As Func(Of Integer, Boolean),
          ByRef GRslt As GRslt_Struct)

    Dim Link_Productif As Boolean = False

    For Each cel As Integer In Grp
      If cel = Celi Or cel = Celj Then Continue For
      If Not Filtre(cel) Then Continue For
      If U_temp(cel, 2) <> " " Then Continue For

      ' Test des deux candidats
      For Each cdd As String In Cddi
        If U_temp(cel, 3).Contains(cdd) Then
          Link_Productif = True
          GRslt.CelExcl.Add(New GCel_Excl_Cls With {.Cel = cel, .Cdd = cdd, .Exc = {Celi, Celj}})
          Dim CelExcl_key As Tuple(Of Integer, String) = Tuple.Create(cel, cdd)
          GRslt.CelExcl_hs.Add(CelExcl_key)
        End If
      Next
    Next

    If Link_Productif Then
      GRslt.RoadRight.Add(New GLink_Cls With {
          .Cel = {Celi, Celj},
          .Cdd = {Cddi(0), Cddi(1), Cddj(0), Cddj(1), "0"},
          .Type = "S",
          .Unité = UnitéNom,
          .Cdd_Composition = "0123"
      })
    End If
  End Sub
End Module