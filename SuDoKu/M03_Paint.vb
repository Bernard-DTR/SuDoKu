Imports System.Drawing.Drawing2D
Friend Module M03_Paint
  '-------------------------------------------------------------------------------
  ' 20/09/2022 L'ensemble des dessins sont faits à l'intérieur de Sqr_Cel
  '            avec Top-Left + 1 et Width-Height - 3
  '-------------------------------------------------------------------------------   
#Region "G4 Couche Stratégie"
  '   La couche G4 stratégies n'est appelée que dans G4_Grid_Stratégie_All,
  Public Sub G4_Grid_Stratégie_All(g As Graphics)
    If Plcy_Gnrl = "Nrm" And Plcy_Strg <> "   " Then

      For i As Integer = 0 To 80
        U_Strg_Val_Ins(i) = ""
        U_Strg_Cdd_Exc(i) = Cnddts_Blancs
      Next i
      G4_Grid_Stratégie_DCd(g)
      G4_Grid_Stratégie_Cdd(g)
      G4_Grid_Stratégie_CdU(g)
      G4_Grid_Stratégie_CdO(g)
      G4_Grid_Stratégie_Flt(g)
      G4_Grid_Stratégie_Cbl(g)
      G4_Grid_Stratégie_Tpl(g)
      G4_Grid_Stratégie_Xwg(g)
      G4_Grid_Stratégie_XYw(g)
      G4_Grid_Stratégie_Swf(g)
      G4_Grid_Stratégie_Jly(g)
      G4_Grid_Stratégie_XYZ(g)
      G4_Grid_Stratégie_SKy(g)
      G4_Grid_Stratégie_Unq(g)
      G4_Grid_Stratégie_Obj(g)
      G4_Grid_Stratégie_GCx(g)
      G4_Grid_Stratégie_XCy_XNl(g)
      G4_Grid_Stratégie_XRp(g)
      G4_Grid_Stratégie_WgX_WgY_WgZ_WgW(g)
      G4_Grid_Stratégie_GLk(g)
      G4_Grid_Stratégie_Gbl(g)
      G4_Grid_Stratégie_Gbv(g)
      G4_Grid_Stratégie_GCs(g)
      G4_Grid_Stratégie_Animation(g)
    End If
  End Sub
  Public Sub G4_Grid_Stratégie_Animation(g As Graphics)
    If Not Plcy_Strg = "Ani" Then Exit Sub
    Dim Cellule_Clct As New Collection

    Dim cellule As Integer
    'Collection des valeurs initiales
    For i As Integer = 0 To 80
      If U(i, 1) <> " " Then Clct_Add(Cellule_Clct, i)
    Next i
    ' Animation de la Grille
    Dim Rect As Rectangle
    Dim Inflate As Integer = 0
    For i As Integer = 1 To Cellule_Clct.Count
      cellule = Clct_Random(Cellule_Clct)
      Rect = Sqr_Cel(cellule)
      Inflate += 2
      If Inflate > WH \ 2 Then Exit For
      Rect.Inflate(New Size(Inflate - 2, Inflate - 2))
      g.DrawIcon(My.Resources.SuDoKu, Rect)
      Threading.Thread.Sleep(100)
      Rect.Inflate(New Size(Inflate, Inflate))
      g.DrawIcon(My.Resources.SuDoKu, Rect)
      Threading.Thread.Sleep(100)
      Rect.Inflate(New Size(Inflate + 1, Inflate + 1))
      g.DrawIcon(My.Resources.SuDoKu, Rect)
      Threading.Thread.Sleep(100)
    Next i

    Plcy_Strg = "   "
    Frm_SDK.Invalidate()
  End Sub
  Public Sub G4_Grid_Stratégie_DCd(g As Graphics)
    If Plcy_Strg <> "DCd" Then Exit Sub

    ' Récupérer les indices des cellules concernées
    Dim indices As List(Of Integer) = Enumerable.Range(0, 81).
                             Where(Function(i) U(i, 2) = Pbl_Valeur_CdS).
                             ToList()

    Dim figure As String =
        If(indices.Count = 9, "Cercle", "Double_Carré")

    For Each i As Integer In indices
      G0_Cell_Figure(g, i, figure, Color_Stratégique)
    Next
  End Sub
  Public Sub G4_Grid_Stratégie_DCd_2(g As Graphics)
    If Not Plcy_Strg = "DCd" Then Exit Sub
    Dim n As Integer
    For i As Integer = 0 To 80
      If U(i, 2) = Pbl_Valeur_CdS Then n += 1
    Next i


    For i As Integer = 0 To 80
      If U(i, 2) = Pbl_Valeur_CdS Then
        If n = 9 Then
          G0_Cell_Figure(g, i, "Cercle", Color_Stratégique)
        Else
          G0_Cell_Figure(g, i, "Double_Carré", Color_Stratégique)
        End If
      End If
    Next i
  End Sub
  Public Sub G4_Grid_Stratégie_Cdd(g As Graphics)
    If Not Plcy_Strg = "Cdd" Then Exit Sub
    For i As Integer = 0 To 80
      If U(i, 2) = " " Then
        Dim sc_a As New Cellule_Cls With {.Numéro = i}
        sc_a.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      End If
    Next i
  End Sub
  Public Sub G4_Grid_Stratégie_CdU(g As Graphics)
    ' La stratégie CdU calcule TOUS les CdU, UN SEUL CdU au hasard est présenté 
    Dim Cellule As Integer
    Dim Candidat As String
    If Not Plcy_Strg = "CdU" Then Exit Sub
    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Cellule = RRslt.Cellule(0)
    Candidat = RRslt.Candidat
    U_Strg_Val_Ins(Cellule) = Candidat
    G0_Cell_Figure(g, Cellule, "Double_Carré", Color_Stratégique)

    U_MdC_Clear()
    G4_MdC_Row_Col_Box("Row", U_Row(Cellule))
    G4_MdC_Row_Col_Box("Col", U_Col(Cellule))
    G4_MdC_Row_Col_Box("Box", U_Reg(Cellule))
    G4_MdC_Paint(g) ' Les figures sont dessinées et les candidats affichés
    'Re-dessine le candidat à placer dans un cercle plein Jaune
    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Insérer)
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & Candidat & " jaune à placer."

  End Sub
  Public Sub G4_Grid_Stratégie_CdO(g As Graphics)
    Dim Cellule As Integer
    Dim Candidat As String
    If Not Plcy_Strg = "CdO" Then Exit Sub

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Cellule = RRslt.Cellule(0)
    Candidat = RRslt.Candidat
    U_Strg_Val_Ins(Cellule) = Candidat
    G0_Cell_Figure(g, Cellule, "Double_Carré", Color_Stratégique)

    Dim Code_LCR As String = RRslt.Code_LCR
    Dim LCR As Integer = RRslt.LCR
    U_MdC_Clear()
    Select Case Code_LCR
      Case "L" : G4_MdC_Row_Col_Box("Row", LCR)
      Case "C" : G4_MdC_Row_Col_Box("Col", LCR)
      Case "R" : G4_MdC_Row_Col_Box("Box", LCR)
    End Select
    G4_MdC_Paint(g) ' Les figures sont dessinées et les candidats affichés
    'Re-dessine le candidat à placer dans un cercle plein Jaune
    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Insérer)
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & Candidat & " jaune à placer."

  End Sub
  Public Sub G4_Grid_Stratégie_Flt(g As Graphics)
    ' Affichage des Valeurs Filtrés
    If Mid$(Plcy_Strg, 1, 2) = "FV" Then
      Dim Valeur_Filtrée As String = Mid$(Plcy_Strg, 3, 1)
      For i As Integer = 0 To 80
        If U(i, 2) = Valeur_Filtrée Then
          G0_Cell_Figure(g, i, "Double_Carré", Color_Stratégique)
        End If
      Next i
      MW_Prv_Val = CInt(Valeur_Filtrée)
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte
    End If

    ' Affichage des Candidats Filtrés
    If Mid$(Plcy_Strg, 1, 2) = "FC" Then
      Dim Candidat As String = Mid$(Plcy_Strg, 3, 1)
      Dim Color As Color = Color_BySymbol(Obj_Symbol)
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(Candidat) Then sc.G6_Cellule_Paint_Candidat(g, Candidat, Color)
      Next i
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte
    End If
  End Sub
  Public Sub G4_Grid_Stratégie_Cbl(g As Graphics)
    If Not Plcy_Strg = "Cbl" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next

    Dim Code_LCR As String = RRslt.Code_LCR
    Dim LCR As Integer = RRslt.LCR
    U_MdC_Clear()
    Select Case Code_LCR
      Case "L" : G4_MdC_Row_Col_Box("Row", LCR)
      Case "C" : G4_MdC_Row_Col_Box("Col", LCR)
      Case "R" : G4_MdC_Row_Col_Box("Box", LCR)
    End Select
    G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés

    Candidat = RRslt.Candidat
    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."

    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_Tpl(g As Graphics)
    If Not Plcy_Strg = "Tpl" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next

    Dim Code_LCR As String = RRslt.Code_LCR
    Dim LCR As Integer = RRslt.LCR
    U_MdC_Clear()
    Select Case Code_LCR
      Case "L" : G4_MdC_Row_Col_Box("Row", LCR)
      Case "C" : G4_MdC_Row_Col_Box("Col", LCR)
      Case "R" : G4_MdC_Row_Col_Box("Box", LCR)
    End Select
    G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."

    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_Xwg(g As Graphics)
    If Not Plcy_Strg = "Xwg" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next

    With RRslt
      U_MdC_Clear()
      Select Case .Code_Sous_Strg
        Case "Xwg_C"
          G4_MdC_Row_Col_Box("Row", U_Row(.Cellule(0)))
          G4_MdC_Row_Col_Box("Row", U_Row(.Cellule(2)))
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(3))
        Case "Xwg_L"
          G4_MdC_Row_Col_Box("Col", U_Col(.Cellule(0)))
          G4_MdC_Row_Col_Box("Col", U_Col(.Cellule(2)))
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(3))
        Case "Shm_L", "Shm_C"
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(1))
          G4_MdC_Trait_ou_Rectangle(.Cellule(2), .Cellule(4))
        Case "Fnd_La", "Fnd_LbL", "Fnd_LbR", "Fnd_Lc", "Fnd_Ca", "Fnd_CbT", "Fnd_CbB", "Fnd_Cc"
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(1))
          G4_MdC_Trait_ou_Rectangle(.Cellule(2), .Cellule(4))
          'L'aileron est présenté par une diagonale de Bresenham
          G0_Cell_Diagonale(g, .Cellule(5), .Cellule(6))
      End Select
      G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés
    End With

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next

    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."

    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_XYw(g As Graphics)
    If Not Plcy_Strg = "XYw" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next
    For Each cellexcl As Integer In RRslt.CelExcl
      G0_Cell_Figure(g, cellexcl, "Double_Carré", Color_Stratégique)
    Next

    With RRslt
      U_MdC_Clear()
      Select Case .Code_Sous_Strg(1).ToString()
        Case "C"
          G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(0)))
          G4_MdC_Row_Col_Box("Row", U_Row(.Cellule(0)))
        Case "F"
          G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(0)))
          G4_MdC_Row_Col_Box("Col", U_Col(.Cellule(0)))
        Case Else
          G4_MdC_Row_Col_Box("Row", U_Row(.CelExcl(0)))
          G4_MdC_Row_Col_Box("Col", U_Col(.CelExcl(0)))
      End Select
      G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés
    End With

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next

    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."

    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_Swf(g As Graphics)
    If Not Plcy_Strg = "Swf" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next
    For Each cellexcl As Integer In RRslt.CelExcl
      G0_Cell_Figure(g, cellexcl, "Double_Carré", Color_Stratégique)
    Next

    With RRslt
      U_MdC_Clear()
      For Each cellule As Integer In RRslt.Cellule
        G4_MdC_Row_Col_Box("Row", U_Row(cellule))
        G4_MdC_Row_Col_Box("Col", U_Col(cellule))
      Next
      G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés
    End With

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next

    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."
    RRslt_Control_Cdd_Exclure(Candidat)

  End Sub
  Public Sub G4_Grid_Stratégie_Jly(g As Graphics)
    If Not Plcy_Strg = "Jly" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next
    For Each cellexcl As Integer In RRslt.CelExcl
      G0_Cell_Figure(g, cellexcl, "Double_Carré", Color_Stratégique)
    Next

    With RRslt
      U_MdC_Clear()
      For Each cellule As Integer In RRslt.Cellule
        G4_MdC_Row_Col_Box("Row", U_Row(cellule))
        G4_MdC_Row_Col_Box("Col", U_Col(cellule))
      Next
      G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés
    End With

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next

    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."
    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_XYZ(g As Graphics)
    If Not Plcy_Strg = "XYZ" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next
    For Each cellexcl As Integer In RRslt.CelExcl
      G0_Cell_Figure(g, cellexcl, "Double_Carré", Color_Stratégique)
    Next

    With RRslt
      U_MdC_Clear()
      G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(0)))
      G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés
    End With

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next

    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."
    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_SKy(g As Graphics)
    'La stratégie SKy regroupe 3 Sous-Stratégies:
    '   s/Stratégie   SKY Skyscraper Gratte-Ciel
    '   s/Stratégie   Kyt Kyte       Cerf-Volant
    '   s/Stratégie   EyR Empty Rct  Rectangle Vide

    If Not Plcy_Strg = "SKy" Then Exit Sub
    Dim Candidat As String

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If

    Candidat = RRslt.Candidat
    For Each cell As Integer In RRslt.Cellule
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      G0_Cdd_Figure(g, cell, CInt(Candidat), "Disque", Color_Stratégique)
    Next
    For Each cellexcl As Integer In RRslt.CelExcl
      G0_Cell_Figure(g, cellexcl, "Double_Carré", Color_Stratégique)
    Next

    With RRslt
      Select Case Mid(.Code_Sous_Strg, 1, 3)
        Case "SKy"
          U_MdC_Clear()
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(1))
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(2))
          G4_MdC_Trait_ou_Rectangle(.Cellule(1), .Cellule(3))
          G4_MdC_Paint(g)  ' Les figures sont dessinées et les candidats affichés
        Case "Kyt"
          U_MdC_Clear()
          G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(1))
          G4_MdC_Trait_ou_Rectangle(.Cellule(2), .Cellule(3))
          Select Case .Code_Sous_Strg
            Case "KytH1V1" : G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(2)))
            Case "KytH1V2" : G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(0)))
            Case "KytH2V1" : G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(2)))
            Case "KytH2V2" : G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(1)))
          End Select
          G4_MdC_Paint(g)      ' Les figures sont dessinées et les candidats affichés
        Case "EyR"
          U_MdC_Clear()
          G4_MdC_Row_Col_Box("Box", U_Reg(.Cellule(0)))
          G4_MdC_Trait_ou_Rectangle(.Cellule(1), .Cellule(2))
          G4_MdC_Paint(g)      ' Les figures sont dessinées et les candidats affichés
      End Select
    End With

    For Each cellexcl As Integer In RRslt.CelExcl
      Dim sc As New Cellule_Cls With {.Numéro = cellexcl}
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      'Re-dessine le candidat à placer dans un cercle plein rouge
      sc.G6_Cellule_Paint_Candidat(g, Candidat, Color_Cdd_Exclure)
      U_Strg_Cdd_Exc(cellexcl) = Candidat
    Next
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat & " rouge à enlever."
    RRslt_Control_Cdd_Exclure(Candidat)
  End Sub
  Public Sub G4_Grid_Stratégie_Unq(g As Graphics)
    If Not Plcy_Strg = "Unq" Then Exit Sub

    If RRslt.Productivité = False Then
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Exit Sub
    End If
    Dim Candidat As String = RRslt.Candidat
    Dim Candidats As String() = RRslt.Candidat.Where(Function(c) c <> " "c).
                                               Select(Function(c) c.ToString()).ToArray()

    For Each cell As Integer In RRslt.Cellule
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      G0_Cell_Figure(g, cell, "Double_Carré", Color_Stratégique)
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      For Each cdd As String In Candidats
        G0_Cdd_Figure(g, cell, CInt(cdd), "Disque", Color_Stratégique)
      Next
    Next
    For Each cellexcl As Integer In RRslt.CelExcl
      G0_Cell_Figure(g, cellexcl, "Double_Carré", Color_Stratégique)
    Next

    'Strategy_Unq_Rectangle_1 candidats comporte 2 candidats (placés dans un string 9)
    '                         la typologie est R11, R12, R13, R14 suivant le sens du rectangle
    '                         Les candidats sont correctement placés dans la zone Candidats
    'Strategy_Unq_Rectangle_2 candidats comporte 1 candidat  
    '                         la typologie est R2HD, R2HG, R2VH, R2VB
    '                         Le candidat est correctement placé dans la zone candidats (string 9)
    With RRslt

      U_MdC_Clear()
      G4_MdC_Trait_ou_Rectangle(.Cellule(0), .Cellule(2))
      G4_MdC_Paint(g)      ' Les figures sont dessinées et les candidats affichés

      For Each cellexcl As Integer In RRslt.CelExcl
        For Each cdd As String In Candidats
          If U(cellexcl, 3).Contains(cdd) Then
            U_Strg_Cdd_Exc(cellexcl) = cdd
            G0_Cdd_Figure(g, cellexcl, CInt(cdd), "Disque", Color_Cdd_Exclure)
          End If
        Next
      Next

    End With
    Dim Candidat_txt As String = String.Join(", ", Candidats)
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " (" & RRslt.Code_Sous_Strg & ") :   " & Candidat_txt & " rouge à enlever."

    For Each cdd As String In Candidats
      RRslt_Control_Cdd_Exclure(cdd)
    Next


  End Sub

  Public Sub G4_Grid_Stratégie_XCy_XNl(g As Graphics)
    If Not (Plcy_Strg = "XCy" Or Plcy_Strg = "XNl") Then Exit Sub
    Try
      Dim sc As New Cellule_Cls
      If XRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(XRslt.Candidat(0)) Then
          G0_Cdd_Figure(g, i, CInt(XRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        Select Case Plcy_Strg
          Case "XCy"
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(4)), Link.Cel(1), CInt(Link.Cdd(4)), Link.Type, Nb)
            Dim sc_XCy As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(4)) Then
              sc_XCy.G6_Cellule_Paint_Candidat(g, Link.Cdd(4), Color_Link_W)
            End If
            If U(Link.Cel(1), 3).Contains(Link.Cdd(4)) Then
              sc_XCy.Numéro = Link.Cel(1)
              sc_XCy.G6_Cellule_Paint_Candidat(g, Link.Cdd(4), Color_Link_W)
            End If

          Case "XNl"
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(4)), Link.Cel(1), CInt(Link.Cdd(4)), Link.Type, Nb)
            ' C'est une alternance lien-fort lien-faible 
            Select Case Nb Mod 2
              Case 0
                Dim sc_XNl As New Cellule_Cls With {.Numéro = Link.Cel(1)}
                sc_XNl.G6_Cellule_Paint_Candidat(g, CStr(CInt(Link.Cdd(4))), Color.Green)
              Case Else
                Dim sc_XNl As New Cellule_Cls With {.Numéro = Link.Cel(1)}
                sc_XNl.G6_Cellule_Paint_Candidat(g, CStr(CInt(Link.Cdd(4))), Color.Blue)
            End Select
        End Select

        ' 3 Affichage des Extrémités des liens  
        Dim PremierLien As XLink_Cls = XRslt.RoadRight.First()
        Dim DernierLien As XLink_Cls = XRslt.RoadRight.Last()
        G0_Cell_Icône(g, PremierLien.Cel(0), "Start")
        G0_Cell_Icône(g, DernierLien.Cel(1), "End")
        Select Case Plcy_Strg
          Case "XCy"
            G0_Cdd_Figure(g, PremierLien.Cel(0), CInt(PremierLien.Cdd(0)), "Disque", Color_Link_S)
            G0_Cdd_Figure(g, DernierLien.Cel(1), CInt(DernierLien.Cdd(1)), "Disque", Color_Link_S)
          Case "XNl"
            ' il n'y en a pas car c'est une boucle Arrivée = Départ
        End Select
      Next Link

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As XCel_Excl_Cls In XRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next XCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & XRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_XRp(g As Graphics)
    If Not Plcy_Strg = "XRp" Then Exit Sub
    Try
      Dim sc As New Cellule_Cls
      If XRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(XRslt.Candidat(0)) And U(i, 3).Contains(XRslt.Candidat(1)) Then
          G0_Cdd_Figure(g, i, CInt(XRslt.Candidat(0)), "Cercle", Color.White)
          G0_Cdd_Figure(g, i, CInt(XRslt.Candidat(1)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        Select Case Nb Mod 2 ' Pour alterner les liens entre les séries de candidats
          Case 0
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(0)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
            Dim sc_XRp As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(0)) Then
              sc_XRp.G6_Cellule_Paint_Candidat(g, Link.Cdd(0), Color_Link_W)
            End If
            sc_XRp.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(2)) Then
              sc_XRp.G6_Cellule_Paint_Candidat(g, Link.Cdd(2), Color_Link_W)
            End If

          Case 1
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(1)), Link.Cel(1), CInt(Link.Cdd(3)), Link.Type, Nb)
            Dim sc_XRp As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(1)) Then
              sc_XRp.G6_Cellule_Paint_Candidat(g, Link.Cdd(1), Color_Link_W)
            End If
            sc_XRp.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(3)) Then
              sc_XRp.G6_Cellule_Paint_Candidat(g, Link.Cdd(3), Color_Link_W)
            End If

        End Select
      Next Link

      ' 3 Affichage des Extrémités des liens  
      Dim PremierLien As XLink_Cls = XRslt.RoadRight.First()
      G0_Cell_Icône(g, PremierLien.Cel(0), "Start")
      G0_Cdd_Figure(g, PremierLien.Cel(0), CInt(XRslt.Candidat(0)), "Disque", Color_Link_W)
      G0_Cdd_Figure(g, PremierLien.Cel(0), CInt(XRslt.Candidat(1)), "Disque", Color_Link_W)

      Dim DernierLien As XLink_Cls = XRslt.RoadRight.Last()
      G0_Cell_Icône(g, DernierLien.Cel(1), "End")
      G0_Cdd_Figure(g, DernierLien.Cel(1), CInt(XRslt.Candidat(0)), "Disque", Color_Link_W)
      G0_Cdd_Figure(g, DernierLien.Cel(1), CInt(XRslt.Candidat(1)), "Disque", Color_Link_W)

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As XCel_Excl_Cls In XRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd           ' le candidat est ajouté
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next XCelExcl

      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & XRslt.Candidat(0) & " " & XRslt.Candidat(1) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub

  Public Sub G4_Grid_Stratégie_WgX_WgY_WgZ_WgW_Save(g As Graphics)
    If Not (Plcy_Strg = "WgX" Or Plcy_Strg = "WgY" Or Plcy_Strg = "WgZ" Or Plcy_Strg = "WgW") Then Exit Sub

    Try
      Dim sc As New Cellule_Cls
      If XRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(XRslt.Candidat(0)) Then
          G0_Cdd_Figure(g, i, CInt(XRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As XLink_Cls In XRslt.RoadRight
        Nb += 1
        Select Case Plcy_Strg
          Case "WgX"
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(0)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
            Dim sc_WgX As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(0)) Then
              sc_WgX.G6_Cellule_Paint_Candidat(g, Link.Cdd(0), Color_Link_W)
            End If
            sc_WgX.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(2)) Then
              sc_WgX.G6_Cellule_Paint_Candidat(g, Link.Cdd(2), Color_Link_W)
            End If
          Case "WgY"
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(2)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
            Dim sc_WgY As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(2)) Then
              sc_WgY.G6_Cellule_Paint_Candidat(g, Link.Cdd(2), Color_Link_W)
            End If
            sc_WgY.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(2)) Then
              sc_WgY.G6_Cellule_Paint_Candidat(g, Link.Cdd(2), Color_Link_W)
            End If
            G0_Cdd_Figure(g, Link.Cel(1), CInt(Link.Cdd(3)), "Disque", Color_Link_S) 'Mise en exergue du candidat Z
          Case "WgZ"
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(3)), Link.Cel(1), CInt(Link.Cdd(3)), Link.Type, Nb)
            Dim sc_WgZ As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(3)) Then
              sc_WgZ.G6_Cellule_Paint_Candidat(g, Link.Cdd(3), Color_Link_W)
            End If
            sc_WgZ.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(3)) Then
              sc_WgZ.G6_Cellule_Paint_Candidat(g, Link.Cdd(3), Color_Link_W)
            End If

            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(4)), Link.Cel(1), CInt(Link.Cdd(4)), Link.Type, Nb)
            'Dim sc_WgZ As New Cellule_Cls With {.Numéro = From_Cellule}
            sc_WgZ.Numéro = Link.Cel(0)
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(4)) Then
              sc_WgZ.G6_Cellule_Paint_Candidat(g, Link.Cdd(4), Color_Link_W)
            End If
            sc_WgZ.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(4)) Then
              sc_WgZ.G6_Cellule_Paint_Candidat(g, Link.Cdd(4), Color_Link_W)
            End If
            G0_Cdd_Figure(g, Link.Cel(0), CInt(Link.Cdd(5)), "Disque", Color_Link_S) 'Mise en exergue du candidat du pivot
          Case "WgW"
            G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(0)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
            Dim sc_WgW As New Cellule_Cls With {.Numéro = Link.Cel(0)}
            'Les candidats sont dessinés s'ils existent
            If U(Link.Cel(0), 3).Contains(Link.Cdd(0)) Then
              sc_WgW.G6_Cellule_Paint_Candidat(g, Link.Cdd(0), Color_Link_W)
            End If
            sc_WgW.Numéro = Link.Cel(1)
            If U(Link.Cel(1), 3).Contains(Link.Cdd(2)) Then
              sc_WgW.G6_Cellule_Paint_Candidat(g, Link.Cdd(2), Color_Link_W)
            End If
            G0_Cdd_Figure(g, XRslt.Cellule(0), CInt(XRslt.Candidat(1)), "Disque", Color_Link_S) 'Mise en exergue du candidat  
            G0_Cdd_Figure(g, XRslt.Cellule(1), CInt(XRslt.Candidat(1)), "Disque", Color_Link_S) 'Mise en exergue du candidat  
        End Select
      Next Link

      ' 3 Affichage des Extrémités des liens  
      Select Case Plcy_Strg
        Case "WgX"
          G0_Cell_Icône(g, XRslt.RoadRight.Item(0).Cel(0), "Start")
          G0_Cell_Icône(g, XRslt.RoadRight.Item(1).Cel(0), "Start")
        Case "WgY"
          G0_Cell_Icône(g, XRslt.RoadRight.Item(0).Cel(0), "Start")
        Case "WgZ"
          G0_Cell_Icône(g, XRslt.RoadRight.Item(0).Cel(0), "Start")
        Case "WgW"
          G0_Cell_Icône(g, XRslt.Cellule(0), "Start")
          G0_Cell_Icône(g, XRslt.Cellule(1), "Start")
      End Select

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As XCel_Excl_Cls In XRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next XCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & XRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_GLk(g As Graphics)
    If Not Plcy_Strg = "GLk" Then Exit Sub
    Try
      If GLinks.Count = 0 Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats dans la grille
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Liens Forts de la list GLinks
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GLinks
        Nb += 1
        G0_Cdd_Bézier(g, gLink.Cel(0), CInt(gLink.Cdd(0)), gLink.Cel(1), CInt(gLink.Cdd(2)), gLink.Type, Nb)
        Dim sc_gLink As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(0)) Then
          sc_gLink.G6_Cellule_Paint_Candidat(g, gLink.Cdd(0), Color_Link_S)
        End If
        sc_gLink.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(2)) Then
          sc_gLink.G6_Cellule_Paint_Candidat(g, gLink.Cdd(2), Color_Link_S)
        End If
      Next gLink
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " " & GRslt.Nb_Liens & " Liens Forts."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_Gbl(g As Graphics)
    If Not Plcy_Strg = "Gbl" Then Exit Sub

    Try
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage de toutes les Courbes de Bézier
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        Nb += 1
        G0_Cdd_Bézier(g, gLink.Cel(0), CInt(gLink.Cdd(0)), gLink.Cel(1), CInt(gLink.Cdd(2)), gLink.Type, Nb)
        Dim sc_gLink As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(0)) Then
          sc_gLink.G6_Cellule_Paint_Candidat(g, gLink.Cdd(0), Color_Link_S)
        End If
        sc_gLink.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(2)) Then
          sc_gLink.G6_Cellule_Paint_Candidat(g, gLink.Cdd(2), Color_Link_S)
        End If
      Next gLink

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each gCel As GCel_Excl_Cls In GRslt.CelExcl
        With gCel
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            If Pbl_Cell_Select = .Cel Then
              Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
              U_Strg_Cdd_Exc(.Cel) = Candidats
            End If
          End If
        End With
      Next gCel
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.CelExcl.Count & " candidat(s) rouge(s) à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub
  Public Sub G4_Grid_Stratégie_Gbv(g As Graphics)
    If Not Plcy_Strg = "Gbv" Then Exit Sub

    Try
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
      Next i

      ' 2 Affichage de toutes les Courbes de Bézier
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        Nb += 1
        ' Premier lien
        G0_Cdd_Bézier(g, gLink.Cel(0), CInt(gLink.Cdd(0)), gLink.Cel(1), CInt(gLink.Cdd(2)), gLink.Type, Nb)
        Dim sc_gLink1 As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(0)) Then
          sc_gLink1.G6_Cellule_Paint_Candidat(g, gLink.Cdd(0), Color_Link_S)
        End If
        sc_gLink1.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(2)) Then
          sc_gLink1.G6_Cellule_Paint_Candidat(g, gLink.Cdd(2), Color_Link_S)
        End If
        ' Second lien
        G0_Cdd_Bézier(g, gLink.Cel(0), CInt(gLink.Cdd(1)), gLink.Cel(1), CInt(gLink.Cdd(3)), gLink.Type, 0)
        Dim sc_gLink2 As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
        If U(gLink.Cel(0), 3).Contains(gLink.Cdd(1)) Then
          sc_gLink2.G6_Cellule_Paint_Candidat(g, gLink.Cdd(1), Color_Link_S)
        End If
        sc_gLink2.Numéro = gLink.Cel(1)
        If U(gLink.Cel(1), 3).Contains(gLink.Cdd(3)) Then
          sc_gLink2.G6_Cellule_Paint_Candidat(g, gLink.Cdd(3), Color_Link_S)
        End If
      Next gLink

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each XCelExcl As GCel_Excl_Cls In GRslt.CelExcl
        With XCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            If Pbl_Cell_Select = .Cel Then
              Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
              U_Strg_Cdd_Exc(.Cel) = Candidats
            End If
          End If
        End With
      Next XCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.CelExcl.Count & " candidat(s) rouge(s) à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_GCs(g As Graphics)
    If Not (Plcy_Strg = "GCs") Then Exit Sub
    Try
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each gLink As GLink_Cls In GRslt.RoadRight
        Nb += 1
        With gLink
          G0_Cdd_Bézier(g, .Cel(0), CInt(.Cdd(0)), .Cel(1), CInt(.Cdd(2)), .Type, Nb)
          Select Case Nb Mod 2
            Case 0
              Dim sc_Cdd As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
              sc_Cdd.G6_Cellule_Paint_Candidat(g, gLink.Cdd(0), Color.Green)
            Case Else
              Dim sc_cdd As New Cellule_Cls With {.Numéro = gLink.Cel(0)}
              sc_cdd.G6_Cellule_Paint_Candidat(g, gLink.Cdd(0), Color.Blue)
          End Select
        End With
      Next gLink

      ' 3 Affichage des Extrémités des liens et du Dernier candidat de la chaîne
      Dim firstLien As GLink_Cls = GRslt.RoadRight.First()
      Dim lastLien As GLink_Cls = GRslt.RoadRight.Last()
      Dim sc_LCdd As New Cellule_Cls With {.Numéro = lastLien.Cel(1)}
      Dim lastIndex As Integer = GRslt.RoadRight.Count - 1
      Select Case lastIndex Mod 2
        Case 0 : sc_LCdd.G6_Cellule_Paint_Candidat(g, lastLien.Cdd(0), Color.Green)
        Case 1 : sc_LCdd.G6_Cellule_Paint_Candidat(g, lastLien.Cdd(0), Color.Blue)
      End Select
      G0_Cell_Icône(g, firstLien.Cel(0), "Start")
      G0_Cell_Icône(g, lastLien.Cel(1), "End")

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each gCelExcl As GCel_Excl_Cls In GRslt.CelExcl
        With gCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next gCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub
  Public Sub G4_Grid_Stratégie_GCx(g As Graphics)
    If Not (Plcy_Strg = "GCx") Then Exit Sub
    Try
      Dim sc As New Cellule_Cls
      If GRslt.Productivité = False Then
        Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
        Exit Sub
      End If

      ' 1 Affichage des Candidats
      For i As Integer = 0 To 80
        sc.Numéro = i
        sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
        If U(i, 3).Contains(GRslt.Candidat(0)) Then
          G0_Cdd_Figure(g, i, CInt(GRslt.Candidat(0)), "Cercle", Color.White)
        End If
      Next i

      ' 2 Affichage des Courbes de Bézier
      Dim Nb As Integer = 0
      For Each Link As GLink_Cls In GRslt.RoadRight
        Nb += 1
        G0_Cdd_Bézier(g, Link.Cel(0), CInt(Link.Cdd(0)), Link.Cel(1), CInt(Link.Cdd(2)), Link.Type, Nb)
        Dim sc_GCx As New Cellule_Cls With {.Numéro = Link.Cel(0)}
        'Les candidats sont dessinés s'ils existent
        If U(Link.Cel(0), 3).Contains(Link.Cdd(0)) Then
          sc_GCx.G6_Cellule_Paint_Candidat(g, Link.Cdd(0), Color_Link_W)
        End If
        sc_GCx.Numéro = Link.Cel(1)
        If U(Link.Cel(1), 3).Contains(Link.Cdd(2)) Then
          sc_GCx.G6_Cellule_Paint_Candidat(g, Link.Cdd(2), Color_Link_W)
        End If

        ' 3 Affichage des Extrémités des liens  
        Dim PremierLien As GLink_Cls = GRslt.RoadRight.First()
        Dim DernierLien As GLink_Cls = GRslt.RoadRight.Last()
        G0_Cell_Icône(g, PremierLien.Cel(0), "Start")
        G0_Cell_Icône(g, DernierLien.Cel(1), "End")
      Next Link

      ' 4 Affichage des Candidats à exclure 
      Dim Candidats As String = Cnddts_Blancs
      For Each gCelExcl As GCel_Excl_Cls In GRslt.CelExcl
        With gCelExcl
          sc.Numéro = .Cel
          If U(.Cel, 3).Contains(.Cdd) Then
            sc.G6_Cellule_Paint_Candidat(g, .Cdd, Color_Cdd_Exclure)
            ' Coloration du menu contextuel avec les 2 candidats
            Mid$(Candidats, CInt(.Cdd), 1) = .Cdd
            U_Strg_Cdd_Exc(.Cel) = Candidats
          End If
        End With
      Next gCelExcl
      Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & GRslt.Candidat(0) & " rouge à enlever."

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Public Sub G4_Grid_Stratégie_Obj(g As Graphics)
    Dim sc As New Cellule_Cls
    If Not Plcy_Strg = "Obj" Then Exit Sub
    For i As Integer = 0 To 80
      sc.Numéro = i
      sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")
    Next i

    For Each Obj As Objet_Cls In Objet_List
      With Obj

        Select Case .Forme
          Case "Cadre", "Carré", "Disque", "Cercle", "Croix"
            Select Case .Cdd_From
              Case 0    'Cellule
                G0_Cell_Figure(g, .Cel_From, .Forme, Color_BySymbol(.Symbol))
              Case Else 'Candidat
                G0_Cdd_Figure(g, .Cel_From, .Cdd_From, .Forme, Color_BySymbol(.Symbol))
            End Select
          Case "Flèche"
            G0_Cdd_Flèche(g, .Cel_From, .Cdd_From, .Cel_To, .Cdd_To, Color_BySymbol(.Symbol))
          Case Else
        End Select
      End With
    Next Obj
  End Sub

#End Region

#Region "Paint Figures et Formes de base"

  Public Sub G0_Cell_Icône(g As Graphics, Cellule As Integer, Icône As String)
    ' Place une icône de Début dans le premier candidat "libre" de la cellule
    '       une icône de Fin   dans le second candidat "libre" de la cellule

    Dim Cdd_i As Integer() = {5, 1, 3, 7, 9, 2, 4, 6, 8}
    Dim candidatsVides As Integer() = Cdd_i.
       Where(Function(c) Not U(Cellule, 3).Contains(c.ToString())).Take(2).ToArray()

    Select Case Icône
      Case "Start"
        Dim rect As Rectangle = Sqr_Cdd((Cellule * 10) + candidatsVides(0))
        Using img As Image = My.Resources.XC_Start
          g.DrawImage(img, rect)
        End Using
      Case "End"
        Dim rect As Rectangle = Sqr_Cdd((Cellule * 10) + candidatsVides(1))
        Using img As Image = My.Resources.XC_End
          g.DrawImage(img, rect)
        End Using
    End Select
  End Sub
  Public Sub G0_Cell_Figure(g As Graphics, Cellule As Integer, Figure As String, Couleurp As Color)
    Dim Couleur As Color = Color.FromArgb(192, Couleurp) ' 128+64 = 192

    'Dessine des figures à partir de 4 points A0, B0, C0, D0 représentant les 4 points du square A0= Left_Top, C0= Right_Bottom
    '                             déclinés en A1, B1, C1, D1 décalé de 1/4 
    '                             déclinés en A2, B2, C2, D2 décalé de 2/4 
    '                             déclinés en A3, B3, C3, D3 décalé de 3/4 
    Const Inflate As Integer = -5

    Dim Rct_Cercle As New Rectangle(x:=CInt(U_Pt20(Cellule, 0).X), CInt(U_Pt20(Cellule, 0).Y), width:=WH - 3, height:=WH - 3)
    Dim Pt1 As Point, Pt2 As Point, Pt3 As Point, Pt4 As Point
    Pt1 = New Point(CInt(U_Pt20(Cellule, 0).X), CInt(U_Pt20(Cellule, 0).Y))
    Pt2 = New Point(CInt(U_Pt20(Cellule, 1).X), CInt(U_Pt20(Cellule, 1).Y))
    Pt3 = New Point(CInt(U_Pt20(Cellule, 2).X), CInt(U_Pt20(Cellule, 2).Y))
    Pt4 = New Point(CInt(U_Pt20(Cellule, 3).X), CInt(U_Pt20(Cellule, 3).Y))

    Using pen As New Pen(Couleur, 2),
          brsh As New SolidBrush(Couleur),
          pen_Inflate As New Pen(Couleur, Math.Abs(Inflate) * 2)
      Select Case Figure
        Case "Cadre"
          ' Comme le trait du rectangle dépasse de la moitié à droite et à gauche du trait, l'inflate doit être le double
          Rct_Cercle.Inflate(Inflate, Inflate)
          g.DrawRectangle(pen_Inflate, Rct_Cercle)
        Case "Carré"
          g.FillRectangle(brsh, Rct_Cercle)
        Case "Croix"
          g.DrawLine(pen, Pt1, Pt3)
          g.DrawLine(pen, Pt2, Pt4)
        Case "Cercle"
          g.DrawArc(pen, Rct_Cercle, 0.0F, 360.0F)
        Case "Disque"
          g.FillPie(brsh, Rct_Cercle, 0.0F, 360.0F)
        Case "Ellipse"
          '28/01/2025 Cette figure n'est plus utilisée
          Dim Rct_V As New Rectangle(x:=Sqr_Cel(Cellule).X + (1 * WHthird), y:=Sqr_Cel(Cellule).Y, width:=WHthird, height:=WH - 3)
          Dim Rct_H As New Rectangle(x:=Sqr_Cel(Cellule).X + 1, y:=Sqr_Cel(Cellule).Y + (1 * WHthird), width:=WH - 3, height:=WHthird)
          g.DrawEllipse(pen, Rct_V)
          g.DrawEllipse(pen, Rct_H)
          g.DrawArc(pen, Rct_Cercle, 0.0F, 360.0F)
        Case "Double_Carré"
          g.DrawPolygon(pen, {U_Pt20(Cellule, 4), U_Pt20(Cellule, 5), U_Pt20(Cellule, 6), U_Pt20(Cellule, 7)})
          g.DrawPolygon(pen, {U_Pt20(Cellule, 12), U_Pt20(Cellule, 13), U_Pt20(Cellule, 14), U_Pt20(Cellule, 15)})
        Case Else
          Jrn_Add(, {Proc_Name_Get() & " Figure Inconnue : " & Figure}, "Erreur")
      End Select
    End Using
  End Sub

  Public Sub G0_Cdd_Figure(g As Graphics, Cellule As Integer, Candidat As Integer, Figure As String, Couleurp As Color)
    Dim Couleur As Color = Color.FromArgb(192, Couleurp) ' 128+64 = 192
    'Dessine des figures à partir de 4 points A0, B0, C0, D0 représentant les 4 points du square A0= Left_Top, C0= Right_Bottom
    '                             déclinés en A1, B1, C1, D1 décalé de 1/4 
    '                             déclinés en A2, B2, C2, D2 décalé de 2/4 
    '                             déclinés en A3, B3, C3, D3 décalé de 3/4 

    If Cellule < 0 Or Cellule > 80 Then Exit Sub
    If Candidat < 1 Or Candidat > 9 Then Exit Sub

    Dim Cdd_n As Integer = (Cellule * 10) + Candidat
    Dim Sqr_Cdd_n As Rectangle = Sqr_Cdd(Cdd_n)
    Sqr_Cdd_n.Inflate(-1, -1)    'Diminution du cercle du candidat  
    ' Définir les coins du rectangle
    Dim Pt1 As New Point(Sqr_Cdd_n.X, Sqr_Cdd_n.Y)                                      ' Coin supérieur gauche
    Dim Pt2 As New Point(Sqr_Cdd_n.X + Sqr_Cdd_n.Width, Sqr_Cdd_n.Y)                    ' Coin supérieur droit
    Dim Pt3 As New Point(Sqr_Cdd_n.X + Sqr_Cdd_n.Width, Sqr_Cdd_n.Y + Sqr_Cdd_n.Height) ' Coin inférieur droit
    Dim Pt4 As New Point(Sqr_Cdd_n.X, Sqr_Cdd_n.Y + Sqr_Cdd_n.Height)                   ' Coin inférieur gauche

    Using pen As New Pen(Couleur, 2),
          brsh As New SolidBrush(Couleur)
      Select Case Figure
        Case "Cadre"
          g.DrawRectangle(pen, Sqr_Cdd_n)
        Case "Carré"
          g.FillRectangle(brsh, Sqr_Cdd_n)
        Case "Croix"
          g.DrawLine(pen, Pt1, Pt3)
          g.DrawLine(pen, Pt2, Pt4)
        Case "Cercle"
          g.DrawArc(pen, Sqr_Cdd_n, 0.0F, 360.0F)
        Case "Disque"
          g.FillPie(brsh, Sqr_Cdd_n, 0.0F, 360.0F)
        Case Else
          Jrn_Add(, {Proc_Name_Get() & " Figure Inconnue : " & Figure}, "Erreur")
      End Select
    End Using
  End Sub

  Function DeCasteljau(t As Single, p0 As PointF, p1 As PointF, p2 As PointF, p3 As PointF) As PointF
    ' Algorithme de De Casteljau
    Dim x As Double = (1 - t) ^ 3 * p0.X +
          3 * (1 - t) ^ 2 * t * p1.X +
          3 * (1 - t) * t ^ 2 * p2.X +
          t ^ 3 * p3.X
    Dim y As Double = (1 - t) ^ 3 * p0.Y +
          3 * (1 - t) ^ 2 * t * p1.Y +
          3 * (1 - t) * t ^ 2 * p2.Y +
          t ^ 3 * p3.Y
    Return New PointF(CSng(x), CSng(y))
  End Function

  Public Sub G0_Cdd_Bézier(g As Graphics, From_Cellule As Integer, From_Candidat As Integer, To_Cellule As Integer, To_Candidat As Integer, Link_Type As String, Link_Numéro As Integer)
    ' 1 Calcul des Centres et des Points de contrôle pour une courbe de Bézier
    Dim From_Centre As PointF = Get_CentreF(From_Cellule, From_Candidat)
    Dim To_Centre As PointF = Get_CentreF(To_Cellule, To_Candidat)
    Dim Pts As Points_Struct = Get_AdjustedPoints(From_Centre, To_Centre)
    Dim Décalage As Integer = 5
    Dim From_Ctrl As New PointF(From_Centre.X + Décalage, From_Centre.Y + Décalage)
    Dim To_Ctrl As New PointF(To_Centre.X + Décalage, To_Centre.Y + Décalage)

    ' 2 Dessine la courbe de Bézier 
    '   la couleur dépend du type de lien, ainsi que le style
    Dim Couleur As Color
    Dim Style As DashStyle

    Select Case Link_Type
      Case "S"
        Couleur = Color_Link_S
        Style = DashStyle.Solid
      Case "W"
        Couleur = Color_Link_W
        Style = DashStyle.Dash
    End Select

    Using Crayon As New Pen(Couleur, 2)
      Dim Cap As New AdjustableArrowCap(4, 6)
      Crayon.CustomEndCap = Cap
      Crayon.DashStyle = Style
      g.SmoothingMode = SmoothingMode.AntiAlias
      g.DrawBezier(Crayon, Pts.Pt_From, From_Ctrl, To_Ctrl, Pts.Pt_To)
    End Using

    ' 3 Affichage de la lettre du lien
    Dim PointMid As PointF = DeCasteljau(0.5, Pts.Pt_From, From_Ctrl, To_Ctrl, Pts.Pt_To)
    Using font As New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Italic),
          brsh As New SolidBrush(Color.Black)
      g.DrawString(ChrW(Link_Numéro + Lettre_Flèche_ChrW), font, brsh, PointMid.X, PointMid.Y)
    End Using
  End Sub

  Public Sub G0_Cell_Diagonale(g As Graphics, From_Cellule As Integer, To_Cellule As Integer)
    'Dessine une Flèche
    Dim Pt_From_Cellule, Pt_To_Cellule As Point
    Dim sc_From As New Cellule_Cls With {.Numéro = From_Cellule}
    Pt_From_Cellule = New Point(sc_From.Position_Center.X, sc_From.Position_Center.Y)
    Dim sc_To As New Cellule_Cls With {.Numéro = To_Cellule}
    Pt_To_Cellule = New Point(sc_To.Position_Center.X, sc_To.Position_Center.Y)
    'Cells_Bresenham_Get(U_Strg, From_Cellule, To_Cellule)
    Using Pen As New Pen(Color_Stratégique, WH \ 2)
      g.DrawLine(Pen, Pt_From_Cellule, Pt_To_Cellule)
    End Using
  End Sub
  Public Sub Cells_Bresenham_Get(ByRef U_Strg() As Boolean, ByVal From_Cellule As Integer, ByVal To_Cellule As Integer)
    ' Documente U_Strg des cellules traversées par la ligne Pt_From Pt_To
    ' Algorithme de la droite de Bresenham IBM 1962
    ' La position Top-Left ou center des Pt_From_To donne des résultats différents, 
    '    ainsi que l'épaisseur du trait  
    Dim sc_From As New Cellule_Cls With {.Numéro = From_Cellule}
    Dim Pt_From As New Point(sc_From.Position_Center.X, sc_From.Position_Center.Y)
    Dim sc_To As New Cellule_Cls With {.Numéro = To_Cellule}
    Dim Pt_To As New Point(sc_To.Position_Center.X, sc_To.Position_Center.Y)

    Dim dx As Integer = Math.Abs(Pt_To.X - Pt_From.X)
    Dim dy As Integer = Math.Abs(Pt_To.Y - Pt_From.Y)
    Dim sx As Integer = If(Pt_From.X < Pt_To.X, 1, -1)
    Dim sy As Integer = If(Pt_From.Y < Pt_To.Y, 1, -1)
    Dim err As Integer = dx - dy

    Dim x As Integer = Pt_From.X
    Dim y As Integer = Pt_From.Y

    While True
      Dim Cellule As Integer = Array.FindIndex(Sqr_Cel, Function(cel) cel.Contains(New Point(x, y)))
      If Cellule <> -1 Then U_Strg(Cellule) = True
      If x = Pt_To.X And y = Pt_To.Y Then Exit While
      Dim e2 As Integer = 2 * err
      If e2 > -dy Then
        err -= dy
        x += sx
      End If
      If e2 < dx Then
        err += dx
        y += sy
      End If
    End While
  End Sub
  Public Sub G4_MdC_Row_Col_Box(Code_LCR As String, LCR As Integer)
    If LCR < 0 Or LCR > 8 Then Exit Sub
    Dim Grp(0 To 8) As Integer
    Dim Cellule As Integer
    Select Case Code_LCR
      Case "Row"
        Grp = U_9CelRow(LCR)
        For i As Integer = 0 To 8
          Cellule = Grp(i)
          Select Case i
            Case 0 : G4_Cell_MdC(Cellule, "PD")
            Case 8 : G4_Cell_MdC(Cellule, "PG")
            Case Else : G4_Cell_MdC(Cellule, "RH")
          End Select
        Next i
      Case "Col"
        Grp = U_9CelCol(LCR)
        For i As Integer = 0 To 8
          Cellule = Grp(i)
          Select Case i
            Case 0 : G4_Cell_MdC(Cellule, "PB")
            Case 8 : G4_Cell_MdC(Cellule, "PH")
            Case Else : G4_Cell_MdC(Cellule, "RV")
          End Select
        Next i
      Case "Box"
        Grp = U_9CelReg(LCR)
        'Les cellules de la région sont numérotées
        ' 0 1 2
        ' 3 4 5
        ' 6 7 8
        Dim b As Integer
        b = 0 : G4_Cell_MdC(Grp(b), "CHG")
        b = 1 : G4_Cell_MdC(Grp(b), "RH")
        b = 2 : G4_Cell_MdC(Grp(b), "CHD")
        b = 3 : G4_Cell_MdC(Grp(b), "RV")
        b = 4 : G4_Cell_MdC(Grp(b), "SQ")
        b = 5 : G4_Cell_MdC(Grp(b), "RV")
        b = 6 : G4_Cell_MdC(Grp(b), "CBG")
        b = 7 : G4_Cell_MdC(Grp(b), "RH")
        b = 8 : G4_Cell_MdC(Grp(b), "CBD")
      Case Else
        Exit Select
    End Select
  End Sub

  Public Sub G4_MdC_Trait_ou_Rectangle(Cellule_A As Integer, Cellule_B As Integer)
    'La procédure dessine soit un trait horizontal, soit un trait vertical, soit un rectangle
    'Toujours de 1 vers 2 en croissant, id de gauche à droite, de haut en bas, de la cellule 0 à 81
    Dim Cel, Cel_1, Cel_2 As Integer
    If Cellule_A < Cellule_B Then
      Cel_1 = Cellule_A : Cel_2 = Cellule_B
    Else
      Cel_1 = Cellule_B : Cel_2 = Cellule_A
    End If
    Dim w, h As Integer ' w et h sont toujours positifs
    w = U_Col(Cel_2) - U_Col(Cel_1)
    h = U_Row(Cel_2) - U_Row(Cel_1)

    If Cellule_A = Cellule_B Then Exit Sub

    If h = 0 Then    'Trait Horizontal 
      G4_Cell_MdC(Cel_1, "PD")
      For i As Integer = U_Col(Cel_1) + 1 To U_Col(Cel_2) - 1 Step 1
        Cel = Wh_Cellule_ColRow(i, U_Row(Cel_1))
        G4_Cell_MdC(Cel, "RH")
      Next i
      G4_Cell_MdC(Cel_2, "PG")
      Exit Sub
    End If '

    If w = 0 Then    'Trait vertical 
      G4_Cell_MdC(Cel_1, "PB")
      For i As Integer = U_Row(Cel_1) + 1 To U_Row(Cel_2) - 1 Step 1
        Cel = Wh_Cellule_ColRow(U_Col(Cel_1), i)
        G4_Cell_MdC(Cel, "RV")
      Next i
      G4_Cell_MdC(Cel_2, "PH")
      Exit Sub
    End If '

    'Forme rectangulaire
    G4_Cell_MdC(Cel_1, "CHG")
    For i As Integer = U_Col(Cel_1) + 1 To U_Col(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(i, U_Row(Cel_1))
      G4_Cell_MdC(Cel, "RH")
    Next i
    G4_Cell_MdC(Cel_1 + w, "CHD")
    For i As Integer = U_Row(Cel_1) + 1 To U_Row(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(U_Col(Cel_1), i)
      G4_Cell_MdC(Cel, "RV")
    Next i
    G4_Cell_MdC(Cel_2, "CBD")
    For i As Integer = U_Col(Cel_1) + 1 To U_Col(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(i, U_Row(Cel_2))
      G4_Cell_MdC(Cel, "RH")
    Next i
    G4_Cell_MdC(Cel_2 - w, "CBG")
    For i As Integer = U_Row(Cel_1) + 1 To U_Row(Cel_2) - 1 Step 1
      Cel = Wh_Cellule_ColRow(U_Col(Cel_2), i)
      G4_Cell_MdC(Cel, "RV")
    Next i
  End Sub

  Public Sub G4_Cell_MdC(Cellule As Integer, Modèle As String)
    'Exemple pour dessiner des Modèles
    'U_MdC_Clear()           'Initialisation pour chaque cellule d'une suite de modèle
    'G4_Cell_MdC(40, "CHG") 'Ajout d'un modèle dans une cellule 
    'G4_Cell_MdC(41, "CHD")
    'G4_Cell_MdC(50, "CBD")
    'G4_Cell_MdC(49, "CBG")
    'G4_MdC_Row_Col_Box("Box", 7)
    'G4_MdC_Paint()         'Les figures sont dessinées et les candidats affichés

    For j As Integer = 0 To U_MdC(Cellule).MdE.Count - 1
      If U_MdC(Cellule).MdE(j) = "" Then
        U_MdC(Cellule).MdE(j) = Modèle
        U_MdC(Cellule).MdE_Exist = True
        Exit For
      End If
    Next j
  End Sub
  Public Sub G4_MdC_Paint(g As Graphics)
    ' Construction et peinture des Modèles Composites
    ' Un modèle est dit composite quand il s'agit de SUPERPOSER plusieurs modèles élémentaires
    '                             comme un croisement de colonne et de ligne
    '           afin qu'il n'y ait pas de superposition de couleurs
    ' Ces modèles sont construits à/p des Modèles Elémentaires
    '   un modèle composite est constitué de PLUSIEURS modèles ELEMENTAIRES
    '                       avec une Région et Union  
    ' Les candidats sont également affichés après la couche stratégique.

    Dim MdC_Exist As Boolean
    Dim MdC_Cellule As Integer
    Dim MdC_N As Integer
    Dim MdC_Modèle As String
    For i As Integer = 0 To 80
      ' Phase 1 Construit une Région d'UNION de tous les modèles à dessiner
      MdC_Exist = False
      MdC_Cellule = U_MdC(i).Cellule
      MdC_N = 0
      Dim MdC_Région As New Region
      For j As Integer = 0 To U_MdC(i).MdE.Count - 1
        If U_MdC(i).MdE(j) <> "" Then
          MdC_Exist = True
          MdC_Modèle = U_MdC(i).MdE(j)
          Dim Pth_Modèle(MdC_N) As GraphicsPath
          Pth_Modèle(MdC_N) = G0_MdE_Build(MdC_Cellule, MdC_Modèle)
          If MdC_N = 0 Then MdC_Région = New Region(Pth_Modèle(MdC_N))
          If MdC_N > 0 Then MdC_Région.Union(Pth_Modèle(MdC_N))
          Pth_Modèle(MdC_N).Dispose()
          MdC_N += 1
        End If
      Next j

      ' Phase 2 Dessine la Région et Peint les valeurs et les candidats concernés
      If MdC_Exist Then
        ' Il est à noter que l'ordre des couches est respecté,
        ' Il n'est donc pas nécessaire de conserver le composant alpha à 128 pour la couche stratégique
        ' La couche stratégique a une étendue "grille". 
        Using brsh As New SolidBrush(Color_Stratégique)
          g.FillRegion(brsh, MdC_Région)                    ' G4
        End Using
        Dim sc As New Cellule_Cls With {.Numéro = i}
        If sc.Valeur <> 0 Then sc.G5_Cellule_Paint_Valeur(g)                            ' G5
        If sc.Valeur = 0 Then sc.G6_Cellule_Paint_Candidats(g, "LesCandidatsEligibles")   ' G6
      End If
      MdC_Région.Dispose()
    Next i

  End Sub

  '-------------------------------------------------------------------------------
  'Idée     : Créer des formes primaires comme un rectangle ou un pouce ou un coude
  '           Décliner ces formes en 2 formes élémentaires Rectangle Horizontal-Vertical
  '                               en 4 formes élémentaires Pouce Haut-Bas et Droit-Gauche
  '                               en 4 formes élémentaires Coude Haut-Droit, Haut-Gauche, Bas-Droit, Bas-Gauche
  '           Combiner des formes élémentaires pour créer une intersection de Rectangles, des Lettres T, 
  '                               des associations rectangles et coude,
  '-------------------------------------------------------------------------------   
  'Vocabulaire:
  ' Modèle Primaire        R(V)         C (CHG)        P(B)
  '        Elémentaire     RV, RH       C H/B-D/G      P B/G/H/D/G
  '        Composite       La croix = RV Union RH
  '                        
  Public Sub U_MdC_Clear()
    'Initialisation de U_Mdc : 1 ligne par Cellule
    'Redim doit être fait pour chaque ligne
    '42 L5-C7 :   RV   PD                                         
    '43 L5-C8 :  CHG  CHD  CBD  CBG                               
    '44 L5-C9 :   RV   PG                                         
    '51 L6-C7 :  CBG
    For i As Integer = 0 To 80
      ReDim U_MdC(i).MdE(9)
      U_MdC(i).Cellule = i
      U_MdC(i).MdE_Exist = False
      For j As Integer = 0 To 9
        U_MdC(i).MdE(j) = ""
      Next j
    Next i
  End Sub
  Public Sub U_MdC_Display()
    Dim M1 As String
    Dim M2 As String
    Jrn_Add(, {"Liste de U_MdC"})
    For i As Integer = 0 To 80
      Dim MdC_Exist As Boolean = False
      M2 = ""
      For j As Integer = 0 To U_MdC(i).MdE.Count - 1
        If U_MdC(i).MdE(j) <> "" Then MdC_Exist = True
        M2 &= U_MdC(i).MdE(j).PadLeft(4) & " "
      Next j
      M1 = CStr(i).PadLeft(3) & " " & U_Coord(U_MdC(i).Cellule)
      If MdC_Exist = True Then Jrn_Add(, {M1 & " : " & M2})
    Next i
  End Sub

  Public Function G0_MdP_Build(Cellule As Integer, Modèle As String) As GraphicsPath

    'Nom des Formes, les formes diffèrent selon Format_Ang "Angles droits" ou "Coins arrondis ou biseautés"
    'C  Couronne       B/H Bas-Haut             D/G Droite-Gauche
    'R  Rectangle      H/V Horizontal-Vertical    
    'CR Croix
    'P  Pouce          BHGD
    'PT Point          (Non utilisé!)
    'T  Lettre T       B/H/D/G Bas-Haut-Droit-Gauche

    'Construction des Modèles Primaires
    '----------------------------------
    'Exemple: Modèle primaire du coin haut gauche (droit, arrondi ou biseauté)
    'à/p de ce modèle les modèles élémentaires CHG, CHD, CBD et CBG seront construits par rotation
    'G0_MdP_Build n'est utilisée que dans G0_MdE_Build

    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    Dim Cell_A As New Point(sc.Position.X, sc.Position.Y)
    Dim Cell_B As New Point(sc.Position.X + WH, sc.Position.Y)
    Dim Cell_C As New Point(sc.Position.X + WH, sc.Position.Y + WH)
    Dim Cell_D As New Point(sc.Position.X, sc.Position.Y + WH)

    Dim Pth_Modèle As New GraphicsPath
    Select Case Modèle
      Case "R" 'Rectangle Vertical
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y)
        Dim B As New Point(Cell_B.X - WHquart, Cell_B.Y)
        Dim C As New Point(Cell_C.X - WHquart, Cell_C.Y)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y)
        With Pth_Modèle
          .StartFigure()
          .AddPolygon({A, B, C, D, A})
          .CloseFigure()
        End With
      Case "S" 'Carré Central (Utilisé dans le centre des régions)
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y + WHquart)
        Dim B As New Point(Cell_B.X - WHquart, Cell_B.Y + WHquart)
        Dim C As New Point(Cell_C.X - WHquart, Cell_C.Y - WHquart)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y - WHquart)
        With Pth_Modèle
          .StartFigure()
          .AddPolygon({A, B, C, D, A})
          .CloseFigure()
        End With

      Case "P" 'Pouce Bas
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y + WHquart)
        Dim B As New Point(Cell_B.X - WHquart, Cell_B.Y + WHquart)
        Dim C As New Point(Cell_A.X + WHquart + WHhalf, Cell_C.Y)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y)
        Dim AD As New Point(Cell_A.X + WHquart, Cell_A.Y + WHhalf)
        Dim BC As New Point(Cell_A.X + WHquart + WHhalf, Cell_A.Y + WHhalf)
        Dim Rect_A_m As New Rectangle(A.X, A.Y, WHhalf, WHhalf)
        Dim A1 As New Point(Cell_A.X + WHquart, Cell_A.Y + WHhalf)
        Dim A2 As New Point(Cell_A.X + WHhalf, Cell_A.Y + WHquart)
        Dim A3 As New Point(Cell_A.X + WHhalf + WHquart, Cell_A.Y + WHhalf)
        With Pth_Modèle
          .StartFigure()
          Select Case Plcy_Format_DAB
            Case 0       ' Droits 
              .AddPolygon({A, B, C, D, A})
            Case 1 ', 3, 5 ' Arrondis
              .AddArc(Rect_A_m, 180, 180)
              .AddLine(BC, C)
              .AddLine(C, D)
              .AddLine(D, AD)
          End Select
          .CloseFigure()
        End With
      Case "C" 'Coin Haut Gauche
        Dim A As New Point(Cell_A.X + WHquart, Cell_A.Y + WHquart)
        Dim B As New Point(Cell_B.X, Cell_B.Y + WHquart)
        Dim BC As New Point(Cell_C.X, Cell_C.Y - WHquart)
        Dim C As New Point(Cell_C.X - WHquart, Cell_C.Y - WHquart)
        Dim CD As New Point(Cell_C.X - WHquart, Cell_C.Y)
        Dim D As New Point(Cell_D.X + WHquart, Cell_D.Y)
        Dim Rect_A_g As New Rectangle(A.X, A.Y, WH + WHhalf, WH + WHhalf)
        Dim Rect_C_p As New Rectangle(C.X, C.Y, WHhalf, WHhalf)
        Dim A1 As New Point(Cell_A.X + WHquart, Cell_A.Y + WHhalf + WHquart)
        Dim A2 As New Point(Cell_A.X + WHhalf + WHquart, Cell_A.Y + WHquart)

        With Pth_Modèle
          .StartFigure()
          Select Case Plcy_Format_DAB
            Case 0         ' Droits 
              .AddPolygon({A, B, BC, C, CD, D, A})
            Case 1
              'Un arc est compris dans un rectangle
              'Un rectangle est défini par un Left-Top (x-y) et Width-Height
              'L'angle Start à 0 sur l'axe x et tourne dans le sens des aiguilles d'une montre
              '                90, 180 , 270
              'la longueur de l'angle est de 90° dans un sens et -90° dans l'autre                
              .AddArc(Rect_A_g, 180, 90)
              .AddLine(B, BC)
              .AddArc(Rect_C_p, 270, -90)
              .AddLine(CD, D)
          End Select
          .CloseFigure()
        End With
      Case Else
    End Select
    Return Pth_Modèle
  End Function
  Public Function G0_MdE_Build(Cellule As Integer, Modèle As String) As GraphicsPath
    'Construction des Modèles Elémentaires
    ' Ces modèles sont construits à/p des Modèles Primaires à une lettre avec RotateAt
    ' Les modèles Elémentaires ont 2 ou 3 lettres.
    Dim sc As New Cellule_Cls With {.Numéro = Cellule}
    Dim Pth_Modèle As GraphicsPath = Nothing

    Select Case Modèle
      Case "RV" 'Rectangle Vertical
        Pth_Modèle = G0_MdP_Build(Cellule, "R")

      Case "RH" 'Rectangle Horizontal
        Pth_Modèle = G0_MdP_Build(Cellule, "R")
        Using mat As New Matrix
          mat.RotateAt(90, New PointF(sc.Position_Center.X, sc.Position_Center.Y))
          Pth_Modèle.Transform(mat)
        End Using

      Case "SQ" 'Carré central
        Pth_Modèle = G0_MdP_Build(Cellule, "S")

      Case "CHG", "CHD", "CBD", "CBG" 'Coin Haut/Bas Gauche/Droit
        Dim Angle As Integer = 0
        If Modèle = "CHG" Then Angle = 0
        If Modèle = "CHD" Then Angle = 90
        If Modèle = "CBD" Then Angle = 180
        If Modèle = "CBG" Then Angle = 270
        Pth_Modèle = G0_MdP_Build(Cellule, "C")
        Using mat As New Matrix
          mat.RotateAt(Angle, New PointF(sc.Position_Center.X, sc.Position_Center.Y))
          Pth_Modèle.Transform(mat)
        End Using

      Case "PB", "PG", "PH", "PD" 'Pouce Haut/Bas Gauche/Droit
        Dim Angle As Integer = 0
        If Modèle = "PB" Then Angle = 0
        If Modèle = "PG" Then Angle = 90
        If Modèle = "PH" Then Angle = 180
        If Modèle = "PD" Then Angle = 270
        Pth_Modèle = G0_MdP_Build(Cellule, "P")
        Using mat As New Matrix
          mat.RotateAt(Angle, New PointF(sc.Position_Center.X, sc.Position_Center.Y))
          Pth_Modèle.Transform(mat)
        End Using
      Case Else
    End Select
    Return Pth_Modèle
  End Function
#End Region
End Module