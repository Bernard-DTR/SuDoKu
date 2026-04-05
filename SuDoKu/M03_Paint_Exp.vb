Module M03_Paint_Exp
  Public Sub G4_Grid_Stratégie_WgX_WgY_WgZ_WgW(g As Graphics)
    If Not {"WgX", "WgY", "WgZ", "WgW"}.Contains(Plcy_Strg) Then Exit Sub
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
            DrawBezierWithCandidates(g,
                                     Link.Cel(0), CInt(Link.Cdd(0)),
                                     Link.Cel(1), CInt(Link.Cdd(2)),
                                     Link.Type, Nb)
          Case "WgY"
            DrawBezierWithCandidates(g,
                                     Link.Cel(0), CInt(Link.Cdd(2)),
                                     Link.Cel(1), CInt(Link.Cdd(2)),
                                     Link.Type, Nb)
            G0_Cdd_Figure(g, Link.Cel(1), CInt(Link.Cdd(3)), "Disque", Color_Link_S)
          Case "WgZ"
            DrawBezierWithCandidates(g,
                                     Link.Cel(0), CInt(Link.Cdd(3)),
                                     Link.Cel(1), CInt(Link.Cdd(3)),
                                     Link.Type, Nb)
            DrawBezierWithCandidates(g,
                                     Link.Cel(0), CInt(Link.Cdd(4)),
                                     Link.Cel(1), CInt(Link.Cdd(4)),
                                     Link.Type, Nb)
            G0_Cdd_Figure(g, Link.Cel(0), CInt(Link.Cdd(5)), "Disque", Color_Link_S)
          Case "WgW"
            DrawBezierWithCandidates(g,
                                     Link.Cel(0), CInt(Link.Cdd(0)),
                                     Link.Cel(1), CInt(Link.Cdd(2)),
                                     Link.Type, Nb)
            G0_Cdd_Figure(g, XRslt.Cellule(0), CInt(XRslt.Candidat(1)), "Disque", Color_Link_S)
            G0_Cdd_Figure(g, XRslt.Cellule(1), CInt(XRslt.Candidat(1)), "Disque", Color_Link_S)
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
  Private Sub PaintIfCandidateExists(g As Graphics, cell As Integer, cdd As Integer, color As Color)
    If U(cell, 3).Contains(CStr(cdd)) Then
      Dim sc As New Cellule_Cls With {.Numéro = cell}
      sc.G6_Cellule_Paint_Candidat(g, CStr(cdd), color)
    End If
  End Sub
  Private Sub DrawBezierWithCandidates(g As Graphics,
                                       fromCell As Integer, fromCdd As Integer,
                                       toCell As Integer, toCdd As Integer,
                                       linkType As String, index As Integer)
    G0_Cdd_Bézier(g, fromCell, fromCdd, toCell, toCdd, linkType, index)
    PaintIfCandidateExists(g, fromCell, fromCdd, Color_Link_W)
    PaintIfCandidateExists(g, toCell, toCdd, Color_Link_W)
  End Sub
  Private Sub DrawStartIcons_Old(g As Graphics)
    Select Case Plcy_Strg
      Case "WgX"
        For i As Integer = 0 To 1
          G0_Cell_Icône(g, XRslt.RoadRight(i).Cel(0), "Start")
        Next
      Case "WgY", "WgZ"
        G0_Cell_Icône(g, XRslt.RoadRight(0).Cel(0), "Start")
      Case "WgW"
        G0_Cell_Icône(g, XRslt.Cellule(0), "Start")
        G0_Cell_Icône(g, XRslt.Cellule(1), "Start")
    End Select
  End Sub
End Module
