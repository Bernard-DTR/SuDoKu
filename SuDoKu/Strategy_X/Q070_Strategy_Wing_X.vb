Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
' Dimanche 17/08/2025
' Stratégie Wing_X
'  
' Préfixe WgX, 
' Les prochaines stratégies seront XY, XYZ et W, soit WgY, WgZ et WgW


'...9...5..5.26.91........7..9...16......42..1...3.....5.817..6.6.4..35....3......
'.....27...3.........8..7.9.........58......6..6.1.59.4........8.45.19.3...34..51.
'-------------------------------------------------------------------------------------------
Friend Module Q070_Strategy_Wing_X

  Public Sub Strategy_WgX(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "WgX" 
    XRslt_Init()

    ' 2 Inventaire des candidats par ordre décroissant
    XCdds_List.Clear()
    XCdds_List_Get(U_temp)
    If Xap Then XCdds_List_Display()

    For Each XCdd As XCdd_Cls In XCdds_List
      If XRslt.Productivité Then Exit For

      ' 10 Initialisation de XRslt avec Plcy_Strg = "WgX"
      XRslt_Init()
      XRslt.Candidat(0) = XCdd.Cdd

      If Xap Then Jrn_Add("SDK_Space")
      If Xap Then Jrn_Add(, {"Candidat traité : " & XCdd.Cdd & " ,Qté : " & XCdd.Nb})

      ' 20 Création des Liens Forts, il faut seulement 2 candidats par unité pour faire un lien fort
      XLinks_List.Clear()
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Row", U_9CelRow) ' Analyse des rangées
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Col", U_9CelCol) ' Analyse des colonnes
      'Links_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Reg", U_9CelReg) ' l'Analyse des régions n'est pas concernée par cette stratégie
      If Xap Then XLinks_List_Display("1")

      ' 30 Analyse des Liens Forts
      If XLinks_List.Count < 2 Then Continue For
      For i As Integer = 0 To XLinks_List.Count - 2
        Dim LinkA As XLink_Cls = XLinks_List(i)
        XRslt.RoadRight.Clear()

        For j As Integer = i + 1 To XLinks_List.Count - 1
          Dim LinkB As XLink_Cls = XLinks_List(j)
          XRslt.RoadRight.Clear()

          Select Case LinkA.Unité.Substring(0, 3)
            Case "Row"
              If Is_SameCol(LinkA.Cel(0), LinkB.Cel(0)) AndAlso Is_SameCol(LinkA.Cel(1), LinkB.Cel(1)) Then
                XRslt.RoadRight.Add(LinkA)
                XRslt.RoadRight.Add(LinkB)
                If Candidats_Exclure_WgX(U_temp, XRslt.Candidat(0), "Row", LinkA, LinkB) > 0 Then
                  XRslt.Productivité = True
                  Exit For
                End If
              End If
            Case "Col"
              If Is_SameRow(LinkA.Cel(0), LinkB.Cel(0)) AndAlso Is_SameRow(LinkA.Cel(1), LinkB.Cel(1)) Then
                XRslt.RoadRight.Add(LinkA)
                XRslt.RoadRight.Add(LinkB)
                If Candidats_Exclure_WgX(U_temp, XRslt.Candidat(0), "Col", LinkA, LinkB) > 0 Then
                  XRslt.Productivité = True
                  Exit For
                End If
              End If
          End Select
        Next j
        If XRslt.Productivité Then Exit For
      Next i
      If XRslt.Productivité Then Exit For
    Next XCdd

    Stratégies_G_End()
  End Sub

  Public Function Candidats_Exclure_WgX(U_temp(,) As String, Candidat As String, Code_LCR As String, LinkA As XLink_Cls, LinkB As XLink_Cls) As Integer
    Dim Grp() As Integer
    Dim LCR As Integer

    XRslt.CelExcl.Clear()

    Select Case Code_LCR
      Case "Row"
        LCR = U_Col(LinkA.Cel(0))
        Grp = U_9CelCol(LCR)
        For i As Integer = 0 To 8
          If Grp(i) = LinkA.Cel(0) Or Grp(i) = LinkB.Cel(0) Then Continue For
          If U_temp(Grp(i), 3).Contains(Candidat) Then
            XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = Grp(i), .Cdd = Candidat, .Exc = {LinkA.Cel(0), LinkA.Cel(1)}})
          End If
        Next i
        LCR = U_Col(LinkA.Cel(1))
        Grp = U_9CelCol(LCR)
        For i As Integer = 0 To 8
          If Grp(i) = LinkA.Cel(1) Or Grp(i) = LinkB.Cel(1) Then Continue For
          If U_temp(Grp(i), 3).Contains(Candidat) Then
            XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = Grp(i), .Cdd = Candidat, .Exc = {LinkB.Cel(0), LinkB.Cel(1)}})
          End If
        Next i

      Case "Col"
        LCR = U_Row(LinkA.Cel(0))
        Grp = U_9CelRow(LCR)
        For i As Integer = 0 To 8
          If Grp(i) = LinkA.Cel(0) Or Grp(i) = LinkB.Cel(0) Then Continue For
          If U_temp(Grp(i), 3).Contains(Candidat) Then
            XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = Grp(i), .Cdd = Candidat, .Exc = {LinkA.Cel(0), LinkA.Cel(1)}})
          End If
        Next i
        LCR = U_Row(LinkA.Cel(1))
        Grp = U_9CelRow(LCR)
        For i As Integer = 0 To 8
          If Grp(i) = LinkA.Cel(1) Or Grp(i) = LinkB.Cel(1) Then Continue For
          If U_temp(Grp(i), 3).Contains(Candidat) Then
            XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = Grp(i), .Cdd = Candidat, .Exc = {LinkB.Cel(0), LinkB.Cel(1)}})
          End If
        Next i
    End Select
    Return XRslt.CelExcl.Count
  End Function

End Module
