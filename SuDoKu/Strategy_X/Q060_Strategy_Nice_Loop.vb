Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------------------
'Vendredi 25/07/2025
' Stratégie Nice_Loop
'  
' Préfixe XNl, Nl
'Strategy_XNl 
'2.......4.1..6..8...83.56....97324..3...5...6..16.93....75.32...3..1..7.5.......3
'2137.8..44.9521.83..843..2..24..7..68.56.2........3..2...18.24..823.4.15941275..8
'-------------------------------------------------------------------------------------------

Friend Module Q060_Strategy_Nice_Loop

  Public Sub Strategy_XNl(U_temp(,) As String)
    If Xap Then Jrn_Add(, {Proc_Name_Get()})

    ' 1 Initialisation de XRslt avec Plcy_Strg = "XNl"
    XRslt_Init()

    ' 2 Inventaire des candidats par ordre décroissant
    XCdds_List.Clear()
    XCdds_List_Get(U_temp)
    If Xap Then XCdds_List_Display()


    For Each XCdd As XCdd_Cls In XCdds_List
      If XRslt.Productivité Then Exit For

      ' 10 Initialisation de XRslt avec Plcy_Strg = "XNl"
      XRslt_Init()
      XRslt.Candidat(0) = XCdd.Cdd

      '    Initialisation de GCels As New List(Of GCel_Cls)
      '    pour une phase Candidats_Exclure_Cc plus rapide
      GCels.Clear()
      For i As Integer = 0 To 80
        If U_temp(i, 3).Contains(XRslt.Candidat(0)) Then
          GCels.Add(New GCel_Cls With {.Cel = i, .Cdd = New String() {"0", "0"}})
        End If
      Next i

      If Xap Then Jrn_Add("SDK_Space")
      If Xap Then Jrn_Add(, {"Candidat traité : " & XCdd.Cdd & " ,Qté : " & XCdd.Nb})

      ' 20 Création des Liens Forts, il faut seulement 2 candidats par unité pour faire un lien fort
      XLinks_List.Clear()
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Row", U_9CelRow) ' Analyse des rangées
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Col", U_9CelCol) ' Analyse des colonnes
      XLinks_List_Generate_XCs_XCx_XNl(U_temp, XCdd.Cdd, "Reg", U_9CelReg) ' Analyse des régions
      If Xap Then XLinks_List_Display("1")

      ' 30 Construction des chemins
      XAllRoads_List.Clear()
      XRoads_Inventory_XCs_XCx_XNl(XLinks_List, New List(Of XLink_Cls)(), XAllRoads_List)
      'If Xap Then Jrn_Add(, {CStr(XAllRoads_List.Count) & " XAllRoads_List.Count"})
      XRslt.XRoads_Nombre = XAllRoads_List.Count

      ' 40 Vérification des chemins
      If XRoads_Vérification_XNl(XAllRoads_List) Then Exit For

    Next XCdd

    Stratégies_G_End()
  End Sub

  Public Function XRoads_Vérification_XNl(XAllRoads_List As List(Of List(Of XLink_Cls))) As Boolean
    'Lien Fort Lien-Faible-Lien-Fort ...n fois... Bouclage sur lien-Faible  
    Dim Road_Numéro As Integer
    Dim Link_Numéro As Integer
    Dim Candidat As String
    Dim Road_CelDéb As Integer
    Dim Link_Fin_Prv As Integer

    Dim Unité As String()
    Dim CelWeak() As Integer
    Dim CddWeak() As String

    For Each Road As List(Of XLink_Cls) In XAllRoads_List
      If XRslt.Productivité Then Exit For

      Road_Numéro += 1
      Link_Numéro = 0
      XRslt.RoadRight.Clear()
      Road_CelDéb = -1
      Link_Fin_Prv = -1
      Candidat = XRslt.Candidat(0)
      For i As Integer = 0 To 80 : U_Road(i) = False : Next i
      XRslt.Productivité = False

      For Each Link As XLink_Cls In Road

        Link_Numéro += 1

        If Link_Numéro = 1 Then ' Premier lien
          Road_CelDéb = Link.Cel(0)
          ' Ajout systématique du premier lien Fort dans le Chemin étudié  " " 
          XRslt.RoadRight.Add(Link)
        Else                    ' Liens suivants
          Unité = Is_SameUnité(Link_Fin_Prv, Link.Cel(0))
          ' Contrôle des conditions
          If Link.Cel(0) = Link_Fin_Prv _                     ' le lien boucle sur le précédent
          Or U_Road(Link.Cel(0)) _
          Or U_Road(Link.Cel(1)) _                            ' le lien (début ou fin) est déjà passé sur ces cellules
          Or Unité(0) = "#" Then Exit For                     ' le lien faible est dans une même unité

          CelWeak = {Link_Fin_Prv, Link.Cel(0)}
          CddWeak = {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), Candidat}
          Dim LinkWeak As New XLink_Cls With {.Cel = CelWeak, .Cdd = CddWeak, .Type = "W", .Unité = Unité(0) & Unité(1)}
          XRslt.RoadRight.Add(LinkWeak) ' Ajout du lien faible
          XRslt.RoadRight.Add(Link)     ' Ajout du lien Fort
        End If '/ Premier Lien et Liens Suivants

        ' Marquage des cellules pour éviter les boucles
        U_Road(Link.Cel(0)) = True : U_Road(Link.Cel(1)) = True

        Link_Fin_Prv = Link.Cel(1)

        ' Le chaînage est-il fermé ?
        If Link_Numéro >= 2 AndAlso (Link_Numéro + 1) Mod 2 = 0 Then      ' à/p de 4 liens et un nombre de liens pairs 
          Unité = Is_SameUnité(Link.Cel(1), Road_CelDéb)
          If Unité(0) <> "#" Then  ' Le chemin est fermé
            CelWeak = {Link.Cel(1), Road_CelDéb}
            CddWeak = {Link.Cdd(0), Link.Cdd(1), Link.Cdd(2), Link.Cdd(3), Candidat}
            Dim LinkWeak As New XLink_Cls With {.Cel = CelWeak, .Cdd = CddWeak, .Type = "W", .Unité = Unité(0) & Unité(1)}
            XRslt.RoadRight.Add(LinkWeak) ' Ajout du lien faible

            ' Le chemin est-il productif ?
            If Candidats_Exclure_XNl() > 0 Then
              XRslt.Candidat(0) = Candidat
              XRslt.Candidat(1) = "0"
              XRslt.XLinks_Nombre = Road.Count
              XRslt.XRoads_Numéro = Road_Numéro
              XRslt.Productivité = True
              If Xap Then XAllRoads_List_Display(Road_Numéro)
            End If

          End If '/ Le chemin est fermé
        End If '/ le chaînage est-il bouclé ?
        If XRslt.Productivité Then Exit For
      Next Link
    Next Road
    Return XRslt.Productivité
  End Function

  Public Function Candidats_Exclure_XNl() As Integer
    For Each Link_Prd As XLink_Cls In XRslt.RoadRight
      If Link_Prd.Type = "W" Then ' Uniquement les liens faibles
        Dim idx1 As Integer = Link_Prd.Cel(0)
        Dim idx2 As Integer = Link_Prd.Cel(1)

        For Each XCel As GCel_Cls In GCels
          ' La cellule est XCel.Cel Integer
          If XCel.Cel = idx1 OrElse XCel.Cel = idx2 Then Continue For
          If Is_Vu(XCel.Cel, idx1) AndAlso Is_Vu(XCel.Cel, idx2) Then
            XRslt.CelExcl.Add(New XCel_Excl_Cls With {.Cel = XCel.Cel, .Cdd = XRslt.Candidat(0), .Exc = {idx1, idx2}})
          End If '/Is_Vu(XCel.Cel, idx1) AndAlso Is_Vu(XCel.Cel, idx2) Then
        Next XCel
      End If ' /If Link_Prd.Type = "W" 

    Next Link_Prd
    Return XRslt.CelExcl.Count
  End Function

End Module