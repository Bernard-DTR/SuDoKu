Option Strict On
Option Explicit On

Friend Module En_Cours
#Region "Menus Test A à J"
  Public Sub TestA()
    Jrn_Add(, {Proc_Name_Get()})
    Pzzl_Crt_Interactif_81("P")
  End Sub
  Public Sub TestB()
    '///////////////////////////////////////////////////////////////////////////////////////
    '#616 Samedi 14/02/2026
    'Je bloque temporairement la génération de grilles en arrière-plan 
    ' Pour des problèmes de stabilisation de l'application
    ' je soupçonne ce traitement de grilles en arrière-plan d'être à l'origine de certains arrêts anormaux de l'application
    '///////////////////////////////////////////////////////////////////////////////////////
    Jrn_Add_Red(Proc_Name_Get() & " #616 La génération de grilles en arrière-plan est temporairement désactivée.")
    Exit Sub

    Jrn_Add(, {Proc_Name_Get()})
    If Batch_en_Cours Then Exit Sub
    If Plcy_Generate_Batch Then
      Jrn_Add(, {"Lancement du traitement de création de grille en arrière-plan"})
      Frm_SDK.Batch_Timer.Interval = 5000
      'Process_16x est un png 16x16 32 bits
      'Me.Mnu08.Size = New System.Drawing.Size(98, 32)
      Frm_SDK.Mnu08.Image = SuDoKu.My.Resources.Resources.Process_16x
      Frm_SDK.Mnu08.Font = New Font(Frm_SDK.Mnu08.Font, FontStyle.Italic)
      Batch_Initial()
    Else
      Jrn_Add(, {"L'option Plcy_Generate_Batch est unabled."})
    End If
  End Sub
  Public Sub TestC()
    Jrn_Add(, {Proc_Name_Get()})
    '///////////////////////////////////////////////////////////////////////////////////////
    '#616 Samedi 14/02/2026
    'Je bloque temporairement la génération de grilles en arrière-plan 
    ' Pour des problèmes de stabilisation de l'application
    ' je soupçonne ce traitement de grilles en arrière-plan d'être à l'origine de certains arrêts anormaux de l'application
    ' Cette option créé des grilles en mode interactif 

    '///////////////////////////////////////////////////////////////////////////////////////
    Jrn_Add_Red(Proc_Name_Get() & " #616 Génération de grilles interactivement.")
    Dim Prd_Number_b As Integer = My.Settings.Prd_Number
    Jrn_Add_Red(CStr(Prd_Number_b).PadLeft(5, CChar("0")))
    Application.DoEvents()

    Frm_SDK.Cursor = Cursors.WaitCursor
    Batch_Sudoku()
    Frm_SDK.Cursor = Cursors.Default

    Dim Prd_Number_a As Integer = My.Settings.Prd_Number
    Jrn_Add_Red(CStr(Prd_Number_a).PadLeft(5, CChar("0")))
    Jrn_Add_Red(Proc_Name_Get() & " #616 Fin de génération de grilles interactivement.")


  End Sub
  Public Sub TestD()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestE()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestF()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestG()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestH()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestI()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestJ()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
#End Region
#Region "RRslt: Les Résultats d'une stratégie 'Classique'"
  Public Sub RRslt_Display()
    With RRslt
      Jrn_Add(, {"Affichage des données RRslt _ Occurence " & .Occurence & " / " & .Nb_Occurences})
      Jrn_Add(, {"Code Stratégie " & .Code_Strg & ", " & Stg_Get(.Code_Strg).Texte})
      Jrn_Add(, {"Sous_Stratégie " & .Code_Sous_Strg})
      Jrn_Add(, {"Code_LCR_LCR   " & .Code_LCR & (.LCR + 1)})
      Jrn_Add(, {"Candidat       " & .Candidat})
      If .Cellule IsNot Nothing Then Jrn_Add(, {"Cellule        " & String.Join(", ", .Cellule.Select(Function(c) U_Coord(c)))})
      If .CelExcl IsNot Nothing Then Jrn_Add(, {"CelExcl        " & String.Join(", ", .CelExcl.Select(Function(c) U_Coord(c)))})
      Jrn_Add(, {"Productivité   " & .Productivité.ToString()})
    End With

  End Sub

#End Region
  Sub Strategy_Code(strg_Code As String, Origine As String)
    ' TOUTES les stratégies de la barre d'outils et les 2 stratégies DCd et CdS passent ici
    ' 1 Active/désactive la stratégie passée en paramètre
    ' 2 Documente Plcy_Strg ou la remet à blanc
    ' 3 Affiche dans B_Info soit le texte de la stratégie, soit les infos de la grille
    ' 4 Lance éventuellement les calculs de la stratégie et en récupère un seul résultat au hasard
    ' 5 Invalide le formulaire pour déclencher un repaint et l'affichage du résultat

    Strategy_Switch(strg_Code)
    Select Case Plcy_Strg_Swt
      Case +1 : Plcy_Strg = strg_Code
      Case -1 : Plcy_Strg = "   "
    End Select

    If Plcy_Strg <> "   " Then Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte
    If Plcy_Strg = "   " Then Frm_SDK.B_Info.Text = Msg_Read("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})
    Jrn_Add(, {Proc_Name_Get() & " " & strg_Code & " " & Origine})
    Dim U_temp(80, 3) As String
    Dim Strategy_Rslt(,) As String = Nothing
    Array.Copy(U, U_temp, UNbCopy)
    Select Case Plcy_Strg
      Case "Cdd"
      Case "CdU" : Strategy_Rslt = Strategy_CdU(U_temp)
      Case "CdO" : Strategy_Rslt = Strategy_CdO(U_temp)
      Case "FC1", "FC2", "FC3", "FC4", "FC5", "FC6", "FC7", "FC8", "FC9"
      Case "FV1", "FV2", "FV3", "FV4", "FV5", "FV6", "FV7", "FV8", "FV9"
      Case "Cbl" : Strategy_Rslt = Strategy_Cbl(U_temp)
      Case "Tpl" : Strategy_Rslt = Strategy_Tpl(U_temp)
      Case "Xwg" : Strategy_Rslt = Strategy_Xwg(U_temp)
      Case Else
    End Select

    Dim index As Integer = RRslt_Copy_Rnd(Strategy_Rslt)
    Event_OnPaint_MAP = Proc_Name_Get() & " Plcy_Strg: '" & Plcy_Strg & "'"
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub

  Public Function RRslt_Copy_Rnd(Strategy_Rslt(,) As String) As Integer
    ' La procédure documente RRSlt d'un résultat pris au hasard de Strategy_Rslt
    If Strategy_Rslt Is Nothing OrElse UBound(Strategy_Rslt, 2) <= 0 Then     'Strategy_Rslt.GetLength(1) = 0
      RRslt.Productivité = False
      Return -1
    End If

    Dim rnd As New Random()
    Dim Index As Integer = rnd.Next(1, Strategy_Rslt.GetUpperBound(1) + 1)   ' Tire un nombre entre min inclus et max non inclus
    With RRslt
      .Occurence = Index
      .Nb_Occurences = Strategy_Rslt.GetUpperBound(1)
      .Code_Strg = Strategy_Rslt(1, Index)
      .Code_Sous_Strg = Strategy_Rslt(2, Index)
      .Code_LCR = Strategy_Rslt(3, Index)
      .LCR = CInt(Strategy_Rslt(4, Index))
      .Candidat = Strategy_Rslt(5, Index)

      Dim valeurs As New List(Of Integer)
      For k As Integer = 0 To 44
        Dim s As String = Strategy_Rslt(10 + k, Index)
        If s = "__" Then Exit For
        valeurs.Add(CInt(s))
      Next
      .Cellule = valeurs.ToArray()
      valeurs.Clear()
      For k As Integer = 0 To 44
        Dim s As String = Strategy_Rslt(55 + k, Index)
        If s = "__" Then Exit For
        valeurs.Add(CInt(s))
      Next
      .CelExcl = valeurs.ToArray()

      .Productivité = True
    End With
    Strategy_Rslt_Display(Strategy_Rslt, -1)
    RRslt_Display()
    Return Index

  End Function

End Module