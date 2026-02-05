Option Strict On
Option Explicit On
Imports System.Security

Friend Module En_Cours
#Region "Menus Test A à J"
  Public Sub TestA()
    Jrn_Add(, {Proc_Name_Get()})
    Pzzl_Crt_Interactif_81("P")
  End Sub
  Public Sub TestB()
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
  Public Sub RRslt_Init()
    With RRslt
      .Code_Strg = Plcy_Strg
      .Code_Sous_Strg = ""
      .Code_LCR = "#"
      .LCR = -1
      .Candidat = "0"
      .Candidats = "000000000"
      '.Cellule()
      '.CelExcl()
      .Productivité = False
    End With
  End Sub
  Public Sub RRslt_Display()

    Jrn_Add(, {"Affichage des données RRslt"})
    With RRslt
      Jrn_Add(, {"Code Stratégie " & .Code_Strg & ", " & Stg_Get(.Code_Strg).Texte})
      Jrn_Add(, {"Sous_Stratégie " & String.Join(", ", .Candidat)})
      Jrn_Add(, {"Code_LCR       " & .Code_LCR & (.LCR + 1)})
      Jrn_Add(, {"Candidat       " & .Candidat})
      'Jrn_Add(, {"Candidats      " & .Candidat})
      Jrn_Add(, {"Cellule        " & U_Coord(.Cellule(0))})
      Jrn_Add(, {"Cellules       " & String.Join(", ", .Cellule.Select(Function(c) U_Coord(c)))})
      Jrn_Add(, {"Productivité   " & .Productivité.ToString()})
    End With

  End Sub

  Public Function RRslt_Copy_Rnd(Strategy_Rslt(,) As String) As Integer
    If Strategy_Rslt Is Nothing Then
      RRslt.Productivité = False
      Return -1
    End If
    If UBound(Strategy_Rslt, 2) <= 0 Then
      RRslt.Productivité = False
      Return -1
    End If
    Dim rnd As New Random()
    Dim Index As Integer = rnd.Next(1, Strategy_Rslt.GetUpperBound(1))   ' Tire un nombre entre 1 et 99 inclus
    With RRslt
      .Code_Sous_Strg = Strategy_Rslt(2, Index)
      .Code_LCR = Strategy_Rslt(3, Index)
      .LCR = CInt(Strategy_Rslt(4, Index))
      .Candidat = Strategy_Rslt(5, Index)
      .Candidats = Strategy_Rslt(6, Index)
      ReDim .Cellule(0)
      .Cellule(0) = CInt(Strategy_Rslt(10, Index))
      .Productivité = True
    End With
    Return Index
  End Function
#End Region
  Sub Strategy_Code(strg_Code As String)
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
    If Plcy_Strg = "   " Then Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_00114", {CStr(Wh_Nb_Cell(U).Initiales), CStr(Wh_Nb_Cell(U).Vides), CStr(Wh_Grid_Nb_Candidats(U))})

    Dim U_temp(80, 3) As String
    Dim Strategy_Rslt(,) As String = Nothing
    Array.Copy(U, U_temp, UNbCopy)
    Select Case Plcy_Strg
      Case "Cdd"
      Case "CdU" : Strategy_Rslt = Strategy_CdU(U_temp)
      Case "CdO" : Strategy_Rslt = Strategy_CdO(U_temp)

      Case Else
    End Select

    Dim index As Integer = RRslt_Copy_Rnd(Strategy_Rslt)
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub

End Module