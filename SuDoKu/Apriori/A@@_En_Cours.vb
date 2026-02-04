Option Strict On
Option Explicit On

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
      .Productivité = False
    End With
    Return Index
  End Function
#End Region
  Sub Strategy_Code(strg_Code As String)
    ' Active/désactive la stratégie passée en paramètre
    Strategy_Switch(strg_Code)

    Select Case Plcy_Strg_Swt
      Case +1 : Plcy_Strg = strg_Code
      Case -1 : Plcy_Strg = "   "
    End Select
    Select Case strg_Code
      Case "CdU"
        Dim U_temp(80, 3) As String
        Dim Strategy_Rslt(,) As String
        Array.Copy(U, U_temp, UNbCopy)
        Strategy_Rslt = Strategy_CdU(U_temp)
        'Strategy_Rslt_Display(Strategy_Rslt, -1)
        Dim index As Integer = RRslt_Copy_Rnd(Strategy_Rslt)
        'Strategy_Rslt_Display(Strategy_Rslt, index)
        'RRslt_Display()
      Case Else
    End Select
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
  End Sub


  Public Sub G4_Grid_Stratégie_CdU_g(g As Graphics)
    ' La stratégie CdU calcule TOUS les CdU, UN SEUL CdU au hasard est présenté 
    Dim Cellule As Integer
    Dim Candidat As String
    If Not Plcy_Strg = "CdU" Then Exit Sub

    Try
      'Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & " sans résultat."
      Cellule = RRslt.Cellule(0)
      Candidat = RRslt.Candidat
      U_Strg_Val_Ins(Cellule) = Candidat
      G0_Cell_Figure_g(g, Cellule, "Double_Carré", Color_Stratégique)

      ' 2 Aide Graphique
      U_MdC_Init()
      G4_MdC_Row_Col_Box("Row", U_Row(Cellule))
      G4_MdC_Row_Col_Box("Col", U_Col(Cellule))
      G4_MdC_Row_Col_Box("Box", U_Reg(Cellule))
      G4_MdC_Paint_g(g) ' Les figures sont dessinées et les candidats affichés
      'Re-dessine le candidat à placer dans un cercle plein Jaune
      Dim sc As New Cellule_Cls With {.Numéro = Cellule}
      sc.G6_Cellule_Paint_Candidat_g(g, Candidat, Color_Cdd_Insérer)
      'U_Strg(Cellule) = True
      'Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte & ": " & Candidat & " jaune à placer."
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
      MsgBox(ex.Message)
    End Try
  End Sub


















End Module