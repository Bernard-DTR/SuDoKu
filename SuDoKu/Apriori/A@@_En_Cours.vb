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
    Dim Grid As String = ClipBoard_Copier_New("1")
    Jrn_Add_Yellow(Grid)
    Dim AllCandidates(728) As Candidate
    AllCandidates_Init(AllCandidates)


    SDK_AllCandidate(Grid, AllCandidates)
    AllCandidates_Display(AllCandidates)

    AllCandidates_Display_IsSolved(AllCandidates)

  End Sub
  Public Sub TestG()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestH()
    Jrn_Add(, {Proc_Name_Get()})
  End Sub
  Public Sub TestI()
    Jrn_Add(, {Proc_Name_Get()})
    Dim Grid As String = ClipBoard_Copier_New("1")
    Jrn_Add_Yellow(Grid)
    Dim AllCandidates(728) As Candidate
    AllCandidates_Init(AllCandidates)
    SDK_AllCandidate(Grid, AllCandidates)
    'AllCandidates_Display(AllCandidates)
    PropagateSolvedCandidates(AllCandidates)
    'AllCandidates_Display_IsSolved(AllCandidates)
    DB_Solution(AllCandidates)
    AllCandidates_SDK(AllCandidates)
  End Sub
  Public Sub TestJ()
    Jrn_Add(, {Proc_Name_Get()})
    Dim U_temp(80, 3) As String

    Plcy_Strg = "GCx"
    Jrn_Add(, {"Strategy_" & Plcy_Strg})
    Array.Copy(U, U_temp, UNbCopy)
    Strategy_GCx(U_temp)
    Frm_SDK.B_Solution.Text = Stg_Get(Plcy_Strg).Family.ToString()
    Frm_SDK.B_Info.Text = Stg_Get(Plcy_Strg).Texte
  End Sub
#End Region
End Module