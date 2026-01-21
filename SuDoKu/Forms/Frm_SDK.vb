Option Strict On
Option Explicit On
Imports System.Threading
Imports SuDoKu.DancingLink

'-------------------------------------------------------------------------------
'
' Formulaire SDK
'
'-------------------------------------------------------------------------------

Public NotInheritable Class Frm_SDK
  Private MouseClick_Middle_ToolTip As CustomToolTip

  Public Journal As New RichTextBox()
  Public B_Solution As New TextBox
  Public B_Position As New TextBox()
  Public B_Pourcentage As New TextBox()
  Public B_Info As New TextBox()

  Public B_ProgressBar As New ProgressBar()
  'La ProgressBar ne peut pas adopter la couleur souhaitée
  Dim Prv_MM_Pt As Point
  Dim Prv_Rct_Cdd_Numéro As Integer

  Dim Key_Numlock As Boolean = False
  Dim Prv_Key_Numlock As Boolean = False
  Dim Key_CapsLock As Boolean = False
  Dim Prv_Key_CapsLock As Boolean = False

  Dim Key_CtrlDown As Boolean = False
  Dim Prv_Key_CtrlDown As Boolean = False

  Public Sub New()
    ' Cet appel est requis par le concepteur.
    InitializeComponent()
  End Sub

  Private Sub Frm_SDK_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'Interruption temporaire de la logique de présentation du formulaire, pendant la mise en place des contrôles
    SuspendLayout()
    Phase_Démarrage_Terminée = False
    OO_000_SDK_Load()
    AutoScaleMode = Me_AutoScaleMode_Standard
    Size = New Size(1692, 1036) 'Taille maximale :SDK   
    SetStyle(ControlStyles.UserPaint Or ControlStyles.AllPaintingInWmPaint, True)
    '        ControlStyles.OptimizedDoubleBuffer, True empêche l'affichage de la grille
    UpdateStyles()

#Region "Contrôles B_*"
    'Mise en place des contrôles B_* de Frm_SDK
    'Il n'est pas possible de copier-coller des infos du journal, sauf avec l'opération Copier le Grid qui effectue un ReadOnly On-Off
    '  La grille ne comporte AUCUN contrôle susceptible d'être écrit, donc on intercepte une touche de 1 à 9 pour afficher une valeur
    '  Le journal ne peut donc pas être "écrit"
    '  La scrollBar verticale est comprise dans la largeur du journal, affichée ou non.
    With Journal
      .Name = "Journal"
      .BackColor = Color_Fond_Typ_I
      .Font = Font_Journal
      .Multiline = True
      .ScrollBars = RichTextBoxScrollBars.ForcedVertical 'Y compris si le texte est inférieur au contrôle
      .MaxLength = 262144
      .ReadOnly = True
      .ContextMenuStrip = Mnu_Journal
      .DetectUrls = True  'Pour afficher des URL's
      '.HideSelection = True
    End With
    Controls.Add(Journal)
    AddHandler Journal.LinkClicked, AddressOf Journal_Linkcliked
    AddHandler Journal.Click, AddressOf Journal_Click

    With B_Solution
      .Name = "B_Solution"
      .ReadOnly = True
      .ShortcutsEnabled = False 'Pour ne pas avoir de menu contextuel
      .TextAlign = HorizontalAlignment.Center
    End With
    Controls.Add(B_Solution)

    With B_Position
      .Name = "B_Position"
      .ReadOnly = True
      .ShortcutsEnabled = False
      .TextAlign = HorizontalAlignment.Center
    End With
    Controls.Add(B_Position)

    With B_Pourcentage
      .Name = "B_Pourcentage"
      .ReadOnly = True
      .ShortcutsEnabled = False
      .TextAlign = HorizontalAlignment.Center
    End With
    Controls.Add(B_Pourcentage)

    'B_Info est affiché en alternance avec B_ProgressBar qui affiche la Progress Bar
    With B_Info
      .Name = "B_Info"
      .ReadOnly = True
      .ShortcutsEnabled = False
      .TextAlign = HorizontalAlignment.Center
    End With
    Controls.Add(B_Info)

    With B_ProgressBar
      .Name = "B_ProgressBar"
      .ForeColor = SystemColors.Highlight
      'La barre de progression ne supporte pas .backcolor
      .Minimum = 0
      .Maximum = 100
      .Value = 0
      .Step = 1
    End With
    Controls.Add(B_ProgressBar)
#End Region

#Region "Mnu04_Stratégies"
    'Text et TTT des Btn_Stratégies
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Dsp_BO = "O" Then
        Dim Btn_Name As String = "Btn_" & Stg.Code
        If BarreOutils.Items.ContainsKey(Btn_Name) Then
          BarreOutils.Items(Btn_Name).Text = Stg.Lettre
          BarreOutils.Items(Btn_Name).ToolTipText = Stg.Texte
        End If
      End If
    Next

    'Stg_List_Code: CdU , CdO, Cbl, Tpl, Xwg, XYw, Swf, Jly, XYZ, SKy, Unq
    For i As Integer = 0 To Stg_Bll.Count - 1
      Dim Btn_Name As String = "Btn_" & Stg_List_Code.Item(i)
      If BarreOutils.Items.ContainsKey(Btn_Name) Then
        If Stg_Bll(i) = True Then BarreOutils.Items(Btn_Name).Enabled = True
        If Stg_Bll(i) = False Then BarreOutils.Items(Btn_Name).Enabled = False
      End If
    Next

    'Le menu 04 diffère des autres menus dans la mesure où il est capable d'afficher une image ET le symbole checked/non checked
    'https://docs.microsoft.com/en-us/dotnet/desktop/winforms/controls/how-to-enable-check-margins-and-image-margins-in-contextmenustrip-controls?view=netframeworkdesktop-4.8
    Mnu04.DropDown = New ContextMenuStrip()
    CType(Mnu04.DropDown, ContextMenuStrip).ShowImageMargin = True
    CType(Mnu04.DropDown, ContextMenuStrip).ShowCheckMargin = True

    Dim Mnu04n_Cdd As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Cdd").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Cdd.Click, AddressOf Mnu04n_Cdd_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Cdd)

    Dim Mnu04n_CdU As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("CdU").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_CdU.Click, AddressOf Mnu04n_CdU_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_CdU)

    Dim Mnu04n_CdO As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("CdO").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_CdO.Click, AddressOf Mnu04n_CdO_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_CdO)

    Dim Mnu04n_Sep01 As New ToolStripSeparator
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Sep01)

    Dim Mnu04n_Cbl As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Cbl").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Cbl.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Cbl)

    Dim Mnu04n_Tpl As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Tpl").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Tpl.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Tpl)

    Dim Mnu04n_Xwg As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Xwg").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Xwg.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Xwg)

    Dim Mnu04n_XYw As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("XYw").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_XYw.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_XYw)

    Dim Mnu04n_Swf As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Swf").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Swf.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Swf)

    Dim Mnu04n_Jly As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Jly").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Jly.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Jly)

    Dim Mnu04n_XYZ As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("XYZ").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_XYZ.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_XYZ)

    Dim Mnu04n_SKy As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("SKy").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_SKy.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_SKy)

    Dim Mnu04n_Unq As New ToolStripMenuItem() With
        {
        .Text = Stg_Get("Unq").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_Unq.Click, AddressOf Mnu04n_Stratégie_BTXYSJZKQ
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Unq)

    Dim Mnu04n_AnnulerLaDerniereOption As New ToolStripMenuItem() With
        {
        .Text = "Annuler la dernière option",
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_AnnulerLaDerniereOption.Click, AddressOf Mnu04n_AnnulerLaDerniereOption_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_AnnulerLaDerniereOption)

    Dim Mnu04n_Sep02 As New ToolStripSeparator
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Sep02)

    Dim Mnu04n_XCx As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "XCx",
        .Text = Stg_Get("XCx").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_XCx.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_XCx)

    Dim Mnu04n_XCy As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "XCy",
        .Text = Stg_Get("XCy").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_XCy.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_XCy)

    Dim Mnu04n_XRp As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "XRp",
        .Text = Stg_Get("XRp").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_XRp.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_XRp)

    Dim Mnu04n_XNl As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "XNl",
        .Text = Stg_Get("XNl").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_XNl.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_XNl)

    Dim Mnu04n_WgX As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "WgX",
        .Text = Stg_Get("WgX").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_WgX.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_WgX)

    Dim Mnu04n_WgY As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "WgY",
        .Text = Stg_Get("WgY").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_WgY.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_WgY)

    Dim Mnu04n_WgZ As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "WgZ",
        .Text = Stg_Get("WgZ").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_WgZ.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_WgZ)

    Dim Mnu04n_WgW As New ToolStripMenuItem() With
        {
        .Name = "Mnu04n_" & "WgW",
        .Text = Stg_Get("WgW").Texte,
        .ForeColor = SystemColors.ControlText
         }
    AddHandler Mnu04n_WgW.Click, AddressOf Mnu04n_Stratégie_XW_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_WgW)

    Mnu07n_Gbl.Text = Stg_Get("Gbl").Texte
    Mnu07n_Gbv.Text = Stg_Get("Gbv").Texte
    Mnu07n_GCs.Text = Stg_Get("GCs").Texte
    Mnu07n_XCx.Text = Stg_Get("XCx").Texte
    Mnu07n_XCy.Text = Stg_Get("XCy").Texte
    Mnu07n_XRp.Text = Stg_Get("XRp").Texte
    Mnu07n_XNl.Text = Stg_Get("XNl").Texte
    Mnu07n_WgX.Text = Stg_Get("WgX").Texte
    Mnu07n_WgY.Text = Stg_Get("WgY").Texte
    Mnu07n_WgZ.Text = Stg_Get("WgZ").Texte
    Mnu07n_WgW.Text = Stg_Get("WgW").Texte

    Mnu0902.Enabled = False
    Mnu0910_GLk.Text = Stg_Get("GLk").Texte & "..."
    Mnu0930_Gbl.Text = Stg_Get("Gbl").Texte
    Mnu0950_Gbv.Text = Stg_Get("Gbv").Texte
    Mnu0970_GCs.Text = Stg_Get("GCs").Texte

    Dim Mnu04n_Sep03 As New ToolStripSeparator
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Sep03)

    Dim Mnu04n_RésoudreUneCellule As New ToolStripMenuItem() With
        {
        .Text = "Résoudre Une Cellule",
        .ForeColor = SystemColors.ControlText,
        .ShortcutKeys = Keys.F11
         }
    AddHandler Mnu04n_RésoudreUneCellule.Click, AddressOf Mnu04n_RésoudreUneCellule_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_RésoudreUneCellule)

    Dim Mnu04n_Suggérer As New ToolStripMenuItem() With
        {
        .Text = "Suggérer",
        .ForeColor = SystemColors.ControlText,
        .ShortcutKeys = Keys.Shift Or Keys.F11
        }
    AddHandler Mnu04n_Suggérer.Click, AddressOf Mnu04n_Suggérer_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Suggérer)

    Dim Mnu04n_Sep04 As New ToolStripSeparator
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Sep04)

    Dim Mnu04n_MettreEnÉvidenceLaDernièreCellule As New ToolStripMenuItem() With
      {
      .Name = "Mnu04n_DCd",
      .Text = "Mettre en évidence La Dernière Valeur",
      .Image = SuDoKu.My.Resources.Resources.Terre,
      .Checked = False,
      .CheckOnClick = True,
      .ForeColor = SystemColors.ControlText
      }
    AddHandler Mnu04n_MettreEnÉvidenceLaDernièreCellule.Click, AddressOf Mnu04n_MettreEnÉvidenceLaDernièreCellule_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_MettreEnÉvidenceLaDernièreCellule)

    Dim Mnu04n_MettreEnÉvidenceLeCandidatSaisi As New ToolStripMenuItem() With
      {
      .Name = "Mnu04n_CdS",
      .Text = "Mettre en évidence Le Candidat Saisi",
      .Checked = False,
      .CheckOnClick = True,
      .ForeColor = SystemColors.ControlText
      }
    AddHandler Mnu04n_MettreEnÉvidenceLeCandidatSaisi.Click, AddressOf Mnu04n_MettreEnÉvidenceLeCandidatSaisi_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_MettreEnÉvidenceLeCandidatSaisi)


    Dim Mnu04n_Sep05 As New ToolStripSeparator
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_Sep05)

    Dim Mnu04n_ModeSuggestion As New ToolStripMenuItem() With
        {
        .Text = "Mode Suggestion",
        .Checked = False,
        .CheckOnClick = True,
        .ForeColor = SystemColors.ControlText,
        .CheckState = 0
         }
    'C'est la propriété CheckOnClick qui gère automatiquement la coche
    AddHandler Mnu04n_ModeSuggestion.Click, AddressOf Mnu04n_ModeSuggestion_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_ModeSuggestion)
#End Region

#Region "Mnu06_Divers Internet Explorer et Url diverses"
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE1"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE2"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE3"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE4"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE5"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE6"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE7"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE8"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IE9"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IEa"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IEb"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IEc"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read_IA("MNU_06IEd"))
    Mnu06_CB.AllowDrop = False
    Mnu06_CB.AutoSize = True
    Mnu06_CB.Width = 450
    'Rappel de la dernière Url sélectionnée
    Dim Index06_CB As Integer = My.Settings.SDK_IE_Last_Url
    Mnu06_CB.SelectedIndex = Index06_CB
#End Region

#Region "ToolTip_B "
    'Il existe 1 Contrôle ToolTip_B  pour les zones B_* 
    With ToolTip_B
      .SetToolTip(B_Position, "Cellule sélectionnée en Ligne-Colonne")
      .SetToolTip(B_Pourcentage, "Taux de remplissage")
      .SetToolTip(B_Info, "Information")
    End With
#End Region

#Region "Derniers réglages"
    'Il est créé un menu contextuel Vide pour ne pas afficher de menu au clic droit  
    '   dans la barre de menus et dans la barre d'outils.
    ' Ces informations sont NECESSAIRES et sont placées dans les PROPRIETES
    Mnu.ContextMenuStrip = Mnu_Ctxt_Vide
    BarreOutils.ContextMenuStrip = Mnu_Ctxt_Vide
    ContextMenuStrip = Mnu_Cel                               ' Définition du Menu Contextuel pour les cellules
    AllowDrop = False                                        ' Il n'a pas de copier-coller
    FormBorderStyle = FormBorderStyle.FixedDialog
    MaximizeBox = False
    MinimizeBox = True
    Icon = My.Resources.SuDoKu
    Top = 10
    Left = 10

#End Region

    OO_999_SDK_Load_End()

#Region "Lancement de la partie précédente"
    Jrn_Add(, {"Application " & Application.ProductName & " " & SDK_Version})
    Jrn_Add(, {"Chemin      " & Path_SDK})
    Jrn_Add(, {"Traitement de création de grilles en Batch : " & CStr(Plcy_Generate_Batch)})
    Jrn_Add(, {"Nombre de grilles à créer en Batch         : " & CStr(My.Settings.Prf_02C_Nb_Batch_Generate)})
    Jrn_Add(, {"Nombre de grilles existantes dans " & Path_Batch})
    Jrn_Add(, {"          Faciles                        F : " & CStr(File_Nb_IA("SDK_F"))})
    Jrn_Add(, {"          Moyennes                       M : " & CStr(File_Nb_IA("SDK_M"))})
    Jrn_Add(, {"          Difficiles                     D : " & CStr(File_Nb_IA("SDK_D"))})
    Jrn_Add(, {"          Expertes                       E : " & CStr(File_Nb_IA("SDK_E"))})
    Jrn_Add(, {Procédure_Name_Get()})
    Jrn_Add("SDK_00010", JourDateHeure())
    Jrn_Add("SDK_00101")
    Plcy_Gnrl = My.Settings.LP_Plcy_Gnrl
    LP_Nom = My.Settings.LP_Nom
    LP_Prb = My.Settings.LP_Prb
    LP_Jeu = My.Settings.LP_Jeu
    LP_Sol = My.Settings.LP_Sol
    LP_Frc = My.Settings.LP_Frc
    Select Case Plcy_Gnrl
      Case "Nrm"
        LP_Cdd = ""
        LP_CddExc = ""
      Case "Sas"
        LP_Cdd = My.Settings.LP_Cdd.Replace("0", " ")
        LP_CddExc = My.Settings.LP_CddExc.Replace("0", " ")
    End Select
    Jrn_Add("SDK_00100", {LP_Nom})
    Jrn_Add(, {"/" & Procédure_Name_Get()})
    Game_New_Game(Plcy_Gnrl, LP_Nom, LP_Prb, LP_Jeu, LP_Sol, LP_Cdd, LP_Frc)
#End Region

    'Reprise de la logique de présentation du formulaire, les contrôles sont mis en place
    Phase_Démarrage_Terminée = True
    ResumeLayout()

    'Plcy_Generate_Batch autorise la création de grilles par lot en arrière-plan
    If Plcy_Generate_Batch Then
      Batch_Timer.Interval = 5000
      'Process_16x est un png 16x16 32 bits
      'Me.Mnu08.Size = New System.Drawing.Size(98, 32)
      Mnu08.Image = SuDoKu.My.Resources.Resources.Process_16x
      Mnu08.Font = New Font(Mnu08.Font, FontStyle.Italic)
      Batch_Initial()
    End If
  End Sub

  Private Sub Batch_Timer_Tick(sender As Object, e As EventArgs) Handles Batch_Timer.Tick
    If Mnu08.Image Is Nothing Then Exit Sub
    Select Case Batch_en_Cours
      Case True
      Case False
        Mnu08.Image = Nothing
        Mnu08.Font = New Font(Mnu08.Font, FontStyle.Regular)
        Batch_Timer.Interval = 50000
    End Select
  End Sub
  Protected Overrides Sub OnPaint(e As PaintEventArgs)
    ' Ajout du 26/10/2025
    ' Cela garantit que la logique de peinture de la classe de base (le cas échéant) est exécutée.
    ' Par exemple, elle peut gérer la peinture de l'arrière-plan ou d'autres comportements par défaut
    MyBase.OnPaint(e)
    If Not Phase_Démarrage_Terminée Then Exit Sub

    'SourceOver signifie que les pixels de la source (ce qui est dessiné) sont composés au-dessus des pixels de la destination (le formulaire)
    e.Graphics.CompositingMode = Drawing2D.CompositingMode.SourceOver
    e.Graphics.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
    e.Graphics.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
    e.Graphics.PixelOffsetMode = Drawing2D.PixelOffsetMode.HighQuality
    e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
    e.Graphics.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

    Select Case Event_OnPaint
      Case "Frm_SDK_Activated",
           "Frm_SDK_Resize",
           "Frm_SDK_MinimumSizeChanged",
           "Frm_SDK_LocationChanged",
           "Frm_SDK_Move",
           "Frm_SDK_SizeChanged",
           "Frm_SDK_VisibleChanged",
           "Game_New_Game",
           "Solution_Affichée",
           "Frm_Préférences",
           "Stratégie_X",
           "AfficherDCdd",
           "Batch_Sudoku",
           "Frm_SDK_FormClosing"
        Dim Gril As New Grille_Cls
        Gril.Grille_Refresh()
        Dim sc As New Cellule_Cls With {.Numéro = Pbl_Cell_Select}
        sc.Cellule_Refresh()
        sc.G7_Cellule_Paint_Select()
        '#529
      Case "Stratégie_G"
        Jrn_Add_Yellow(Procédure_Name_Get() & " Plcy_strg=" & Plcy_Strg & " Event_OnPaint=" & Event_OnPaint)
        Dim Gril As New Grille_Cls
        Gril.Grille_Refresh_g(e.Graphics)
        Dim sc As New Cellule_Cls With {.Numéro = Pbl_Cell_Select}
        sc.Cellule_Refresh()
        sc.G7_Cellule_Paint_Select()
      Case Else
        Jrn_Add("SDK_00000", {"Protected Overrides Sub OnPaint(e As PaintEventArgs) est activée: "})
        Jrn_Add("SDK_00000", {"e As PaintEventArgs      : " & e.ToString})
        Jrn_Add("SDK_00000", {"Valeur Event_OnPaint     : " & Event_OnPaint})
        Jrn_Add("SDK_00000", {"Valeur Event_OnPaint_MAP : " & Event_OnPaint_MAP})
        Jrn_Add("SDK_00000", {"La grille n'est pas rafraîchie, La cellule " & U_cr(Pbl_Cell_Select) & " n'est pas sélectionnée."})
        Jrn_Add("SDK_00010", JourDateHeure())
    End Select
    Event_OnPaint = "#"
  End Sub
  Private Sub Frm_SDK_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
    Event_OnPaint = "Frm_SDK_Activated"
  End Sub

  Private Sub Frm_SDK_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
    'En plaçant à cet endroit l'enregistrement des LP_*, 
    'Ceux-ci sont enregistrés correctement pour les 12 ouvertures de jeux de SDK.
    Dim P As String = ""
    Dim J As String = ""
    Dim S As String = ""
    For i As Integer = 0 To 80
      P &= U(i, 1)                                ' Problème
      J &= If(U(i, 2) <> " ", U(i, 2), U(i, 1))   ' Jeu
      S &= U_Sol(i)                               ' Solution
    Next i
    My.Settings.LP_Nom = LP_Nom
    My.Settings.LP_Frc = LP_Frc
    My.Settings.LP_Prb = P.Replace(" ", "0")    '0 remplace une valeur vide
    My.Settings.LP_Jeu = J.Replace(" ", "0")
    My.Settings.LP_Sol = S.Replace(" ", "0")
    'Projet / Propriété de SuDoKu ... / Paramètres pour créer les nouveaux paramètres
    'SuDoKu est ouvert soit en Mode Nrm, soit en mode Sas
    Select Case Plcy_Gnrl
      Case "Sas" : My.Settings.LP_Plcy_Gnrl = "Sas"
      Case Else : My.Settings.LP_Plcy_Gnrl = "Nrm"
    End Select

    ' En Mode Sas seulement, les candidats sont enregistrés, LP_Cdd ne sera utilisé que dans le cadre Sas        
    LP_Cdd = ""                ' Les candidats  
    LP_CddExc = ""             ' Les candidats exclus
    For i As Integer = 0 To 80
      Select Case Plcy_Gnrl
        Case "Sas"
          LP_Cdd &= U(i, 3)
          LP_CddExc &= U_CddExc(i)
        Case Else 'Cas Nrm
          LP_Cdd &= Cnddts_Blancs
          LP_CddExc &= Cnddts_Blancs
      End Select
    Next i
    My.Settings.LP_Cdd = LP_Cdd.Replace(" ", "0")
    My.Settings.LP_CddExc = LP_CddExc.Replace(" ", "0")
    My.Settings.Save()

    If Batch_Thread IsNot Nothing AndAlso Batch_Thread.IsAlive Then
      ' Optionnel : afficher un message ou une animation d'attente
      Enabled = False
      Cursor = Cursors.WaitCursor

      ' Attendre la fin du thread
      Event_OnPaint = "Frm_SDK_FormClosing"
      Invalidate()
      MsgBox("SuDoKu est en train de calculer des grilles " & vbCrLf & "Merci de patienter ! ",
      MsgBoxStyle.Information, "SuDoKu")
      Batch_Thread.Join()

      ' Restaurer l'état de l'interface
      Cursor = Cursors.Default
      Enabled = True
    End If
  End Sub
  Private Sub Frm_SDK_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
  End Sub
  Private Sub Frm_SDK_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
    'Se produit quand le contrôle est redimensionné par exemple après une Réduction 
    Event_OnPaint = "Frm_SDK_Resize"
  End Sub
  Private Sub Frm_SDK_MinimumSizeChanged(sender As Object, e As EventArgs) Handles MyBase.MinimumSizeChanged
    Event_OnPaint = "Frm_SDK_MinimumSizeChanged"
  End Sub
  Private Sub Frm_SDK_LocationChanged(sender As Object, e As EventArgs) Handles MyBase.LocationChanged
    Event_OnPaint = "Frm_SDK_LocationChanged"
  End Sub
  Private Sub Frm_SDK_Move(sender As Object, e As EventArgs) Handles MyBase.Move
    'Génére OnPaint, lorsque le Move est hors écran
    Event_OnPaint = "Frm_SDK_Move"
  End Sub
  Private Sub Frm_SDK_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged
    Event_OnPaint = "Frm_SDK_SizeChanged"
  End Sub
  Private Sub Frm_SDK_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
    Event_OnPaint = "Frm_SDK_VisibleChanged"
  End Sub

#Region "Clavier et Mouse Clic"
  Private Sub Frm_SDK_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
    ' Le contrôle a le focus
    ' La touche est enfoncée pour la première fois
    ' Il faut que la cellule précisée par la souris soit correcte
    Dim ToolTipText As String = Nothing
    'Test du Clavier Numérique
    'Le clavier numérique permet soit de saisir des valeurs, soit de se positionner sur une cellule et de se déplacer
    If My.Computer.Keyboard.NumLock Then
      Key_Numlock = True
    Else
      Key_Numlock = False
    End If
    Prv_Key_Numlock = Key_Numlock

    'Test du Clavier Majuscule
    'Bien que le mode Clavier Majuscule ne serve à rien, je le conserve
    'Sur le Clavier (Ligne 2) la séquence Maj+9 donne 9, la touche "ç" donne 9
    If My.Computer.Keyboard.CapsLock Then
      Key_CapsLock = True
    Else
      Key_CapsLock = False
    End If
    If Key_CapsLock <> Prv_Key_CapsLock Then    'Le message n'est donné qu'en cas de changement
      If My.Computer.Keyboard.CapsLock Then
        Jrn_Add("SDK_00170")        'SDK_00170 = Le clavier est en Mode Majuscule
      Else
        Jrn_Add("SDK_00171")        'SDK_00171 = Le clavier est en Mode Minuscule
      End If
    End If
    Prv_Key_CapsLock = Key_CapsLock

    'La sélection multiple est faite indifféremment avec Shift ou Ctrl .
    Dim Cellule_KD As Integer = Pbl_Cell_Select
    Try
      Select Case e.KeyCode.ToString()
        Case "Escape"                ' Rien
        Case "Capital"               ' Rien
        Case "NumLock"               ' Rien
        Case "CapsLock"              ' Rien
        Case "ShiftKey"              ' Rien
        Case "ControlKey"            ' Rien

                                             ' Ces 2 actions sont effectuées par les raccourcis des menus.
        Case "Insert"                ' ACTION Il faut activer FN + INS ou INS   Keyboard Numérique (Mode déplacement)
        Case "Delete"                ' ACTION Il faut activer SUPPR    ou SUPPR Keyboard Numérique (Mode déplacement)
'               --------------------------------------------------------------------------------------

        Case "Home"                  ' POSITION en haut à gauche
          Cellule_KD = 0
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "End"                   ' POSITION en bas  à gauche
          Cellule_KD = 72
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "PageUp"                ' POSITION en haut à droite
          Cellule_KD = 8
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "Next"                  ' POSITION en bas  à droite
          Cellule_KD = 80
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "Clear"                 ' POSITION au centre
          Cellule_KD = 40
          Frm_SDK_KeyDown_Selected(Cellule_KD)

        Case "Left", "Back"          ' POSITION Déplacement à gauche
          Cellule_KD -= 1 : If Cellule_KD < 0 Then Cellule_KD = 80
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "Right"                 ' POSITION Déplacement à droite
          Cellule_KD += 1 : If Cellule_KD > 80 Then Cellule_KD = 0
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "Up"                    ' POSITION Déplacement vers le haut
          Dim Row As Integer = U_Row(Cellule_KD)
          Dim Col As Integer = U_Col(Cellule_KD)
          Select Case Row
            Case 0
              If Col = 0 Then Cellule_KD = Wh_Cellule_ColRow(8, 8)
              If Col <> 0 Then Cellule_KD = Wh_Cellule_ColRow(Col - 1, 8)
            Case Else
              Cellule_KD -= 9
          End Select
          Frm_SDK_KeyDown_Selected(Cellule_KD)
        Case "Down"                  ' POSITION Déplacement vers le bas
          Dim Row As Integer = U_Row(Cellule_KD)
          Dim Col As Integer = U_Col(Cellule_KD)
          Select Case Row
            Case 8
              If Col = 8 Then Cellule_KD = Wh_Cellule_ColRow(0, 0)
              If Col <> 8 Then Cellule_KD = Wh_Cellule_ColRow(Col + 1, 0)
            Case Else
              Cellule_KD += 9
          End Select
          Frm_SDK_KeyDown_Selected(Cellule_KD)

'               --------------------------------------------------------------------------------------

        Case "NumPad1" To "Numpad9"  ' INSERTION
          Dim Valeur As Integer = e.KeyCode - 96     ' Touches du clavier numérique de 1 à 9
          Cell_Val_Insert(CStr(Valeur), Cellule_KD, "Kbd_Num")
        Case "D1" To "D9"            ' INSERTION
          Dim Valeur As Integer = e.KeyCode - 48      ' Touches du clavier alphanumérique de 1 à 9 (MAJ OFF ou ON)
          Cell_Val_Insert(CStr(Valeur), Cellule_KD, "Kbd_Alp")

        Case Else
          'HPOmen16 la touche e.KeyCode.ToString() = LaunchApplication2
          '                   e.KeyData = 183 lance l'application Calc
      End Select

      'Ce n'est pas la touche qui entre la valeur, mais son effet true
      'il y a donc un son d'erreur qu'il faut enlever
      e.SuppressKeyPress = True            'Pour supprimer le son, car il n'y a AUCUNE zone d'entrée

    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {"Cellule_KD :  " & CStr(Cellule_KD)})
      Jrn_Add("ERR_00000", {Procédure_Name_Get() & " hors square "})
    End Try

  End Sub
  Private Sub Frm_SDK_KeyDown_Selected(Cellule_KDS As Integer)
    ' Test des sélections multiples
    '     La sélection multiple est indifférente à Ctrl ou à Shift
    With My.Computer.Keyboard
      If Plcy_Gnrl = "Sas" Then
        If .CtrlKeyDown Or .ShiftKeyDown Then Plcy_Slm = True
      End If
      If Not .CtrlKeyDown And Not .ShiftKeyDown Then Plcy_Slm = False
    End With

    'La cellule sur laquelle passe la souris devient la Pbl_Cell_Select
    Pbl_Cell_Select = Cellule_KDS

    Dim sc As New Cellule_Cls With {.Numéro = Pbl_Cell_Select}
    Select Case Plcy_Slm
      Case True
        '1 On sélectionne la cellule  
        U_Slm(Pbl_Cell_Select) = True
        sc.Cellule_Refresh()
        sc.G7_Cellule_Paint_Select()
      Case False
        '0 Si slm, on remet les cellules précédentes en état
        If Prv_Plcy_Slm Then
          Dim sc_Slm As New Cellule_Cls
          For i As Integer = 0 To 80
            If U_Slm(i) Then
              sc_Slm.Numéro = i
              sc_Slm.G2_Cellule_Paint_Fond()
              sc_Slm.G5_Cellule_Paint_Valeur()
              sc_Slm.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
              U_Slm(i) = False
            End If
          Next i
        End If
        '1 On remet la cellule précédente en état
        Dim sc_Prv As New Cellule_Cls With {.Numéro = Prv_Pbl_Cell_Select}
        sc_Prv.G2_Cellule_Paint_Fond()
        sc_Prv.G5_Cellule_Paint_Valeur()
        sc_Prv.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
        '2 On sélectionne la cellule  
        sc.Cellule_Refresh()
        sc.G7_Cellule_Paint_Select()
        '3 
        Prv_Pbl_Cell_Select = Pbl_Cell_Select
    End Select
    Prv_Plcy_Slm = Plcy_Slm
  End Sub

  Private Sub Frm_SDK_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
    'Le pointeur de la souris est passé sur le composant
    Dim Cntl As Windows.Forms.Form = DirectCast(sender, Windows.Forms.Form)
    Dim MM_Pt As New Point(x:=e.X, y:=e.Y)
    Dim Cellule_MM As Integer      ' Il s'agit de la Cellule où se trouve la souris
    Dim Candidat_MM_Pt As Integer  ' Il s'agit du Candidat dans la Cellule où se trouve la souris
    ' Test des sélections multiples
    With My.Computer.Keyboard
      If Plcy_Gnrl = "Sas" Then
        If .CtrlKeyDown Or .ShiftKeyDown Then Plcy_Slm = True
      End If
      If Not .CtrlKeyDown And Not .ShiftKeyDown Then Plcy_Slm = False
    End With

    Try
      'Se produit après l'effacement d'une valeur
      If Prv_MM_Pt = MM_Pt Then Exit Sub
      Prv_MM_Pt = MM_Pt

      Cellule_MM = Wh_Cellule_Pt_IA(MM_Pt)
      If Cellule_MM = -1 Then Exit Sub
      Candidat_MM_Pt = Wh_Cellule_Candidat_Pt_IA(Cellule_MM, MM_Pt)
      If Candidat_MM_Pt = -1 Then Exit Sub
      'Le traitement est effectué quelque soit la typologie de la cellule
      '0 La cellule sur laquelle passe la souris devient la Pbl_Cell_Select
      Pbl_Cell_Select = Cellule_MM
      'Le 13/08/2025 Je ne sais à quoi sert cette instruction
      'Cursor.Current = DefaultCursor
      Dim Rct_Cdd_Numéro As Integer = (Pbl_Cell_Select * 10) + Candidat_MM_Pt
      Pbl_Cell_Candidat_Select = Candidat_MM_Pt

      ' On doit avoir changé de Rectangle de Cellule

      If Prv_Pbl_Cell_Select < 0 Then Prv_Pbl_Cell_Select = 0
      If Prv_Pbl_Cell_Select < 0 _
      Or Prv_Pbl_Cell_Select > Sqr_Cel.GetUpperBound(0) Then Prv_Rct_Cdd_Numéro = 0


      'Select Case Plcy_Zoom
      '  Case True
      '    If Sqr_Cdd(Prv_Rct_Cdd_Numéro).Contains(MM_Pt) Then Exit Sub
      '  Case False
      If Sqr_Cel(Prv_Pbl_Cell_Select).Contains(MM_Pt) Then Exit Sub
      '      End Select

      Dim sc As New Cellule_Cls With {.Numéro = Cellule_MM}
      Select Case Plcy_Slm
        Case True
          '1 On sélectionne la cellule  
          U_Slm(Pbl_Cell_Select) = True
          sc.Cellule_Refresh()
          sc.G7_Cellule_Paint_Select()
        Case False
          '0 Si slm, on remet les cellules précédentes en état
          If Prv_Plcy_Slm Then
            Dim sc_Slm As New Cellule_Cls
            For i As Integer = 0 To 80
              If U_Slm(i) Then
                sc_Slm.Numéro = i
                sc_Slm.G2_Cellule_Paint_Fond()
                sc_Slm.G5_Cellule_Paint_Valeur()
                sc_Slm.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
                U_Slm(i) = False
              End If
            Next i
          End If
          '1 On remet la cellule précédente en état
          Dim sc_Prv As New Cellule_Cls With {.Numéro = Prv_Pbl_Cell_Select}
          sc_Prv.G2_Cellule_Paint_Fond()
          sc_Prv.G5_Cellule_Paint_Valeur()
          sc_Prv.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()

          '2 On sélectionne la cellule passée sous la souris
          ' Si Plcy_AideGraphique, il faut alors rafraîchir les cellules explicatives, sinon Uniquement la cellule concernée
          If Plcy_AideGraphique Then
            ' les valeurs ont été dessinées pour la grille.
            G4_Grid_Stratégie_All()
          Else
            sc.Cellule_Refresh()
          End If
          sc.G7_Cellule_Paint_Select()
          '3 La cellule sur laquelle passe la souris devient la Prv_Pbl_Cell_Select
          Prv_Pbl_Cell_Select = Pbl_Cell_Select
          Prv_Pbl_Cell_Candidat_Select = Pbl_Cell_Candidat_Select
      End Select

      'Positionnement du Formulaire d'insertion de Candidats
      Prv_Rct_Cdd_Numéro = Rct_Cdd_Numéro
      If Plcy_FIC_Frm_Insérer_Candidats Then Aimantation(Pbl_Cell_Select)
      Prv_Plcy_Slm = Plcy_Slm
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Public Sub Aimantation(Cellule As Integer)
    'Positionnement du Formulaire d'insertion de Candidats
    Frm_Insérer_Candidats.Btn_Insérer.Text = Msg_Read_IA("SDK_10070", {U_cr(Cellule)})
    Dim Pos_L As Integer = Left + Sqr_Cel(Cellule).Left
    Dim Pos_T As Integer = Top + Sqr_Cel(Cellule).Top
    Select Case Plcy_FIC_Zone_Aimantée
      Case "G"
        Pos_L += WH
        Pos_T += 0
        Frm_Insérer_Candidats.SetDesktopLocation(Pos_L, Pos_T)
      Case "BG"
        Pos_L += -CInt(Frm_Insérer_Candidats.Width / 2) + CInt((3 * WH) / 2)
        Pos_T += -Frm_Insérer_Candidats.Height + CInt((5 * WH) / 8)
        Frm_Insérer_Candidats.SetDesktopLocation(Pos_L, Pos_T)
      Case "BD"
        Pos_L += -CInt(Frm_Insérer_Candidats.Width / 2) - WHhalf
        Pos_T += -Frm_Insérer_Candidats.Height + CInt((5 * WH) / 8)
        Frm_Insérer_Candidats.SetDesktopLocation(Pos_L, Pos_T)
      Case "D"
        Pos_L += -Frm_Insérer_Candidats.Width + CInt(WH / 4)
        Pos_T += 0
        Frm_Insérer_Candidats.SetDesktopLocation(Pos_L, Pos_T)
      Case Else
        'Annulation de l'aimantation
    End Select
  End Sub
  Private Sub Frm_SDK_MouseClick(sender As Object, e As MouseEventArgs) Handles MyBase.MouseClick
    ' e as MouseEventArgs permet de localiser la souris
    ' seuls les clics gauche et milieu sont détectés, le clic droit affiche le menu contextuel
    ' Après un clic de souris sur le contrôle
    ' L'utilisateur appuie sur le bouton de la souris
    Dim Pt As New Point(x:=e.X, y:=e.Y)
    Try
      'Il n'est pas besoin de calculer la cellule, c'est Pbl_Cell_Select
      If Pbl_Cell_Select = -1 Then Exit Sub
      Dim Candidat_Pt As Integer = Wh_Cellule_Candidat_Pt_IA(Pbl_Cell_Select, Pt)
      'Le mouse-click se fait soit sur un square, soit ailleurs
      'Il est de 3 types: Gauche, Milieu, Droit
      'Gauche: Rien de particulier, le menu contextuel est préparé,
      'Milieu: Affichage des candidats de l'unité si Plcy_MouseClick_Middle = True 
      'Droit : Le menu contextuel est affiché, ainsi le bouton droit n'est pas traité par Frm_SDK_MouseClick
      Select Case e.Button
        Case MouseButtons.Left
          Frm_SDK_MouseClick_Left(Candidat_Pt)

        Case MouseButtons.Middle
          ' Affiche les candidats de l'unité en fonction de la position du clic et de l'option Afficher les candidats en InfoBulle
          ' Affichage des candidats de la Ligne (Donc colonne 0 ou 8)
          If Plcy_MouseClick_Middle Then Frm_SDK_MouseClick_Middle_Mistral(sender, Candidat_Pt)

        Case MouseButtons.Right
          'le clic droit n'est pas détecté sur les Cellules R et V, car il y a un ContextMenuStrip sur le formulaire
          'le clic droit est détecté sur les Cellules       I, car il n'y a pas de ContextMenuStrip
          'le Frm_SDK_MouseMove précédent lance G7_Cellule_Paint_Select
          'et G7_Cellule_Paint_Select    lance Mnu_Mngt_Nrm ou Mnu_Mngt_Sas
      End Select
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Frm_SDK_MouseClick_Left(Candidat_Pt As Integer)
    ' Provient UNIQUEMENT de Frm_SDK_MouseClick
    'Procédure de Saisie par simple clic gauche
    '             sur un candidat affiché ou sur la grille PCA de saisie
    'Utilisation de Cellule pour simplifier la lecture du code
    Dim Cellule_MCL As Integer = Pbl_Cell_Select
    Dim sc As New Cellule_Cls With {.Numéro = Cellule_MCL}
    Select Case sc.Typologie
      Case "I", "R" 'Il s'agit d'un clic de Sélection, puisque la cellule est remplie
      Case "V"
        'En Mode Sas, on peut saisir n'importe quel candidat
        'En Mode Nrm, sans stratégie, la saisie sera vérifiée avec la solution
        If (Plcy_Gnrl = "Sas" And U(Cellule_MCL, 3) = Cnddts_Blancs) _
        Or (Plcy_Gnrl = "Nrm" And Plcy_Strg = "   ") Then
          Cell_Val_Insert(CStr(Candidat_Pt), Cellule_MCL, "Mse_Clk_PCA")
        End If
        'Comme des candidats sont affichés,
        '      seuls ces derniers sont saisisables
        '      ou s'il n'y en a qu'un, celui-ci est saisi
        '      IL FAUT TOUJOURS CLIQUER sur LE CANDIDAT
        If (Plcy_Gnrl = "Nrm" And Plcy_Strg = "Cdd") _
        Or (Plcy_Gnrl = "Nrm" And Stg_Get(Plcy_Strg).Type = "I") _
        Or (Plcy_Gnrl = "Nrm" And Stg_Get(Plcy_Strg).Type = "E") _
        Or (Plcy_Gnrl = "Sas" And U(Cellule_MCL, 3) <> Cnddts_Blancs) Then
          If U(Cellule_MCL, 3).Contains(CStr(Candidat_Pt)) = True Then
            Cell_Val_Insert(CStr(Candidat_Pt), Cellule_MCL, "Mse_Clk_Cdd")
          End If
        End If
        'Les candidats dans ce cas ne sont pas affichés, mais la grille de saisie
        If (Plcy_Gnrl = "Nrm" And Plcy_Strg <> "   ") _
        Or (Plcy_Gnrl = "Nrm" And Mid$(Plcy_Strg, 1, 2) = "FV") Then
          If U(Cellule_MCL, 3).Contains(CStr(Candidat_Pt)) = True Then
            Cell_Val_Insert(CStr(Candidat_Pt), Cellule_MCL, "Mse_Clk_Cdd")
          End If
        End If
    End Select
    'le Cellule_Refresh permet de ne pas trop griser la cellule
    sc.Cellule_Refresh()
    sc.G7_Cellule_Paint_Select()
  End Sub
  Private Sub Frm_SDK_MouseClick_Middle_Mistral(sender As Object, Candidat As Integer)
    ' Provient UNIQUEMENT de Frm_SDK_Mouse_Click
    Dim TTT_Message As String = Cnddts_Blancs
    Dim Cellule As Integer = Pbl_Cell_Select
    '  True                          
    '-----------+        
    '  1  2  3  |        
    '           |        
    '  4  5  6  |        
    '           |        
    '  7  8  9  |        
    '-----------+        
    ' Pour les régions,  il faut cliquer-milieu dans les centres des régions
    ' Pour les colonnes, il faut cliquer-milieu dans les cellules du haut ou du bas des colonnes 
    '                                           et pour les colonnes de gauche et de droite, il faut cliquer-milieu dans le haut droit  ou haut gauche
    '                                                                                                               ou      bas  droit  ou haut gauche
    '                                           afin de faire la distinction avec les lignes
    ' Pour les lignes, même raisonnement
    ' Pour les cellules des 4 angles, on précise la position pour l'affichage de la ligne ou de la colonne
    If Plcy_Fantasy Then Exit Sub    'Le TTT_Message ne fonctionne pas avec une police fantaisie.

    Select Case Cellule
      Case 1, 2, 3, 4, 5, 6, 7, 73, 74, 75, 76, 77, 78, 79           ' Colonnes haut et bas
        TTT_Message = Wh_Candidats_Unité(U_9CelCol(U_Col(Cellule)))
      Case 9, 18, 27, 36, 45, 54, 63, 17, 26, 35, 44, 53, 62, 71     ' lignes   gauche et droite
        TTT_Message = Wh_Candidats_Unité(U_9CelRow(U_Row(Cellule)))
      Case 10, 13, 16, 37, 40, 43, 64, 67, 70                        ' Régions
        TTT_Message = Wh_Candidats_Unité(U_9CelReg(U_Reg(Cellule)))
      Case 0                                                         ' Coin en haut à gauche
        Select Case Candidat
          Case 2, 3 ' Pour la Colonne : Emplacement 2, 3
            TTT_Message = Wh_Candidats_Unité(U_9CelCol(U_Col(Cellule)))
          Case 4, 7 ' Pour la Ligne   : Emplacement 4, 7
            TTT_Message = Wh_Candidats_Unité(U_9CelRow(U_Row(Cellule)))
          Case Else
        End Select
      Case 8                                                         ' Coin en haut à droite
        Select Case Candidat
          Case 1, 2 ' Pour la Colonne : Emplacement 1, 2
            TTT_Message = Wh_Candidats_Unité(U_9CelCol(U_Col(Cellule)))
          Case 6, 9 ' Pour la Ligne   : Emplacement 6, 9
            TTT_Message = Wh_Candidats_Unité(U_9CelRow(U_Row(Cellule)))
          Case Else
        End Select
      Case 72                                                        ' Coin en bas  à gauche
        Select Case Candidat
          Case 1, 4  ' Pour la Ligne  : Emplacement 1, 4
            TTT_Message = Wh_Candidats_Unité(U_9CelRow(U_Row(Cellule)))
          Case 8, 9  ' Pour la Colonne  : Emplacement 8, 9
            TTT_Message = Wh_Candidats_Unité(U_9CelCol(U_Col(Cellule)))
          Case Else
        End Select
      Case 80                                                        ' Coin en bas  à droite
        Select Case Candidat
          Case 7, 8 ' Pour la Colonne : Emplacement 7, 8
            TTT_Message = Wh_Candidats_Unité(U_9CelCol(U_Col(Cellule)))
          Case 3, 6 ' Pour la Ligne   : Emplacement 3, 6
            TTT_Message = Wh_Candidats_Unité(U_9CelRow(U_Row(Cellule)))
          Case Else
        End Select
      Case Else
    End Select
    Plcy_FIC_TTT = TTT_Message

    Dim TTT_ToolTipText As String
    If TTT_Message <> Cnddts_Blancs Then
      TTT_ToolTipText = TTT_MEF_Cdd(TTT_Message)
    Else
      TTT_ToolTipText = "  M a l" & vbCrLf & "  P l a" & vbCrLf & "  c é ! "
    End If
    Dim TTT_Font As New Font("Consolas", 14, FontStyle.Italic)
    ' Créer une instance de CustomToolTip
    MouseClick_Middle_ToolTip = New CustomToolTip(TTT_ToolTipText, TTT_Font)
    Dim Position As New Point(Left + Get_Centre(Cellule, Candidat).X, Top + Get_Centre(Cellule, Candidat).Y)

    MouseClick_Middle_ToolTip.ShowTooltip(Position)
    TTT_Timer.Interval = 2000 ' 5 secondes
    AddHandler TTT_Timer.Tick, Sub(senderObj As Object, eventArgs As EventArgs)
                                 MouseClick_Middle_ToolTip.HideTooltip()
                                 TTT_Timer.Stop()
                               End Sub
    TTT_Timer.Start()
  End Sub

  Private Sub Frm_SDK_MouseWheel_IA(sender As Object, e As MouseEventArgs) Handles MyBase.MouseWheel
    ' Stratégie des Filtres "FV1" To "FV9" et "FC1" to "FC9"
    ' Détermination du sens de défilement
    Dim ScrollDelta As Integer = e.Delta * SystemInformation.MouseWheelScrollLines \ SystemInformation.MouseWheelScrollDelta
    Dim Sens As Integer = Math.Sign(ScrollDelta)
    If Plcy_Gnrl = "Nrm" AndAlso Plcy_Strg.StartsWith("FV") Then MouseWheel_Valeur(Sens)
    If Plcy_Gnrl = "Nrm" AndAlso Plcy_Strg.StartsWith("FC") Then MouseWheel_Candidat(Sens)
  End Sub
  Public Sub MouseWheel_Valeur(Sens As Integer)
    Dim NewMW As Integer
    If Not Integer.TryParse(Plcy_Strg.Substring(2, 1), NewMW) Then Exit Sub
    ' Compter les occurrences de chaque valeur sur la grille
    Dim Result As Wh_Nb_Cell_Struct = Wh_Nb_Cell(U)
    Dim Val_Nb(9) As Integer
    ' Copie des valeurs de Val_Nb dans Val_Nb
    For i As Integer = 0 To 9
      Val_Nb(i) = Result.Val_Nb(i)
    Next i
    Dim StartVal As Integer = NewMW
    Do
      NewMW = ((NewMW + Sens + 8) Mod 9) + 1
      ' Si la valeur n'est pas présente 9 fois, on la présente
      If Val_Nb(NewMW) < 9 Then Exit Do
    Loop While NewMW <> StartVal

    Strategy_Switch("FV" & CStr(NewMW))
    G4_Grid_Stratégie_Flt()
    Sélection_Pbl_Cell_Standard()
  End Sub
  Public Sub MouseWheel_Candidat(ByVal Sens As Integer)
    Dim FiltreMW As Integer
    If Not Integer.TryParse(Plcy_Strg.Substring(2, 1), FiltreMW) Then Exit Sub
    Dim NewMW As Integer = ((FiltreMW + Sens + 8) Mod 9) + 1
    Strategy_Switch("FC" & CStr(NewMW))
    G4_Grid_Stratégie_Flt()
    Sélection_Pbl_Cell_Standard()
  End Sub
#End Region

#Region "Menu"
  '-------------------------------------------------------------------------------
  '
  ' Menus
  '
  '--------------01---------------------------------------------------------------
  Private Sub Mnu01_Ouvrir_Click(sender As Object, e As EventArgs) Handles Mnu01_Ouvrir.Click
    'Chargement d'une nouvelle partie en mode normal ou Sas
    Frm_LoadParties.Show()
  End Sub
  Private Sub Mnu01_RejouerLaPartie_Click(sender As Object, e As EventArgs) Handles Mnu01_RejouerLaPartie.Click
    Game_New_Game(Plcy_Gnrl, LP_Nom, LP_Prb, LP_Prb, LP_Sol, Cdd729:=StrDup(729, " "), LP_Frc)
  End Sub
  Private Sub Mnu01_Saisir_Click(sender As Object, e As EventArgs) Handles Mnu01_Saisir.Click
    Dim Nom As String = "Pzzl_" & "_" & Format(Now, "yyyy_MM_dd_HH_mm_ss")
    'Le nom d'une partie est limitée à 15 caractères
    Dim Prb As String = StrDup(81, " ")
    Dim Jeu As String = StrDup(81, " ")
    Dim Sol As String = StrDup(81, " ")
    Dim Frc As String = "0"
    Plcy_Saisir_Commencer = True
    'La procédure Saisir diffère de celle de l'Edition
    'U(i,1) est égal à " " tant que Commencer n'est pas lancé.
    'Les candidats collatéraux sont enlevés au fur et à mesure de la Saisie 
    Game_New_Game(Plcy_Gnrl, Nom, Prb, Jeu, Sol, StrDup(729, " "), Frc)
    Mnu01_Saisir.Enabled = False
    Mnu01_Commencer.Enabled = True
  End Sub
  Private Sub Mnu01_Commencer_Click(sender As Object, e As EventArgs) Handles Mnu01_Commencer.Click
    Dim Nom As String = "Pzzl_" & "_" & Format(Now, "yyyy_MM_dd_HH_mm_ss")
    'Le nom d'une partie est limitée à 15 caractères
    Dim Prb As String = ""
    Dim Jeu As String = ""
    Plcy_Saisir_Commencer = False
    'Les valeurs saisies en 2 deviennent valeurs initiales
    For i As Integer = 0 To 80
      Prb &= U(i, 2) : Jeu &= U(i, 2)
    Next i
    Dim Sol As String = StrDup(81, " ")
    Dim Frc As String = "0"
    Game_New_Game(Plcy_Gnrl, Nom, Prb, Jeu, Sol, StrDup(729, " "), Frc)
    Mnu01_Saisir.Enabled = True
    Mnu01_Commencer.Enabled = False
    Jrn_Add(, {Prb})
  End Sub
  Private Sub Mnu01_EnregistrerUnePartieTest_Click(sender As Object, e As EventArgs) Handles Mnu01_EnregistrerUnePartieTest.Click
    Pzzl_Write_Partie_Test()
  End Sub
  Private Sub Mnu01_ChargerUnePartieTest_Click(sender As Object, e As EventArgs) Handles Mnu01_ChargerUnePartieTest.Click
    Pzzl_Load_Partie_Test()
  End Sub
  Private Sub Mnu01_OuvrirLaBibliothèqueTestDeHodoku_Click(sender As Object, e As EventArgs) Handles Mnu01_OuvrirLaBibliothèqueTestDeHodoku.Click
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Procédure_Name_Get()})
    Frm_LoadPartiesHodoku.Show()
  End Sub
  Private Sub Mnu01_OuvrirLeRépertoire_Click(sender As Object, e As EventArgs) Handles Mnu01_OuvrirLeRépertoire.Click
    Dim Pgm As String = "Explorer /e, /n, " & File_SDK
    Nsd_i = Shell(Pgm, AppWinStyle.NormalFocus)
  End Sub
  Private Sub Mnu01_JeuSansAssistance_Click(sender As Object, e As EventArgs) Handles Mnu01_JeuSansAssistance.Click
    Dim J As String = ""
    For i As Integer = 0 To 80 : J &= U(i, 2) : Next i 'L'ensemble des valeurs est le Jeu
    Select Case Plcy_Gnrl
      Case "Nrm" : Plcy_Gnrl = "Sas"
      Case "Sas" : Plcy_Gnrl = "Nrm"
    End Select
    Game_New_Game(Plcy_Gnrl, LP_Nom, LP_Prb, J, LP_Sol, StrDup(729, "0"), LP_Frc)
  End Sub
  Private Sub Mnu01_Quitter_Click(sender As Object, e As EventArgs) Handles Mnu01_Quitter.Click
    Close()
  End Sub
  '--------------02---------------------------------------------------------------
  Private Sub Mnu02_Refaire_Click(sender As Object, e As EventArgs) Handles Mnu02_Refaire.Click
    Select Case Plcy_Gnrl
      Case "Nrm"
        Undo_Redo("Refaire")
      Case Else
    End Select
  End Sub
  Private Sub Mnu02_Annuler_Click(sender As Object, e As EventArgs) Handles Mnu02_Annuler.Click
    Select Case Plcy_Gnrl
      Case "Nrm"
        Undo_Redo("Annuler")
      Case Else
    End Select
  End Sub
  Private Sub Mnu02_Effacer_Click(sender As Object, e As EventArgs) Handles Mnu02_Effacer.Click
    'Effacer remet à blanc une case et replace tous les candidats.
    'Cette option n'est accessible que si la case a été saisie
    'En réalité cette option NE PEUT PAS ETRE UTILISéE par le menu, puisqu'il faut que la cellule soit sélectionnée
    'SEUL Le raccourci SUPPR exécute cette option
    '     ou la touche SUPPR du clavier numérique en mode déplacement
    Try
      Cell_Val_Delete(Pbl_Cell_Select, "Mnu_Eff")
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu02_InsérerLaSolution_Click(sender As Object, e As EventArgs) Handles Mnu02_InsérerLaSolution.Click
    'En réalité cette option NE PEUT PAS ETRE UTILISéE par le menu, puisqu'il faut que la cellule soit sélectionnée
    'SEUL Le raccourci FN + INS exécute cette option
    '     ou la touche INS du clavier numérique en mode déplacement  
    If Plcy_Solution_Existante = False Then Exit Sub
    Try
      Dim Solution As String = U_Sol(Pbl_Cell_Select)
      Cell_Val_Insert(Solution, Pbl_Cell_Select, "Mnu_Ins_Sol")
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu02_Copier_Click(sender As Object, e As EventArgs) Handles Mnu02_Copier.Click
    'Uniquement les Valeurs initiales
    ClipBoard_Copier_New("1")
  End Sub
  Private Sub Mnu02_Copier2_Click(sender As Object, e As EventArgs) Handles Mnu02_Copier2.Click
    'Toutes les Valeurs  
    ClipBoard_Copier_New("2")
  End Sub
  Private Sub Mnu02_Copie3_Click(sender As Object, e As EventArgs) Handles Mnu02_Copie3.Click
    ClipBoard_Copier_New("3")
  End Sub

  Private Sub Mnu02_Coller_Click(sender As Object, e As EventArgs) Handles Mnu02_Coller.Click
    'Coller le Presse-Papier dans la Grille
    ClipBoard_Coller()
  End Sub
  Private Sub Mnu02_CopierlaGrilleEnRTFDansLeJournal_Click(sender As Object, e As EventArgs) Handles Mnu02_CopierlaGrilleDansLeJournal.Click
    ClipBoard_Coller_RTF()
  End Sub
  '--------------03---------------------------------------------------------------
  Private Sub Mnu03_EffacerLeJournal_Click(sender As Object, e As EventArgs) Handles Mnu03_EffacerLeJournal.Click
    Jrn_Clear()
  End Sub
  Private Sub Mnu03_CopierLeJournalEnModeRTB_Click_1(sender As Object, e As EventArgs) Handles Mnu01_CopierLeJournalEnModeRTF.Click
    Dim File_SDK As String = Jrn_RcdRTF()
    Processing_Start_IA(File_SDK)
  End Sub
  Private Sub Mnu03_AfficherLaSolution_Click(sender As Object, e As EventArgs) Handles Mnu03_AfficherLaSolution.Click
    If Plcy_Solution_Existante = False Then Exit Sub
    'Sauvegarde du jeu en cours
    LP_Jeu = ""
    For i As Integer = 0 To 80 : LP_Jeu &= U(i, 2) : Next i
    For i As Integer = 0 To 80
      If U(i, 1) = " " And U_Sol(i) <> " " Then U(i, 2) = U_Sol(i)
    Next i
    'Un seul Invalidate par fonction
    B_Info.Text = "Affichage de la Solution"
    Dim Gril As New Grille_Cls
    Gril.Grille_Refresh()

    Thread.Sleep(2000) 'Le temps de lire quelques valeurs
    For i As Integer = 0 To 80 : U(i, 2) = LP_Jeu.Substring(i, 1) : Next i
    B_Info.Text = " _ "
    ' A cet endroit, il n'est pas nécessaire de faire un B_Info.Grille_Refresh !!!
    Event_OnPaint = "Solution_Affichée"
    Invalidate()
  End Sub
  Private Sub Mnu03_Rafraîchir_Click(sender As Object, e As EventArgs) Handles Mnu03_Rafraîchir.Click
    Dim Gril As New Grille_Cls
    Gril.Grille_Refresh()
    Dim sc As New Cellule_Cls With {.Numéro = Pbl_Cell_Select}
    sc.G7_Cellule_Paint_Select()
  End Sub
  '--------------Transformation---------------------------------------------------
  Private Sub Mnu031I_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Mnu031I.SelectedIndexChanged
    Transf_Incrémentation(CInt(sender.ToString().Substring(0, 1)))
  End Sub
  Private Sub Mnu03MH_Click(sender As Object, e As EventArgs) Handles Mnu03MH.Click
    Transf_Symétrie(1)
  End Sub
  Private Sub Mnu03MV_Click(sender As Object, e As EventArgs) Handles Mnu03MV.Click
    Transf_Symétrie(2)
  End Sub
  Private Sub Mnu03DD_Click(sender As Object, e As EventArgs) Handles Mnu03DD.Click
    Transf_Symétrie(3)
  End Sub
  Private Sub Mnu03DG_Click(sender As Object, e As EventArgs) Handles Mnu03DG.Click
    Transf_Symétrie(4)
  End Sub
  Private Sub Mnu031_Rotation90_Click(sender As Object, e As EventArgs) Handles Mnu031_Rotation90.Click
    Transf_Rotation()
  End Sub
  Private Sub Mnu031_Aléatoire_Click(sender As Object, e As EventArgs) Handles Mnu031_Aléatoire.Click
    Transf_Aléatoire()
  End Sub
  Private Sub Mnu03RH_Click(sender As Object, e As EventArgs) Handles Mnu03RH.Click
    Transf_Région_H()
  End Sub
  Private Sub Mnu03RV_Click(sender As Object, e As EventArgs) Handles Mnu03RV.Click
    Transf_Région_V()
  End Sub
  '--------------04n--------------------------------------------------------------
  Private Sub Mnu04n_Cdd_Click(sender As Object, e As EventArgs) Handles Btn_Cdd.Click
    Strategy_Dsp_Cdd()
  End Sub
  Private Sub Mnu04n_CdU_Click(sender As Object, e As EventArgs) Handles Btn_CdU.Click
    Strategy_Dsp("CdU", AddressOf Sélection_Pbl_Cell_Standard)
  End Sub
  Private Sub Mnu04n_CdO_Click(sender As Object, e As EventArgs) Handles Btn_CdO.Click
    Strategy_Dsp("CdO", AddressOf Sélection_Pbl_Cell_Standard)
  End Sub
  Public Sub Mnu04n_Stratégie_BTXYSJZKQ(Sender As Object, e As EventArgs) Handles Btn_XYZ.Click, Btn_XYw.Click, Btn_Xwg.Click, Btn_Unq.Click, Btn_Tpl.Click, Btn_Swf.Click, Btn_SKy.Click, Btn_Jly.Click, Btn_Cbl.Click
    'Le Sender provient soit du texte de l'option du menu, le texte du menu est explicite
    '                   soit de la barre d'Outils, le texte d'un bouton est une simple lettre
    Dim stgs() As String = {"Cbl", "Tpl", "Xwg", "XYw", "Swf", "Jly", "XYZ", "SKy", "Unq"}

    For Each stg As String In stgs
      If Sender.ToString() = Stg_Get(stg).Texte OrElse Sender.ToString() = Stg_Get(stg).Lettre Then
        Strategy_Dsp(stg, AddressOf Sélection_Pbl_Cell_Standard)
        Exit Sub
      End If
    Next

    ' Si aucune correspondance trouvée
    Jrn_Add(, {Procédure_Name_Get() & " Sender inconnu : " & Sender.ToString()}, "Erreur")
  End Sub
  Public Sub Mnu04n_Stratégie_XW_Click(Sender As Object, e As EventArgs)
    If TypeOf Sender Is ToolStripMenuItem Then
      Dim Mnu As ToolStripMenuItem = CType(Sender, ToolStripMenuItem)
      Dim Mnu_Name As String = Mnu.Name
      Dim Stratégie_XW As String = Mnu_Name.Substring(7, 3)
      Dim U_temp(80, 3) As String
      Array.Copy(U, U_temp, UNbCopy)
      Select Case Stratégie_XW
        Case "Gbl" : Plcy_Strg = "Gbl" : Strategy_Gbl(U_temp)
        Case "Gbv" : Plcy_Strg = "Gbv" : Strategy_Gbv(U_temp)
        Case "GCs" : Plcy_Strg = "GCs" : Strategy_GCs(U_temp)
        Case "XCx" : Plcy_Strg = "XCx" : Strategy_XCx(U_temp)
        Case "XCy" : Plcy_Strg = "XCy" : Strategy_XCy(U_temp)
        Case "XRp" : Plcy_Strg = "XRp" : Strategy_XRp(U_temp)
        Case "XNl" : Plcy_Strg = "XNl" : Strategy_XNl(U_temp)
        Case "WgX" : Plcy_Strg = "WgX" : Strategy_WgX(U_temp)
        Case "WgY" : Plcy_Strg = "WgY" : Strategy_WgY(U_temp)
        Case "WgZ" : Plcy_Strg = "WgZ" : Strategy_WgZ(U_temp)
        Case "WgW" : Plcy_Strg = "WgW" : Strategy_WgW(U_temp)
        Case Else : Jrn_Add(, {"Mnu_Name inconnu : " & Mnu_Name, "Erreur"})
      End Select
    Else
      Jrn_Add(, {"Sender inconnu : " & Sender.ToString(), "Erreur"})
    End If
  End Sub
  Private Sub Mnu04n_AnnulerLaDerniereOption_Click(sender As Object, e As EventArgs) Handles Btn0.Click
    Strategy_Dsp_Standard()
  End Sub
  Private Sub Mnu04n_RésoudreUneCellule_Click(sender As Object, e As EventArgs)
    Cell_Slv_Interactif("S", "Résoudre une Cellule")
  End Sub
  Private Sub Mnu04n_Suggérer_Click(sender As Object, e As EventArgs)
    Cell_Slv_Interactif("S", "Suggérer une Cellule")
  End Sub
  Private Sub Mnu04n_MettreEnÉvidenceLaDernièreCellule_Click(sender As Object, e As EventArgs)
    Strategy_Dsp("DCd", AddressOf Sélection_Pbl_Cell_Standard)
  End Sub
  Private Sub Mnu04n_MettreEnÉvidenceLeCandidatSaisi_Click(sender As Object, e As EventArgs)
    Strategy_Dsp("CdS", AddressOf Sélection_Pbl_Cell_Standard)
  End Sub

  Private Sub Mnu04n_ModeSuggestion_Click(sender As Object, e As EventArgs)
    Strategy_Dsp_Standard()
    'U_Suggest comporte 0/1 pour afficher un disque jaune dans les cellules à documenter
    Swt_Mode_Suggestion *= -1
    For i As Integer = 0 To 80 : U_Suggest(i) = "0" : Next i

    Select Case Swt_Mode_Suggestion
      Case -1 'Mode Suggestion Hors Fonction
        For i As Integer = 0 To 80 : U_Suggest(i) = "0" : Next i
        Dim Gril As New Grille_Cls
        Gril.Grille_Refresh()
      Case +1 'Mode Suggestion En Fonction
        If Plcy_Gnrl = "Nrm" And Plcy_Strg = "   " Then
          'Affiche du coup dans la zone Info l'explication de la suggestion
          Cell_Slv_Interactif("S", "Mode Suggestion")
          ' Un tableau U_Suggest(i) stocke les cellules suggérées
          ' Le mode suggestion propose toutes les cellules d'une stratégie de type Cbl, Tpl à Unq
          ' Il n'est pas prévu d'explications (à placer ou à enlever)
          ' Il est lancé dans Mnu04n_ModeSuggestion_Click
          '              et ensuite après chaque insertion de valeurs ou de candidats
          '              dans Cell_Cdd_Insert
          '                   Cell_Val_Insert

        End If
    End Select
  End Sub
  '--------------04---------------------------------------------------------------

  Private Sub Btn123456789_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn9.MouseDown, Btn8.MouseDown, Btn7.MouseDown, Btn6.MouseDown, Btn5.MouseDown, Btn4.MouseDown, Btn3.MouseDown, Btn2.MouseDown, Btn1.MouseDown
    Select Case e.Button
      Case MouseButtons.Left
        Strategy_Dsp("FV" & sender.ToString(), AddressOf Sélection_Pbl_Cell_Standard)
      Case MouseButtons.Right
        Strategy_Dsp("FC" & sender.ToString(), AddressOf Sélection_Pbl_Cell_Standard)
    End Select
  End Sub
  '--------------05---------------------------------------------------------------
  Private Sub Mnu05_AideSudokuGraphique_Click(sender As Object, e As EventArgs) Handles Mnu05_AideSudokuGraphique.Click
    Dsp_AideGraphique("Alt")
  End Sub
  Private Sub Mnu05_APropos_Click(sender As Object, e As EventArgs) Handles Mnu05_APropos.Click
    Frm_About.Show()
  End Sub
  Private Sub Mnu05_Préférences_Click(sender As Object, e As EventArgs) Handles Mnu05_Préférences.Click
    Frm_Préférences.Show()
  End Sub
  Private Sub Mnu05_FichierDesMessages_Click(sender As Object, e As EventArgs) Handles Mnu05_FichierDesMessages.Click
    Processing_Start_IA(File_SDKMsg)
  End Sub
  Private Sub Mnu05_Documentation_Click(sender As Object, e As EventArgs) Handles Mnu05_Documentation.Click
    'Fichier de Documentation
    Processing_Start_IA(File_SDKDoc)
  End Sub
  Private Sub Mnu05_Maintenance_Click(sender As Object, e As EventArgs) Handles Mnu05_Maintenance.Click
    Nsd_i = Shell($"Notepad {Path_SDK}SuDoKu\SuDoKu\Apriori\aMaintenance.txt", AppWinStyle.MaximizedFocus)
    SendKeys.Send("^{END}")
  End Sub
  Private Sub Mnu05_ModeEtendu_Click(sender As Object, e As EventArgs) Handles Mnu05_ModeEtendu.Click
    'L'option n'est pas visible
    'le raccourci Ctrl+Maj+E
    'permet de passer en mode étendu sans passer par Préférences / Divers / Mode Etendu
    Select Case Plcy_Gbl_Etendue
      Case True : Plcy_Gbl_Etendue = False
      Case False : Plcy_Gbl_Etendue = True
    End Select
    My.Settings.Prf_05D_Plcy_Globale = Plcy_Gbl_Etendue
    OC_Présentation()
    Event_OnPaint = "Frm_Préférences"
    Invalidate()
  End Sub
  Private Sub Mnu05_Dictionnaire_Click(sender As Object, e As EventArgs) Handles Mnu05_Dictionnaire.Click
    Frm_Dictionnaire.Show()
  End Sub
  '--------------06---------------------------------------------------------------
  Private Sub Mnu06_ListerU_Click(sender As Object, e As EventArgs) Handles Mnu06_ListerU.Click
    U_Display()
    'Affichage des valeurs et des candidats
    U_Display2()
    'Infos diverses
    Dim Wh_Nb As Wh_Nb_Cell_Struct = Wh_Nb_Cell(U)
    Wh_Nb_Cell_Display(Wh_Nb)
  End Sub
  Private Sub Mnu06_ListerA_Click(sender As Object, e As EventArgs) Handles Mnu06_ListerA.Click
    Act_Display()
  End Sub
  Private Sub Mnu06_Manuel_des_Stratégies_Click(sender As Object, e As EventArgs) Handles Mnu06_Manuel_des_Stratégies.Click
    'Fichier Manuel_des_Stratégies.docx
    Dim File As String = Path_SDK & "S01_Documentation\Manuel_Complet_SDK.docx"
    Processing_Start_IA(File)
  End Sub
  Private Sub Mnu063_VérificationDeLaGrille_Click(sender As Object, e As EventArgs) Handles Mnu063_VérificationDeLaGrille.Click
    Dim U_Chk(80, 3) As String
    Array.Copy(U, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    U_Checking_Display(U_Check, True)
  End Sub
  Private Sub Mnu06_EffacerLaGrille_Click(sender As Object, e As EventArgs) Handles Mnu06_EffacerLaGrille.Click
    Dim Nom As String = "Grille_Effacée"
    Dim Prb As String = StrDup(81, " ")
    Dim Jeu As String = StrDup(81, " ")
    Dim Sol As String = StrDup(81, " ")
    Dim Frc As String = "0"
    Game_New_Game(Gnrl:="Nrm", Nom:=Nom, Prb:=Prb, Jeu:=Jeu, Sol:=Sol, Cdd729:=StrDup(729, " "), Frc:=Frc)
    Mnu01_Saisir.Enabled = True
    Mnu01_Commencer.Enabled = True
  End Sub
  Private Sub Mnu06_SudokuPCA_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuPCA.Click
    Dim Shell_St As String = Path_SDKAJ & "PCA_Sudoku\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuAngusJohnson_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuAngusJohnson.Click
    Dim Shell_St As String = Path_SDKAJ & "SimpleSudoku\simplesudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuDarrenColes_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuDarrenColes.Click
    Dim Shell_St As String = Path_SDKAJ & "WinSudoku\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuPatriceHenrion_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuPatriceHenrion.Click
    'Ligne de commande /p lance le jeu en petite taille
    Dim Shell_St As String = Path_SDKAJ & "SudokuPH\Sudoku.exe /p"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuNKH_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuNKH.Click
    'L'application NKH a été installé, ainsi qu'un raccourci pour l'appeler
    Dim Shell_St As String = "C:\Users\Public\Desktop\Sudoku1.lnk"
    Processing_Start_IA(Shell_St)
  End Sub
  Private Sub Mnu06_Hodoku220_Click(sender As Object, e As EventArgs) Handles Mnu06_Hodoku220.Click
    ' L'application est lancée et
    ' Il est ajouté une copie de la grille en cours et un collage dans Hodoku 
    ClipBoard_Copier_New("1") ' Copier les valeurs initiales de SDK
    Dim Shell_St As String = "C:\Program Files (x86)\HoDoKu\hodoku.exe"
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
      Jrn_Add(, {ex.Message})
      Jrn_Add(, {ex.ToString()})
    End Try
    Jrn_Add(, {"Attendre 3 secondes pour que l'application soit lancée..."})
    Thread.Sleep(3000)
    AppActivate("HoDoKu - v2.2.0") 'Active une application qui est déjà en cours d'exécution
    ' AppActivate("Java(TM) Platform SE binary")
    SendKeys.Send("^v")
  End Sub

  Private Sub Mnu06_SudokuFedynaK_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuFedynaK.Click
    'Lancement Grille Python + Escape + Entrée
    Dim Grille As String = LP_Prb.Replace(" ", "0")
    Dim Shell_St As String = "C:\Users\berna\AppData\Local\Programs\Python\Python313\pythonw.exe " & Path_SDKAJ & "Python\FedynaK\Sudoku_FedynaK\Sudoku_FedynaK.py " & Grille
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
    Thread.Sleep(500) ' Attendre que la fenêtre soit prête
    ' Récupération du processus
    Dim procList As Process() = Process.GetProcessesByName("pythonw")
    If procList.Length > 0 Then
      Dim hWnd As IntPtr = procList(0).MainWindowHandle
      If hWnd <> IntPtr.Zero Then
        NativeMethods.SetForegroundWindow(hWnd)
        ' Envoi des touches
        SendKeys.Send("{ESC}")
        SendKeys.Send("{ENTER}")
      Else
        MessageBox.Show("Fenêtre inaccessible (MainWindowHandle = 0)")
      End If
    Else
      MessageBox.Show("Processus Python introuvable.")
    End If
  End Sub
  Private Sub Mnu06_SudoCue_Click(sender As Object, e As EventArgs) Handles Mnu06_SudoCue.Click
    Dim Shell_St As String = Path_SDKAJ & "SudoCue\SudoCue.exe"

    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub

  Private Sub Mnu06_DiufSudokuJava_Click(sender As Object, e As EventArgs) Handles Mnu06_DiufSudokuJava.Click
    ClipBoard_Copier_New("1") ' Copier les valeurs initiales de SDK
    Dim Processing As New Process()
    Processing.StartInfo.FileName = "cmd.exe"
    Processing.StartInfo.Arguments = "/c """ & Path_SDKAJ & "DiufSudoku\gui.bat"""
    'gui.bat comporte java -Xms512m -Xmx2G -jar dist\DiufSudoku.jar SudokuExplainer.log
    Processing.StartInfo.WorkingDirectory = Path_SDKAJ & "DiufSudoku"
    Processing.StartInfo.CreateNoWindow = False ' True = cacher la fenêtre noire
    Processing.StartInfo.UseShellExecute = False
    Processing.Start()
  End Sub

  Private Sub Mnu06_MUDancingLink_Click(sender As Object, e As EventArgs) Handles Mnu06_MUDancingLink.Click
    ClipBoard_Copier_New("1") ' Copier les valeurs initiales de SDK

    Dim Shell_St As String = Path_SDKAJ & "Dancing_Link\Dancing_Links_Library_100\Dancing_Links_Library\Dancing_Links_Library\bin\Debug\Dancing_Links_Library.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_CB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Mnu06_CB.SelectedIndexChanged
    'Traitement des URL
    If Not Phase_Démarrage_Terminée Then Exit Sub
    Dim Shell_St As String = Msg_Read_IA("MNU_06IE0")
    Dim Item As String = Mnu06_CB.SelectedItem.ToString()
    Dim Index02_03 As Integer = Mnu06_CB.FindString(Item)
    My.Settings.SDK_IE_Last_Url = Index02_03
    Shell_St &= Item
    Using Processing As New Diagnostics.Process()
      Processing.StartInfo.WindowStyle = ProcessWindowStyle.Maximized
      Processing.StartInfo.FileName = Shell_St
      Processing.Start()
    End Using
  End Sub
  Private Sub Mnu06_ClassicSudoku_Click(sender As Object, e As EventArgs) Handles Mnu06_ClassicSudoku.Click
    Dim Shell_St As String = Path_SDKAJ & "CP_Sudoku\SudokuSolver_src\SudokuSolver src\bin\Debug\SudokuSolver.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuSolver_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuSolver.Click
    Dim Shell_St As String = Path_SDKAJ & "Planete\VB_Net_Sud1916447222005\Sudoku\bin\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_MicrosoftSudoku_Click(sender As Object, e As EventArgs) Handles Mnu06_MicrosoftSudoku.Click
    Dim Shell_St As String = Path_SDKAJ & "Sudoku_VisualBasic\bin\Debug\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_HHSudokuGame_Click(sender As Object, e As EventArgs) Handles Mnu06_HHSudokuGame.Click
    Dim Shell_St As String = Path_SDKAJ & "SudokuGame\SudokuGame\bin\Debug\SudokuPuzzle.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  '--------------07---------------------------------------------------------------
  Private Sub Mnu07_Strategy_CdU_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_CdU.Click
    Mnu07_Strategy("U")
  End Sub
  Private Sub Mnu07_Strategy_CdO_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_CdO.Click
    Mnu07_Strategy("O")
  End Sub
  Private Sub Mnu07_Strategy_DCd_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_DCd.Click
    Mnu07_Strategy("D")
  End Sub
  Private Sub Mnu07_Strategy_Cbl_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_Cbl.Click
    Mnu07_Strategy("B")
  End Sub
  Private Sub Mnu07_Strategy_Tpl_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_Tpl.Click
    Mnu07_Strategy("T")
  End Sub
  Private Sub Mnu07_Strategy_Xwg_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_Xwg.Click
    Mnu07_Strategy("X")
  End Sub
  Private Sub Mnu07_Strategy_XYw_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_XYw.Click
    Mnu07_Strategy("Y")
  End Sub
  Private Sub Mnu07_Strategy_Swf_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_Swf.Click
    Mnu07_Strategy("S")
  End Sub
  Private Sub Mnu07_Strategy_Jly_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_Jly.Click
    Mnu07_Strategy("J")
  End Sub
  Private Sub Mnu07_Strategy_XYZ_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_XYZ.Click
    Mnu07_Strategy("Z")
  End Sub
  Private Sub Mnu07_Strategy_Sky_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_SKy.Click
    Mnu07_Strategy("K")
  End Sub
  Private Sub Mnu07_Strategy_Unq_Click(sender As Object, e As EventArgs) Handles Mnu07_Strategy_Unq.Click
    Mnu07_Strategy("Q")
  End Sub
  Public Sub Mnu07_Strategy(Strategy As String)
    Dim U_temp(80, 3) As String
    Dim Strategy_Rslt(99, 0) As String
    Try
      Array.Copy(U, U_temp, UNbCopy)
      Select Case Strategy
        Case "U" : Strategy_Rslt = Strategy_CdU(U_temp)
        Case "O" : Strategy_Rslt = Strategy_CdO(U_temp)
        Case "D" : Strategy_DCd(U_temp)
        Case "B" : Strategy_Rslt = Strategy_Cbl(U_temp)
        Case "T" : Strategy_Rslt = Strategy_Tpl(U_temp)
        Case "X" : Strategy_Rslt = Strategy_Xwg(U_temp)
        Case "Y" : Strategy_Rslt = Strategy_XYw(U_temp)
        Case "S" : Strategy_Rslt = Strategy_Swf(U_temp)
        Case "J" : Strategy_Rslt = Strategy_Jly(U_temp)
        Case "Z" : Strategy_Rslt = Strategy_XYZ(U_temp)
        Case "K" : Strategy_Rslt = Strategy_SKy(U_temp)
        Case "Q" : Strategy_Rslt = Strategy_Unq(U_temp)
      End Select

      Select Case Strategy
        Case "D"
          DCdd_List_Display()
        Case Else
          Strategy_Rslt_Display_Ligne(Strategy_Rslt, -1)
      End Select

    Catch ex As Exception
      MsgBox(ex.Message,, Procédure_Name_Get())
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try

  End Sub

  Private Sub Mnu07_Gbl_Click(sender As Object, e As EventArgs) Handles Mnu07n_Gbl.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07001")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_Gbv_Click(sender As Object, e As EventArgs) Handles Mnu07n_Gbv.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07002")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_GCs_Click(sender As Object, e As EventArgs) Handles Mnu07n_GCs.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07010")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_XCx_Click(sender As Object, e As EventArgs) Handles Mnu07n_XCx.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07020")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_XCy_Click(sender As Object, e As EventArgs) Handles Mnu07n_XCy.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07030")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_XRp_Click(sender As Object, e As EventArgs) Handles Mnu07n_XRp.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07040")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_XNl_Click(sender As Object, e As EventArgs) Handles Mnu07n_XNl.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07050")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_WgX_Click(sender As Object, e As EventArgs) Handles Mnu07n_WgX.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07110")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_WgY_Click(sender As Object, e As EventArgs) Handles Mnu07n_WgY.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07120")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_WgZ_Click(sender As Object, e As EventArgs) Handles Mnu07n_WgZ.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07130")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  Private Sub Mnu07_WgW_Click(sender As Object, e As EventArgs) Handles Mnu07n_WgW.Click
    Dim File_Name As String = Path_SDK & Msg_Read_IA("MNU_07140")
    If Plcy_Open_Display Then Processing_Start_IA(File_Name)
    Pzzl_Open(File_Name)
    Mnu04n_Stratégie_XW_Click(sender, e)
  End Sub
  '--------------08---------------------------------------------------------------
  Private Sub Mnu08_Jouer_Click(sender As Object, e As EventArgs) Handles Mnu08_Jouer.Click
    Mnu08J_F.Text = "Facile" & " (" & CStr(File_Nb_IA("SDK_F")) & ") "
    Mnu08J_M.Text = "Moyen" & " (" & CStr(File_Nb_IA("SDK_M")) & ") "
    Mnu08J_D.Text = "Difficile" & " (" & CStr(File_Nb_IA("SDK_D")) & ") "
    Mnu08J_E.Text = "Expert" & " (" & CStr(File_Nb_IA("SDK_E")) & ") "
  End Sub
  Private Sub Mnu08J_F_Click(sender As Object, e As EventArgs) Handles Mnu08J_F.Click
    Mnu08J_Click("F")
  End Sub
  Private Sub Mnu08J_M_Click(sender As Object, e As EventArgs) Handles Mnu08J_M.Click
    Mnu08J_Click("M")
  End Sub
  Private Sub Mnu08J_D_Click(sender As Object, e As EventArgs) Handles Mnu08J_D.Click
    Mnu08J_Click("D")
  End Sub
  Private Sub Mnu08J_E_Click(sender As Object, e As EventArgs) Handles Mnu08J_E.Click
    Mnu08J_Click("E")
  End Sub
  Private Sub Mnu08J_Click(Difficulté As String)
    Jrn_Add(, {Procédure_Name_Get()})
    'Extension / Jouer un Sudoku FMDE
    Dim s, l As Integer
    ' l'absence de Order By File Ascending/Descending renvoie les fichiers dans un ordre non garanti,
    ' dépendant du système de fichiers. Donc: désordre assuré.
    'Dim Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(Path_Batch)
    '                                      Where File.Contains("SDK_" & Difficulté)
    '                                      Order By File Ascending
    ' Classique "shuffle LINQ" (mélange)
    Dim rnd As New Random()
    Dim Files As IEnumerable(Of String) =
    IO.Directory.GetFiles(Path_Batch).
        Where(Function(f) f.Contains("SDK_" & Difficulté)).
        OrderBy(Function(f) rnd.Next())
    If Files.Count > 0 Then
      For Each File As String In Files
        'Ouverture et lecture du premier fichier
        Pzzl_Open(File)
        s = InStrRev(File.ToString(), "\")
        l = File.ToString().Length
        Dim Nom_Physique As String = Mid$(File.ToString(), s + 1, l - s)         ' Extension comprise
        'Copie du fichier
        '05/06/2024 Une fois la partie lancée, elle est copiée dans Batch_Poubelle et pourra ainsi être contrôlée
        Dim Source As String = File
        Dim Destination As String = Path_Batch_Poubelle & Nom_Physique
        My.Computer.FileSystem.CopyFile(Source, Destination)
        ' et ensuite Suppression du fichier
        ' Une fois la partie lancée, le fichier est supprimé
        My.Computer.FileSystem.DeleteFile(File,
           Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
           Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently)
        'Affichage du fichier si Plcy_Open_Display
        'L'affichage permet de lire les commentaires et d'en ajouter  
        If Plcy_Open_Display Then Processing_Start_IA(Destination)
        Exit For
      Next File
    Else
      Dim MsgTit As String = Procédure_Name_Get() & " " & Difficulté & " " & Application.ProductName & " " & SDK_Version
      Nsd_i = MsgBox("Il n'a pas été trouvé de parties à jouer.",, MsgTit)
    End If
  End Sub
  Private Sub Mnu08_JouerAutrement_Click(sender As Object, e As EventArgs) Handles Mnu08_JouerAutrement.Click
    Jrn_Add(, {Procédure_Name_Get()})
    Pzzl_Prd_DL()
  End Sub
  Private Sub Mnu08_Création_Click(sender As Object, e As EventArgs) Handles Mnu08_Création.Click
    Jrn_Add(, {Procédure_Name_Get()})
    Pzzl_Prd_Interactif("P")
  End Sub
  Private Sub Mnu08_Résolution_Click(sender As Object, e As EventArgs) Handles Mnu08_Résolution.Click
    Jrn_Add(, {Procédure_Name_Get()})
    Dim Stat As String() = {"#", "#", "#", "#", "#"} '  VI, Val av, Val ap, Cdd av, Cdd ap
    Stat(0) = CStr(Wh_Grid_Nb_Cellules_Initiales(U))
    Stat(1) = CStr(Wh_Grid_Nb_Cellules_Remplies(U))
    Stat(3) = CStr(Wh_Grid_Nb_Candidats(U))

    Pzzl_Slv_Interactif("S")

    Stat(2) = CStr(Wh_Grid_Nb_Cellules_Remplies(U))
    Stat(4) = CStr(Wh_Grid_Nb_Candidats(U))
    Jrn_Add("PRD_30040", {Stat(0), Stat(1), Stat(2), Stat(3), Stat(4)})

    Strategy_Dsp_Standard()
    'Vérification
    Dim U_Chk(80, 3) As String
    Array.Copy(U, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    U_Checking_Display(U_Check, True)
  End Sub
  Private Sub Mnu08_RésoudreEnForceBrute_Click(sender As Object, e As EventArgs) Handles Mnu08_RésoudreEnForceBrute.Click
    Jrn_Add(, {Procédure_Name_Get()})
    'La stratégie en force brute ne fonctionne que si le grille est correcte
    Dim U_Chk(80, 3) As String
    Array.Copy(U, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    U_Checking_Display(U_Check, False)

    If U_Check.Check = True Then
      Cursor.Current = Cursors.WaitCursor
      Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
      Strategy_Force_Brute()
      Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
      Dim Durée As Integer = Durée_Fin - Durée_Déb
      Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
      Jrn_Add(, {"Durée: " & CStr(Durée).PadLeft(5) & " ms, soit " & String.Format("{0:00}:{1:00}:{2:00}:{3:000}", Ts.Hours, Ts.Minutes, Ts.Seconds, Ts.Milliseconds)})
      Jrn_Add(, {"Durée U_Checking non comprise. "})
      Cursor.Current = Cursors.Default
    End If
  End Sub
  Private Sub Mnu08_RésoudreDancingLink_Click(sender As Object, e As EventArgs) Handles Mnu08_RésoudreDancingLink.Click
    Jrn_Add(, {Procédure_Name_Get()})
    'La stratégie DancingLink ne fonctionne que si le grille est correcte
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Cursor.Current = Cursors.WaitCursor
    Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve_IA(U)
    Jrn_Add(, {"Dancing Link        : " & CStr(DL.Nb_Solution)})
    Select Case DL.Nb_Solution
      Case -1 : Jrn_Add(, {"U_temp(,) est contrôlé Not U_Check.Check."})
      Case 0 : Jrn_Add(, {"U_temp(,) n'a pas de solution. "})
      Case 1
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode})
        For i As Integer = 0 To 80
          U(i, 2) = DL.Solution(0).Substring(i, 1)
        Next i
        Dim Gril As New Grille_Cls
        Gril.Grille_Refresh()
      Case Else
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode & " Solutions multiples."})
        For i As Integer = 1 To DL.Solution.Length - 1
          Jrn_Add(, {CStr(i).PadLeft(2) & " " & DL.Solution(i)})
        Next i
    End Select

    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée: " & CStr(Durée).PadLeft(5) & " ms, soit " & String.Format("{0:00}:{1:00}:{2:00}:{3:000}", Ts.Hours, Ts.Minutes, Ts.Seconds, Ts.Milliseconds)})
    Jrn_Add(, {"Durée U_Checking comprise. "})
    Cursor.Current = Cursors.Default
  End Sub
  Private Sub Mnu08_EditionDuProblème_Click(sender As Object, e As EventArgs) Handles Mnu08_EditionDuProblème.Click
    Jrn_Add(, {Procédure_Name_Get()})
    'Public Swt_ModeEdition As Integer = -1 Position Initiale
    Swt_ModeEdition *= -1
    Select Case Swt_ModeEdition
      Case 1
        Plcy_Gnrl = "Edi"
        Mnu08_EditionDuProblème.Checked = True
        Jrn_Add("SDK_00321")
        ContextMenuStrip = Mnu_EDI
        B_Info.Text = Msg_Read_IA("SDK_00320")
        B_Info.BackColor = Color.Gray
        Strategy_Dsp_Cdd()
      Case -1
        Plcy_Gnrl = "Nrm"
        Mnu08_EditionDuProblème.Checked = False
        Jrn_Add("SDK_00322")
        ContextMenuStrip = Mnu_Cel
        B_Info.Text = ""
        B_Info.BackColor = Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Jrn_Add(, {"Nouvelle partie éditée"})
    End Select
  End Sub

  Private Sub Mnu08_DessinerSurLaGrille_Click(sender As Object, e As EventArgs) Handles Mnu08_DessinerSurLaGrille.Click
    Jrn_Add(, {Procédure_Name_Get()})
    'Public Swt_Mode_Dessin As Integer = -1 Position Initiale
    Swt_Mode_Dessin *= -1
    Select Case Swt_Mode_Dessin
      Case 1
        Plcy_Gnrl = "Nrm" : Plcy_Strg = "Obj"
        Mnu08_DessinerSurLaGrille.Checked = True
        Jrn_Add("SDK_00331")
        ContextMenuStrip = Mnu_Obj

        Obj_Symbol = My.Settings.Obj_Symbol
        Obj_Forme = My.Settings.Obj_Forme
        InitMenuItems()
        Couleurs_ResetAndColorizsation(C1Items, Color_List)
        Couleurs_Ckeck(C1Items, Obj_Symbol)
        Objets_Reset(O1Items)
        Objets_CheckAndColor(O1Items, Obj_Forme, Color_BySymbol(Obj_Symbol))
        B_Info.Text = Msg_Read_IA("SDK_00330")
        B_Info.BackColor = Color_BySymbol(Obj_Symbol)
      Case -1
        Plcy_Gnrl = "Nrm" : Plcy_Strg = "   "
        Mnu08_DessinerSurLaGrille.Checked = False
        Jrn_Add("SDK_00332")
        ContextMenuStrip = Mnu_Cel
        B_Info.Text = ""
        B_Info.BackColor = Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Jrn_Add(, {"Nouvelle partie dessinée"})
    End Select

  End Sub

  Private Sub Mnu08_InsérerTouteLaSolution_Click(sender As Object, e As EventArgs) Handles Mnu08_InsérerTouteLaSolution.Click
    Jrn_Add(, {Procédure_Name_Get()})
    If Plcy_Solution_Existante = False Then Exit Sub
    For i As Integer = 0 To 80
      If U(i, 2) = " " Then U(i, 2) = U_Sol(i)
    Next i
    Dim Gril As New Grille_Cls
    Gril.Grille_Refresh()
  End Sub

#End Region

#Region "Menus Contextuels"
  '-------------------------------------------------------------------------------
  '
  ' Menus Contextuels Cellule
  '
  ' Ces options différent des options du menu contextuel d'édition, 
  '     car elles gèrent les valeurs ou les candidats ainsi que les tableau U et U_CddExc
  ' Mnu_Cel_Val_Insérer
  ' Mnu_Cel_Val_Effacer
  ' Mnu_Cel_Cdd_Insérer
  ' Mnu_Cel_Cdd_InsérerN
  ' Mnu_Cel_Cdd_Exclure
  ' Mnu_Cel_Cdd_ExclureAll
  ' Mnu_Cel_Color
  '
  ' Mnu_Cel_Val_Ins_1 à 9          Insérer la valeur 1 à 9 
  ' Mnu_Cel_Cdd_Exc_Separator      
  ' Mnu_Cel_Cdd_Exc_1 à 9          Exclure le candidat 1 à 9 
  ' Mnu_Cel_Cdd_Exd_N              Exclure les candidats ...
  ' Mnu_Cel_Cdd_Exe_A              Exclure tous les candidats
  ' Mnu_Cel_Cdd_Ins_Separator      
  ' Mnu_Cel_Cdd_Ins_1 à 9          Insérer le candidat 1 à 9 
  ' Mnu_Cel_Cdd_Int_N              Insérer les candidats ...
  ' Mnu_Cel_Val_Eff_Separator      
  ' Mnu_Cel_Val_Eff_x              Effacer la valeur saisie
  '
  '-------------------------------------------------------------------------------
  Private Sub Mnu_Cel_Val_Insérer(sender As Object, e As EventArgs) Handles Mnu_Cel_Val_Ins_9.Click, Mnu_Cel_Val_Ins_8.Click, Mnu_Cel_Val_Ins_7.Click, Mnu_Cel_Val_Ins_6.Click, Mnu_Cel_Val_Ins_5.Click, Mnu_Cel_Val_Ins_4.Click, Mnu_Cel_Val_Ins_3.Click, Mnu_Cel_Val_Ins_2.Click, Mnu_Cel_Val_Ins_1.Click
    'Insérer la Valeur, x se trouve en sender.ToString().Substring(18, 1)
    Try
      Dim Cellule_Ins As Integer = Pbl_Cell_Select
      Dim VP As String = sender.ToString().Substring(18, 1)
      Cell_Val_Insert(VP, Cellule_Ins, "Mnu_Ctx_Ins")
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu_Cel_Val_Effacer(sender As Object, e As EventArgs) Handles Mnu_Cel_Val_Eff_x.Click
    Try
      Dim Cellule_Eff As Integer = Pbl_Cell_Select
      Cell_Val_Delete(Cellule_Eff, "Mnu_Ctx_Eff")
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu_Cel_Cdd_Insérer(Sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Ins_9.Click, Mnu_Cel_Cdd_Ins_8.Click, Mnu_Cel_Cdd_Ins_7.Click, Mnu_Cel_Cdd_Ins_6.Click, Mnu_Cel_Cdd_Ins_5.Click, Mnu_Cel_Cdd_Ins_4.Click, Mnu_Cel_Cdd_Ins_3.Click, Mnu_Cel_Cdd_Ins_2.Click, Mnu_Cel_Cdd_Ins_1.Click
    'Insérer le candidat x se trouve en sender.ToString().Substring(20, 1) dans une cellule
    Try
      Dim Cellule_Ins As Integer = Pbl_Cell_Select
      Dim Candidat As String = Sender.ToString().Substring(20, 1)
      Cell_Cdd_Insert(Candidat, Cellule_Ins, "Mnu_Ctx")
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu_Cel_Cdd_InsérerN(Sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Int_N.Click
    Frm_Insérer_Candidats.Show()
  End Sub
  Private Sub Mnu_Cel_Cdd_Exclure(sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Exc_9.Click, Mnu_Cel_Cdd_Exc_8.Click, Mnu_Cel_Cdd_Exc_7.Click, Mnu_Cel_Cdd_Exc_6.Click, Mnu_Cel_Cdd_Exc_5.Click, Mnu_Cel_Cdd_Exc_4.Click, Mnu_Cel_Cdd_Exc_3.Click, Mnu_Cel_Cdd_Exc_2.Click, Mnu_Cel_Cdd_Exc_1.Click
    'Exclure le candidat, x se trouve en sender.ToString().Substring(20, 1) dans une cellule
    Try
      Dim Cellule_Cdd_Exc As Integer = Pbl_Cell_Select
      Dim Candidat As String = sender.ToString().Substring(20, 1)
      Cell_Cdd_Exclude(Candidat, Cellule_Cdd_Exc)
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu_Cel_Cdd_ExclureAll(sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Exe_A.Click
    'Exclusion de tous les candidats pour une cellule
    Try
      Dim Cellule_Cdd_Exc As Integer = Pbl_Cell_Select
      Dim Candidats As String = U(Cellule_Cdd_Exc, 3)
      For j As Integer = 0 To 8
        Dim Cdd As String = Candidats.Substring(j, 1)
        If Cdd <> " " Then
          Cell_Cdd_Exclude(Cdd, Cellule_Cdd_Exc)
        End If
      Next j
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Mnu_Cel_Cdd_ExclureN(sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Exd_N.Click
    ' Exclure n candidats avec InputBox dans une cellule
    ' La valeur par défaut comporte l'ensemble des candidats à exclure stockés dans U_Strg_Cdd_Exc
    Try
      Dim Cellule_Exc As Integer = Pbl_Cell_Select
      Dim Titre As String = "Exclure un ou plusieurs candidats"
      Dim Texte As String = "Candidat(s) souhaité(s) en " & U_cr(Cellule_Exc) & " :" & vbCrLf
      Texte &= "La chaîne de caractères ne doit pas dépasser 9."
      Texte &= "Les candidats peuvent être présentés dans le désordre, et comporter des blancs ."
      Dim Dftvalue As String = U_Strg_Cdd_Exc(Cellule_Exc)
      ' Permet de positionner InputBox près du clic de la souris
      Dim Position As New Point(Me.Left + Get_Centre(Pbl_Cell_Select, 1).X + WH, Me.Top + Get_Centre(Pbl_Cell_Select, 1).Y + WH)
      Dim Candidats As String = InputBox(Texte, Titre, Dftvalue, Position.X, Position.Y)

      If Candidats Is "" Then Exit Sub 'Cancel enfoncé
      Dim l As Integer = Candidats.Length
      If l > 9 Then l = 9
      For j As Integer = 0 To l - 1
        Dim Cdd As String = Candidats.Substring(j, 1)
        If Cdd >= "1" And Cdd <= "9" Then
          Cell_Cdd_Exclude(Cdd, Cellule_Exc)
        End If
      Next j
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  '-------------------------------------------------------------------------------
  '
  ' Menu Contextuel Objet
  '
  '-------------------------------------------------------------------------------

  Private Sub Mnu_Obj_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles Mnu_Obj.Opening
    ' Mémorisation du point de clique à l'ouverture du menu
    Pbl_PtF = Me.PointToClient(MousePosition)
  End Sub

  Private Sub Mnu_Obj_Click(sender As Object, e As EventArgs) Handles Origine.Click, Lister.Click, Flèche_Supprimer.Click, Flèche.Click, Enlever_Tout.Click, Enlever.Click, Disque.Click, Destination.Click, D.Click, Croix.Click, Cercle.Click, Cel_Cdd.Click, Carré.Click, Cadre.Click, C.Click, B.Click, A.Click
    Dim ClickedMenuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
    Dim ToRefresh As Boolean

    Select Case ClickedMenuItem.Name

      Case "A", "B", "C", "D"
        Obj_Symbol = ClickedMenuItem.Name
        Obj_Color = Color_BySymbol(Obj_Symbol)
        Couleurs_ResetAndColorizsation(C1Items, Color_List)
        Couleurs_Ckeck(C1Items, Obj_Symbol)
        ToRefresh = False

      Case "Cadre", "Carré", "Cercle", "Disque", "Croix"
        Obj_Forme = ClickedMenuItem.Name
        Objets_Reset(O1Items)
        Objets_CheckAndColor(O1Items, Obj_Forme, Color_BySymbol(Obj_Symbol))
        ToRefresh = False

      Case "Cel_Cdd"
        Dim Cellule As Integer = Pbl_Cell_Select
        Dim Candidat As Integer = Pbl_Cell_Candidat_Select
        Dim Cdd As Integer = 0
        ' Pour une cellule  Cellule est compris entre 0 et 80 et Cdd = 0
        ' Pour une candidat Cellule est compris entre 0 et 80 et Cdd est compris entre 1 et 9
        If U(Cellule, 3).Contains(CStr(Candidat)) Then Cdd = Candidat
        Objet_List.Add(New Objet_Cls With {.Symbol = Obj_Symbol, .Forme = Obj_Forme, .Cel_From = Cellule, .Cdd_From = Cdd, .Cel_To = -1, .Cdd_To = 0})
        ToRefresh = True

      Case "Flèche"
        ' Pour dessiner une flèche, il faut 2 candidats
        '                           1 candidat Origine et un candidat Destination
        ' Lorsque la flèche est dessinée, le candidat destination devient le candidat origine
        Obj_Forme = sender.ToString()
        Objets_Reset(O1Items)
        Objets_CheckAndColor(O1Items, Obj_Forme, Color_BySymbol(Obj_Symbol))
        ToRefresh = False

      Case "Origine"
        Flè_Cel_From = Pbl_Cell_Select
        Flè_Cdd_From = Pbl_Cell_Candidat_Select
        Flè_From = 0
        If U(Flè_Cel_From, 3).Contains(CStr(Flè_Cdd_From)) Then Flè_From = Flè_Cdd_From
        ToRefresh = False

      Case "Destination"
        Flè_Cel_To = Pbl_Cell_Select
        Flè_Cdd_To = Pbl_Cell_Candidat_Select
        Flè_To = 0
        If U(Flè_Cel_To, 3).Contains(CStr(Flè_Cdd_To)) Then Flè_To = Flè_Cdd_To
        ' Il faut 2 cellules différentes et 2 candidats
        If Flè_Cel_From <> Flè_Cel_To And Flè_From <> 0 And Flè_To <> 0 Then
          Dim Pts As Points_Struct = Get_Pt_From_To_Flèche(Flè_Cel_From, Flè_Cdd_From, Flè_Cel_To, Flè_Cdd_To)
          Objet_List.Add(New Objet_Cls With {.Symbol = Obj_Symbol, .Forme = Obj_Forme, .Cel_From = Flè_Cel_From, .Cdd_From = Flè_From, .Cel_To = Flè_Cel_To, .Cdd_To = Flè_To, .Point_From = Pts.Pt_From, .Point_To = Pts.Pt_To})
        End If
        ' Pour tracer plus rapidement une seconde flèche
        Flè_Cel_From = Flè_Cel_To
        Flè_Cdd_From = Flè_Cdd_To
        Flè_From = Flè_To
        ToRefresh = True

      Case "Flèche_Supprimer"
        ' Pour supprimer la flèche, il faut connaître les 2 points de la flèche.
        'Candidat_Pt = Me.PointToClient(MousePosition) a été documenté à l'ouverture du menu contextuel
        '👉 Mnu_Obj_Opening
        ' Trouver l'objet correspondant
        Dim ligneàSupprimer As Objet_Cls = Objet_List.FirstOrDefault(Function(l) Get_Distance_Point_Flèche_IA(Pbl_PtF, l.Point_From, l.Point_To) < 5)
        If ligneàSupprimer IsNot Nothing Then
          Objet_List.Remove(ligneàSupprimer) ' Supprimer l'objet trouvé
        End If
        ToRefresh = True

      Case "Enlever"
        Dim Cellule As Integer = Pbl_Cell_Select
        Dim Candidat As Integer = Pbl_Cell_Candidat_Select
        Dim Cdd As Integer = 0
        ' Pour une cellule  Cellule est compris entre 0 et 80 et Cdd = 0
        ' Pour une candidat Cellule est compris entre 0 et 80 et Cdd est compris entre 1 et 9
        If U(Cellule, 3).Contains(CStr(Candidat)) Then Cdd = Candidat
        Objet_List.RemoveAll(Function(Obj) Obj.Cel_From = Cellule And Obj.Cdd_From = Cdd And Obj.Cel_To = -1 And Obj.Cdd_To = 0)
        ToRefresh = True

      Case "Enlever_Tout"
        Objet_List.Clear()                   ' Vider Objet_List
        ToRefresh = True

      Case "Lister"
        Objet_List_Display()

      Case Else
        Jrn_Add(, {ClickedMenuItem.Name & " " & sender.ToString() & " en cours!"})
    End Select

    My.Settings.Obj_Symbol = Obj_Symbol
    My.Settings.Obj_Forme = Obj_Forme
    My.Settings.Save()

    If ToRefresh Then
      Dim Gril As New Grille_Cls
      Gril.Grille_Refresh()
    End If
  End Sub
  '-------------------------------------------------------------------------------
  '
  ' Menus Contextuels Edition Val, Cdd, Eff, Ini, Nrm 2649
  ' 
  '-------------------------------------------------------------------------------
  Private Sub Mnu_EDI_Saisir_Valeur_Click(sender As Object, e As EventArgs) Handles Mnu_EDI_Saisir_Valeur.Click
    M02_Menu_Management.Mnu_EDI("Val", sender, e)
  End Sub
  Private Sub Mnu_EDI_Candidats_Click(sender As Object, e As EventArgs) Handles Mnu_EDI_Candidats.Click
    M02_Menu_Management.Mnu_EDI("Cdd", sender, e)
  End Sub
  Private Sub Mnu_EDI_Effacer_Click(sender As Object, e As EventArgs) Handles Mnu_EDI_Effacer.Click
    M02_Menu_Management.Mnu_EDI("Eff", sender, e)
  End Sub
  Private Sub Mnu_EDI_Val_Initiale_Click(sender As Object, e As EventArgs) Handles Mnu_EDI_Val_Initiale.Click
    M02_Menu_Management.Mnu_EDI("Ini", sender, e)
  End Sub
  Private Sub Mnu_EDI_Val_Normale_Click(sender As Object, e As EventArgs) Handles Mnu_EDI_Val_Normale.Click
    M02_Menu_Management.Mnu_EDI("Nrm", sender, e)
  End Sub

  '-------------------------------------------------------------------------------
  '
  ' Procédures Journal
  '
  '-------------------------------------------------------------------------------
  Private Sub Journal_Linkcliked(sender As Object, e As Windows.Forms.LinkClickedEventArgs)
    'Permet d'ouvrir une URL
    Try
      Nsd_P = Diagnostics.Process.Start(e.LinkText)
    Catch ex As Exception
      MsgBox("Url inconnue.")
    End Try
  End Sub
  Private Sub Journal_Click(sender As Object, e As EventArgs)
    'Permet de déterminer l'endroit de fixation du journal
    Journal_Emp_Blocage = Journal.SelectionStart
  End Sub
  '-------------------------------------------------------------------------------
  '
  ' Menus Contextuels Journal
  '
  '-------------------------------------------------------------------------------
  Private Sub Mnu_Jrn_BloquerLeDéroulement_Click(sender As Object, e As EventArgs) Handles Mnu_Jrn_BloquerLeDéroulement.Click
    Swt_DéroulerJournal *= -1
    Select Case Swt_DéroulerJournal
      Case +1
        Mnu_Jrn_BloquerLeDéroulement.Checked = False
        Jrn_Add("SDK_00241") 'SDK_00241=Le déroulement du journal est débloqué.
      Case -1
        Mnu_Jrn_BloquerLeDéroulement.Checked = True
        Jrn_Add("SDK_00240") 'SDK_00240=Le déroulement du journal est bloqué.
    End Select
  End Sub
  Private Sub Mnu_Jrn_CopierLaSélectionDansLePresse_Papier_RTF_Click(sender As Object, e As EventArgs) Handles Mnu_Jrn_CopierLaSélectionDansLePresse_Papier_RTF.Click
    'La partie sélectionnée du journal est copiée en mode RTF dans le Presse-Papier
    'Word permet de récupérer le PP
    Clipboard.Clear()
    Clipboard.SetData(DataFormats.Rtf, CType(Journal.SelectedRtf, Object))
    '.SelectedText permet d'obtenir le texte sélectionné
    '.SelectedRtf                   le texte sélectionné en mode RTF
  End Sub
  Private Sub Mnu_Jrn_CopierLaSélectionDansLePresse_Papier_Texte_Click(sender As Object, e As EventArgs) Handles Mnu_Jrn_CopierLaSélectionDansLePresse_Papier_Texte.Click
    'La partie sélectionnée du journal est copiée en mode Texte dans le Presse-Papier
    'NotePad permet de récupérer le PP ou l'option Coller Angus Johnson Mots Croisés
    Clipboard.Clear()
    Clipboard.SetData(DataFormats.Text, CType(Journal.SelectedText, Object))
  End Sub
#End Region

#Region "Menus de Tests"
  '-------------------------------------------------------------------------------
  ' Menus de Test A et B
  '       Options placées en fin de module pour simplification de maintenance
  '-------------------------------------------------------------------------------
  Private Sub Mnu08_TestA_Click(sender As Object, e As EventArgs) Handles Mnu08_TestA.Click
    TestA()
  End Sub
  Private Sub Mnu08_TestB_Click(sender As Object, e As EventArgs) Handles Mnu08_TestB.Click
    TestB()
  End Sub
  Private Sub Mnu08_TestC_Click(sender As Object, e As EventArgs) Handles Mnu08_TestC.Click
    TestC()
  End Sub
  Private Sub Mnu08_TestD_Click(sender As Object, e As EventArgs) Handles Mnu08_TestD.Click
    TestD()
  End Sub
  Private Sub Mnu08_TestE_Click(sender As Object, e As EventArgs) Handles Mnu08_TestE.Click
    TestE()
  End Sub
  Private Sub Mnu08_TestF_Click(sender As Object, e As EventArgs) Handles Mnu08_TestF.Click
    TestF()
  End Sub
  Private Sub Mnu08_TestG_Click(sender As Object, e As EventArgs) Handles Mnu08_TestG.Click
    TestG()
  End Sub
  Private Sub Mnu08_TestH_Click(sender As Object, e As EventArgs) Handles Mnu08_TestH.Click
    TestH()
  End Sub
  Private Sub Mnu08_TestI_Click(sender As Object, e As EventArgs) Handles Mnu08_TestI.Click
    TestI()
  End Sub
  Private Sub Mnu08_TestJ_Click(sender As Object, e As EventArgs) Handles Mnu08_TestJ.Click
    TestJ()
  End Sub
#End Region

#Region "Menu Graphe"
  Private Sub Mnu0901_Click(sender As Object, e As EventArgs) Handles Mnu0901.Click
    Jrn_Add(, {Procédure_Name_Get() & " Lancement des Stratégies G "})
    Strategy_Switch("   ")
    B_Info.Text = Procédure_Name_Get()
    Dsp_AideGraphique("Non")
    U_Strg_Effacer()
    Dim Gril As New Grille_Cls
    Gril.Grille_Refresh()

    Dim U_temp(80, 3) As String

    For i As Integer = 0 To Stg_List_Link.Count - 1
      Plcy_Strg = Stg_List_Link(i)
      Jrn_Add(, {"Strategy_" & Plcy_Strg})
      Array.Copy(U, U_temp, UNbCopy)
      Select Case Plcy_Strg
        Case "Gbl" : Strategy_Gbl(U_temp)
        Case "Gbv" : Strategy_Gbv(U_temp)
        Case "GCs" : Strategy_GCs(U_temp)
        Case "XCx" : Strategy_XCx(U_temp)
        Case "XCy" : Strategy_XCy(U_temp)
        Case "XRp" : Strategy_XRp(U_temp)
        Case "XNl" : Strategy_XNl(U_temp)
        Case "WgX" : Strategy_WgX(U_temp)
        Case "WgY" : Strategy_WgY(U_temp)
        Case "WgZ" : Strategy_WgZ(U_temp)
        Case "WgW" : Strategy_WgW(U_temp)
      End Select
      If GRslt.Productivité Then Exit For
    Next i
    If Pzzl_Slv_UO(U_temp) Then Jrn_Add(, {"La grille est résolvable en CdU et CdO."})

  End Sub
  Private Sub Mnu0902_Click(sender As Object, e As EventArgs) Handles Mnu0902.Click
    ' Cette option permet de supprimer les candidats et rétablit l'affichage standard
    Jrn_Add(, {Procédure_Name_Get() & " Suppression des Candidats Exclues "})
    Jrn_Add(, {"Stratégie en cours: " & Plcy_Strg})
    Select Case Plcy_Strg
      Case "Gbl", "Gbv", "GCs"
        For Each XCel As GCel_Excl_Cls In GRslt.CelExcl
          Cell_Cdd_Exclude(XCel.Cdd, XCel.Cel)
        Next XCel
      Case "XCx", "XCy", "Xrp", "XNl", "WgX", "WgY", "WgZ", "WgW"
        For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
          Cell_Cdd_Exclude(XCel.Cdd, XCel.Cel)
        Next XCel
    End Select
    Mnu0902.Enabled = False
    Strategy_Dsp_Standard()
  End Sub

  Private Sub Mnu0910_GLk_Click(sender As Object, e As EventArgs) Handles Mnu0910_GLk.Click
    Dim U_temp(80, 3) As String
    Array.Copy(U, U_temp, UNbCopy)
    Strategy_GLk(U_temp)
  End Sub

  Private Sub Mnu0930_Gbl_Click(sender As Object, e As EventArgs) Handles Mnu0930_Gbl.Click
    Dim U_temp(80, 3) As String
    Array.Copy(U, U_temp, UNbCopy)
    Strategy_Gbl(U_temp)
  End Sub

  Private Sub Mnu0950_Gbv_Click(sender As Object, e As EventArgs) Handles Mnu0950_Gbv.Click
    Dim U_temp(80, 3) As String
    Array.Copy(U, U_temp, UNbCopy)
    Strategy_Gbv(U_temp)
  End Sub

  Private Sub Mnu0970_GCs_Click(sender As Object, e As EventArgs) Handles Mnu0970_GCs.Click
    Dim U_temp(80, 3) As String
    Array.Copy(U, U_temp, UNbCopy)
    Strategy_GCs(U_temp)
  End Sub
#End Region
End Class