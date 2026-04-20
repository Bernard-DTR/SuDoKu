Imports System.Threading
Imports SuDoKu.DancingLink

Public NotInheritable Class Frm_SDK
  Private MouseClick_Middle_ToolTip As CustomToolTip

  Public Journal As New RichTextBox()
  Public B_Famille As New TextBox        ' comporte désormais la famille de la stratégie jouée
  Public B_Position As New TextBox()
  Public B_Pourcentage As New TextBox()
  Public B_Info As New TextBox()

  Public B_ProgressBar As New ProgressBar()
  'La ProgressBar ne peut pas adopter la couleur souhaitée
  Dim Prv_MM_Pt As Point
  Dim Prv_Rct_Cdd_Numéro As Integer
  Public Sub New()
    ' Cet appel est requis par le concepteur.
    InitializeComponent()
  End Sub

  Private Sub Frm_SDK_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'Interruption temporaire de la logique de présentation du formulaire, pendant la mise en place des contrôles
    SuspendLayout()
    Phase_Démarrage_Terminée = False
    OO_000_SDK_Load()
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
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

    With B_Famille
      .Name = "B_Famille"
      .ReadOnly = True
      .ShortcutsEnabled = False 'Pour ne pas avoir de menu contextuel
      .TextAlign = HorizontalAlignment.Center
    End With
    Controls.Add(B_Famille)

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

    '#713
    AddHandler TTT_Timer.Tick, AddressOf TTT_Timer_Tick

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

    Mnu07n_Gbl.Text = Stg_Get("Gbl").Texte
    Mnu07n_Gbv.Text = Stg_Get("Gbv").Texte
    Mnu07n_GCs.Text = Stg_Get("GCs").Texte
    Mnu07n_GCx.Text = Stg_Get("GCx").Texte
    Mnu07n_XCy.Text = Stg_Get("XCy").Texte
    Mnu07n_XRp.Text = Stg_Get("XRp").Texte
    Mnu07n_XNl.Text = Stg_Get("XNl").Texte
    Mnu07n_WgX.Text = Stg_Get("WgX").Texte
    Mnu07n_WgY.Text = Stg_Get("WgY").Texte
    Mnu07n_WgZ.Text = Stg_Get("WgZ").Texte
    Mnu07n_WgW.Text = Stg_Get("WgW").Texte
    Mnu08_Résolution.Text = "Résolution SDK _ Stratégies " & Stg_Profondeur
    Mnu0902.Enabled = False

    Mnu0910_GLk.Text = Stg_Get("GLk").Texte & "..."
    Mnu0915_Gbl.Text = Stg_Get("Gbl").Texte
    Mnu0920_Gbv.Text = Stg_Get("Gbv").Texte
    Mnu0925_GCs.Text = Stg_Get("GCs").Texte
    Mnu0930_GCx.Text = Stg_Get("GCx").Texte
    Mnu0935_XCy.Text = Stg_Get("XCy").Texte
    Mnu0940_XRp.Text = Stg_Get("XRp").Texte
    Mnu0945_XNl.Text = Stg_Get("XNl").Texte
    Mnu0950_WgX.Text = Stg_Get("WgX").Texte
    Mnu0955_WgY.Text = Stg_Get("WgY").Texte
    Mnu0960_WgZ.Text = Stg_Get("WgZ").Texte
    Mnu0965_WgW.Text = Stg_Get("WgW").Texte

    Dim Mnu04n_RésoudreUneCellule As New ToolStripMenuItem() With
        {
        .Text = "Résoudre Une Cellule",
        .ForeColor = SystemColors.ControlText,
        .ShortcutKeys = Keys.F11
         }
    AddHandler Mnu04n_RésoudreUneCellule.Click, AddressOf Mnu04n_RésoudreUneCellule_Click
    Nsd_i = Mnu04.DropDown.Items.Add(Mnu04n_RésoudreUneCellule)

#End Region

#Region "Mnu06_Divers Internet Explorer et Url diverses"
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE1"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE2"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE3"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE4"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE5"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE6"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE7"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE8"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IE9"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IEa"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IEb"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IEc"))
    Nsd_i = Mnu06_CB.Items.Add(Msg_Read("MNU_06IEd"))
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
      .SetToolTip(B_Famille, "Famille Stratégie")
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
    Jrn_Add(, {"Traitement de création de grilles en Batch : " & CStr(Plcy_Generate_Batch)})
    Jrn_Add(, {"Nombre de grilles à créer en Batch         : " & CStr(My.Settings.Prf_02C_Nb_Batch_Generate)})
    Jrn_Add(, {"Nombre de grilles existantes dans " & Path_Batch})
    Jrn_Add(, {"          Faciles                        F : " & CStr(File_Nb("SDK_F"))})
    Jrn_Add(, {"          Moyennes                       M : " & CStr(File_Nb("SDK_M"))})
    Jrn_Add(, {"          Difficiles                     D : " & CStr(File_Nb("SDK_D"))})
    Jrn_Add(, {"          Expertes                       E : " & CStr(File_Nb("SDK_E"))})
    Jrn_Add(, {Proc_Name_Get()})
    Jrn_Add("SDK_00010", JourDateHeure())
    Jrn_Add("SDK_00101")
    Plcy_Gnrl = My.Settings.LP_Plcy_Gnrl
    LP_Nom = My.Settings.LP_Nom
    LP_Prb = My.Settings.LP_Prb
    LP_Jeu = My.Settings.LP_Jeu
    LP_Sol = My.Settings.LP_Sol
    LP_Frc = My.Settings.LP_Frc
    LP_Cdd = ""
    LP_CddExc = ""
    'Jrn_Add("SDK_00100", {LP_Nom})
    Jrn_Add(, {"/" & Proc_Name_Get()})

    OC_Présentation()
    Game_New_Game(Plcy_Gnrl, "   ", LP_Nom, LP_Prb, LP_Jeu, LP_Sol, LP_Cdd, LP_Frc, Proc_Name_Get())
#End Region

    'Reprise de la logique de présentation du formulaire, les contrôles sont mis en place
    Phase_Démarrage_Terminée = True
    ResumeLayout()

    ' Pour afficher la barre d'outils avec les valeurs 0 à 1 de Arial
    Mnu_Mngt_Barre_Outils_Filtres()

    'Plcy_Generate_Batch autorise la création de grilles par lot en arrière-plan
    If Plcy_Generate_Batch Then
      Batch_Timer.Interval = 5000
      'Process_16x est un png 16x16 32 bits
      'Me.Mnu08.Size = New System.Drawing.Size(98, 32)
      Mnu08.Image = SuDoKu.My.Resources.Resources.Process_16x
      Mnu08.Font = New Font(Mnu08.Font, FontStyle.Italic)
      Batch_Initial()
    End If
    Build_Fond_Cellule_Survolee()

  End Sub

  Private Sub TTT_Timer_Tick(sender As Object, e As EventArgs)
    Try
      If MouseClick_Middle_ToolTip IsNot Nothing Then
        MouseClick_Middle_ToolTip.HideTooltip()
      End If
    Catch
      ' Ne pas propager l'erreur depuis le timer
    End Try
    TTT_Timer.Stop()
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
    MyBase.OnPaint(e)
    If Not Phase_Démarrage_Terminée Then Exit Sub
    Dim g As Graphics = e.Graphics

    'Le quadrillage n'est pas redessiné, c'est un bitmap qui est affiché, ce qui améliore les performances d'affichage
    g.DrawImageUnscaled(Bmp_Quadrillage, 0, 0)
    g.DrawImageUnscaled(Bmp_Fond_Valeur, 0, 0)
    G4_Grid_Stratégie_All(g)

    'If Grille.Nb_Cellules_Remplies = 81 Then Grille.G8_Grille_Partie_Terminée(g)

    If Cellule_Survolee >= 0 Then
      If U(Cellule_Survolee, 2) = " " Then
        g.DrawImage(Bmp_Fond_Cellule_Survolee, Sqr_Cel(Cellule_Survolee).X, Sqr_Cel(Cellule_Survolee).Y)
      End If
    End If

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
    My.Settings.LP_Plcy_Gnrl = "Nrm"

    My.Settings.Save()

    If Batch_Thread IsNot Nothing AndAlso Batch_Thread.IsAlive Then
      Enabled = False
      Cursor = Cursors.WaitCursor

      ' Attendre la fin du thread
      Invalidate()
      MsgBox("SuDoKu est en train de calculer des grilles " & vbCrLf & "Merci de patienter ! ",
      MsgBoxStyle.Information, "SuDoKu")
      Batch_Thread.Join()

      ' Restaurer l'état de l'interface
      Cursor = Cursors.Default
      Enabled = True
    End If

    RemoveHandler TTT_Timer.Tick, AddressOf TTT_Timer_Tick
  End Sub

#Region "Mouse Clic"
  Private Sub Frm_SDK_MouseMove(sender As Object, e As MouseEventArgs) Handles MyBase.MouseMove
    'Dim Cellule_MM As Integer = Array.FindIndex(Sqr_Cel, Function(cel) cel.Contains(e.X, e.Y))
    Dim Cellule_MM As Integer = Wh_Cellule_Pt(pt:=New Point(e.X, e.Y))
    If Cellule_MM = -1 Then Exit Sub
    Pbl_Cell_Select = Cellule_MM
    If Prv_Pbl_Cell_Select <> -1 And Prv_Pbl_Cell_Select <> Pbl_Cell_Select Then
      B_Position.Text = U_cr(Pbl_Cell_Select) & " (" & Pbl_Cell_Select & ")"
      Mnu_Mngt(Pbl_Cell_Select)
    End If
    Prv_Pbl_Cell_Select = Pbl_Cell_Select
    If (Stg_Get(Plcy_Strg).Family = 0 Or Stg_Get(Plcy_Strg).Family = 2) AndAlso U(Pbl_Cell_Select, 2) = " " Then
      If Pbl_Cell_Select <> Cellule_Survolee Then
        If Cellule_Survolee >= 0 Then
          Me.Invalidate(Sqr_Cel(Cellule_Survolee))  ' Invalider l’ancienne cellule
        End If
        If Pbl_Cell_Select >= 0 Then
          Me.Invalidate(Sqr_Cel(Pbl_Cell_Select))   ' Invalider la nouvelle cellule
        End If
        Cellule_Survolee = Pbl_Cell_Select
      End If
    End If
  End Sub
  Private Sub Frm_SDK_MouseLeave(sender As Object, e As EventArgs) Handles MyBase.MouseLeave
    If Cellule_Survolee >= 0 Then
      Me.Invalidate(Sqr_Cel(Cellule_Survolee))
    End If
    Cellule_Survolee = -1
  End Sub

  Private Sub Frm_SDK_MouseClick(sender As Object, e As MouseEventArgs) Handles MyBase.MouseClick
    ' e as MouseEventArgs permet de localiser la souris
    ' seuls les clics gauche et milieu sont détectés, le clic droit affiche le menu contextuel
    ' Après un clic de souris sur le contrôle
    ' L'utilisateur appuie sur le bouton de la souris
    Try
      'Il n'est pas besoin de calculer la cellule, c'est Pbl_Cell_Select
      If Pbl_Cell_Select = -1 Then Exit Sub
      Dim Candidat_Pt As Integer = Wh_Cellule_Candidat_Pt(Pbl_Cell_Select, New Point(x:=e.X, y:=e.Y))
      'Le mouse-click se fait soit sur un square, soit ailleurs
      'Il est de 3 types: Gauche, Milieu, Droit
      'Gauche: Rien de particulier, le menu contextuel est préparé,
      'Milieu: Affichage des candidats de l'unité si Plcy_MouseClick_Middle = True 
      'Droit : Le menu contextuel est affiché, ainsi le bouton droit n'est pas traité par Frm_SDK_MouseClick
      Select Case e.Button
        Case MouseButtons.Left
          Cell_Val_Insert(CStr(Candidat_Pt), Pbl_Cell_Select, "Mse_Clk")

        Case MouseButtons.Middle
          ' Affiche les candidats de l'unité en fonction de la position du clic et de l'option Afficher les candidats en InfoBulle
          ' Affichage des candidats de la Ligne (Donc colonne 0 ou 8)
          If Plcy_MouseClick_Middle Then Frm_SDK_MouseClick_Middle(sender, Candidat_Pt)

        Case MouseButtons.Right
          'le clic droit n'est pas détecté sur les Cellules R et V, car il y a un ContextMenuStrip sur le formulaire
          'le clic droit est détecté sur les Cellules       I, car il n'y a pas de ContextMenuStrip
      End Select
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub

  Private Sub Frm_SDK_MouseClick_Middle(sender As Object, Candidat As Integer)
    ' Provient UNIQUEMENT de Frm_SDK_Mouse_Click
    Dim TTT_Message As String = Cnddts_Blancs
    Dim Cellule As Integer = Pbl_Cell_Select
    Pénalités(Proc_Name_Get() & " " & U_Coord(Cellule) & "(" & Candidat & ")")
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
    '#713
    'TTT_Timer.Interval = 2000
    'AddHandler TTT_Timer.Tick, Sub(senderObj As Object, eventArgs As EventArgs)
    '                             MouseClick_Middle_ToolTip.HideTooltip()
    '                             TTT_Timer.Stop()
    '                           End Sub
    'TTT_Timer.Start()
    TTT_Timer.Interval = 2000
    TTT_Timer.Start()
  End Sub

  Private Sub Frm_SDK_MouseWheel(sender As Object, e As MouseEventArgs) Handles MyBase.MouseWheel
    ' Stratégie des Filtres "FV1" To "FV9" et "FC1" to "FC9"
    ' Détermination du sens de défilement
    Dim ScrollDelta As Integer = e.Delta * SystemInformation.MouseWheelScrollLines \ SystemInformation.MouseWheelScrollDelta
    Dim Sens As Integer = Math.Sign(ScrollDelta)

    If Plcy_Gnrl = "Nrm" AndAlso Plcy_Strg.StartsWith("FV") Then MouseWheel_Valeur(Sens)
    If Plcy_Gnrl = "Nrm" AndAlso Plcy_Strg.StartsWith("FC") Then MouseWheel_Candidat(Sens)
  End Sub
  Public Sub MouseWheel_Valeur(Sens As Integer)
    If Not Integer.TryParse(Plcy_Strg.Substring(2, 1), MW_Val) Then Exit Sub
    Dim Result As Wh_Nb_Cell_Struct = Wh_Nb_Cell(U)    ' Compter les occurrences de chaque valeur sur la grille

    Dim StartVal As Integer = MW_Val
    Do
      MW_Val = ((MW_Val + Sens + 8) Mod 9) + 1
      If Result.Val_Nb(MW_Val) < 9 Then Exit Do
      ' Si la valeur n'est pas présente 9 fois, on la présente
    Loop While MW_Val <> StartVal

    Plcy_Strg = "FV" & CStr(MW_Val)
    Pénalités("Stratégie " & Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte)
    Dim MW_Cell_Last As Integer = -1
    MW_Cell_List.Clear()
    For i As Integer = 0 To 80
      If U(i, 2) = CStr(MW_Prv_Val) Or U(i, 2) = CStr(MW_Val) Then
        MW_Cell_List.Add(i)
        MW_Cell_Last = i
      End If
    Next

    If MW_Cell_Last <> -1 Then
      Using reg As New Region(Sqr_Pth(MW_Cell_Last))
        ' On invalide TOUTES les cellules filtrées précédentes et en cours  
        For Each cell As Integer In MW_Cell_List
          reg.Union(Sqr_Pth(cell))
        Next
        B_Info.Text = Stg_Get(Plcy_Strg).Texte
        Invalidate(reg, False)
      End Using
    End If
    MW_Prv_Val = MW_Val
  End Sub
  Public Sub MouseWheel_Candidat(ByVal Sens As Integer)
    Dim FiltreMW As Integer
    If Not Integer.TryParse(Plcy_Strg.Substring(2, 1), FiltreMW) Then Exit Sub
    Dim NewMW As Integer = ((FiltreMW + Sens + 8) Mod 9) + 1
    Plcy_Strg = "FC" & CStr(NewMW)
    Pénalités("Stratégie " & Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte)
    Invalidate()
  End Sub
#End Region

#Region "Menu"
  '-------------------------------------------------------------------------------
  '
  ' Menus
  '
  '--------------01---------------------------------------------------------------
  Private Sub Mnu01_Ouvrir_Click(sender As Object, e As EventArgs) Handles Mnu01_Ouvrir.Click
    'Chargement d'une nouvelle partie en mode normal 
    Frm_LoadParties.Show()
    Invalidate()
  End Sub
  Private Sub Mnu01_RejouerLaPartie_Click(sender As Object, e As EventArgs) Handles Mnu01_RejouerLaPartie.Click
    Game_New_Game(Plcy_Gnrl, "   ", LP_Nom, LP_Prb, LP_Prb, LP_Sol, Cdd729:=StrDup(729, " "), LP_Frc, Proc_Name_Get())
    Invalidate()
  End Sub
  Private Sub Mnu01_Saisir_Click(sender As Object, e As EventArgs) Handles Mnu01_Saisir.Click
    ' Pendant la saisie les nombres des cellules vides et des candidats sont affichés
    '         le nombre des valeurs initiales ne peut pas l'être, il faut attendre / Commencer
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
    'Utilisation de la stratégie
    'Stg_List.Add(New Stg_Cls("Sai", "N", "N", "N", "N", 0, "Saisir une grille"))
    Plcy_Gnrl = "Nrm"
    Plcy_Strg = "Sai"
    Game_New_Game(Plcy_Gnrl, Plcy_Strg, Nom, Prb, Jeu, Sol, StrDup(729, " "), Frc, Proc_Name_Get())
    Mnu01_Saisir.Enabled = False
    Mnu01_Commencer.Enabled = True
    B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()
    B_Info.Text = Stg_Get(Plcy_Strg).Texte

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
    'Utilisation de la stratégie
    'Stg_List.Add(New Stg_Cls("   ", "N", "N", "N", "N", 0, "Aucune Stratégie"))
    Plcy_Gnrl = "Nrm"
    Plcy_Strg = "   "
    Game_New_Game(Plcy_Gnrl, Plcy_Strg, Nom, Prb, Jeu, Sol, StrDup(729, " "), Frc, Proc_Name_Get())
    Mnu01_Saisir.Enabled = True
    Mnu01_Commencer.Enabled = False
  End Sub
  Private Sub Mnu01_EnregistrerUnePartieTest_Click(sender As Object, e As EventArgs) Handles Mnu01_EnregistrerUnePartieTest.Click
    Pzzl_Write_Partie_Test()
  End Sub
  Private Sub Mnu01_ChargerUnePartieTest_Click(sender As Object, e As EventArgs) Handles Mnu01_ChargerUnePartieTest.Click
    Pzzl_Load_Partie_Test()
  End Sub
  Private Sub Mnu01_OuvrirLaBibliothèqueTestDeHodoku_Click(sender As Object, e As EventArgs) Handles Mnu01_OuvrirLaBibliothèqueTestDeHodoku.Click
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})
    Frm_LoadPartiesHodoku.Show()
  End Sub
  Private Sub Mnu01_CopierLeJournalEnModeRTF_Click(sender As Object, e As EventArgs) Handles Mnu01_CopierLeJournalEnModeRTF.Click
    Dim File_SDK As String = Jrn_RcdRTF()
    Processing_Start(File_SDK)
  End Sub
  Private Sub Mnu01_OuvrirLeRépertoire_Click(sender As Object, e As EventArgs) Handles Mnu01_OuvrirLeRépertoire.Click
    Dim Pgm As String = "Explorer /e, /n, " & File_SDK
    Nsd_i = Shell(Pgm, AppWinStyle.NormalFocus)
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
  Private Sub Mnu03_AfficherLaSolution_Click(sender As Object, e As EventArgs) Handles Mnu03_AfficherLaSolution.Click
    If Plcy_Solution_Existante = False Then Exit Sub
    'Sauvegarde du jeu en cours
    Dim jeu_Save As String = ""
    For i As Integer = 0 To 80 : jeu_Save &= U(i, 2) : Next i
    For i As Integer = 0 To 80
      If U(i, 1) = " " And U_Sol(i) <> " " Then U(i, 2) = U_Sol(i)
    Next i
    Build_Fond_Valeur()
    B_Info.Text = "Affichage de la Solution"
    Invalidate()
    Application.DoEvents()    'Affiche la grille sans solutions immédiatement

    Thread.Sleep(2000) 'Le temps de lire quelques valeurs

    For i As Integer = 0 To 80 : U(i, 2) = jeu_Save.Substring(i, 1) : Next i
    Build_Fond_Valeur()
    B_Info.Text = " _ "
    Invalidate()
    Application.DoEvents()
  End Sub
  Private Sub Mnu03_Rafraîchir_Click(sender As Object, e As EventArgs) Handles Mnu03_Rafraîchir.Click
    Invalidate()
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
  ' TOUTES les stratégies de la barre d'outils sont lancées ici
  Private Sub Btn_CUO_BTXYSJZKQ_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn_CdU.MouseDown, Btn_CdO.MouseDown, Btn_Cdd.MouseDown, Btn_Cbl.MouseDown, Btn_Xwg.MouseDown, Btn_Tpl.MouseDown, Btn_XYw.MouseDown, Btn_Swf.MouseDown, Btn_XYZ.MouseDown, Btn_Unq.MouseDown, Btn_SKy.MouseDown, Btn_Jly.MouseDown
    Dim Lettre As String = sender.ToString()
    Dim Strg_Classique As String = "###"
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Lettre = Lettre Then Strg_Classique = Stg.Code
    Next Stg

    Select Case e.Button
      Case MouseButtons.Left
        ' le bouton gauche exécute la stratégie et affiche 1 seul résultat
        Strategy_Code(Strg_Classique, Proc_Name_Get())
      Case MouseButtons.Right
        ' Le bouton droit affiche tous les résultats de la stratégie sauf pour Cdd
        Dim U_temp(80, 3) As String
        Dim Strategy_Rslt(99, 0) As String
        Array.Copy(U, U_temp, UNbCopy)
        Select Case Lettre
          Case "U" : Strategy_Rslt = Strategy_CdU(U_temp)
          Case "O" : Strategy_Rslt = Strategy_CdO(U_temp)
          Case "B" : Strategy_Rslt = Strategy_Cbl(U_temp)
          Case "T" : Strategy_Rslt = Strategy_Tpl(U_temp)
          Case "X" : Strategy_Rslt = Strategy_Xwg(U_temp)
          Case "Y" : Strategy_Rslt = Strategy_XYw(U_temp)
          Case "S" : Strategy_Rslt = Strategy_Swf(U_temp)
          Case "J" : Strategy_Rslt = Strategy_Jly(U_temp)
          Case "Z" : Strategy_Rslt = Strategy_XYZ(U_temp)
          Case "K" : Strategy_Rslt = Strategy_SKy(U_temp)
          Case "Q" : Strategy_Rslt = Strategy_Unq(U_temp)
          Case Else : Exit Sub
        End Select
        Strategy_Rslt_Display(Strategy_Rslt, -1)
    End Select
    Pénalités("Stratégie " & Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte)

  End Sub
  Private Sub Btn0_Click(sender As Object, e As EventArgs) Handles Btn0.Click
    Strategy_Dsp_Standard()
  End Sub

  Private Sub Btn123456789_MouseDown(sender As Object, e As MouseEventArgs) Handles Btn9.MouseDown, Btn8.MouseDown, Btn7.MouseDown, Btn6.MouseDown, Btn5.MouseDown, Btn4.MouseDown, Btn3.MouseDown, Btn2.MouseDown, Btn1.MouseDown
    Dim btn As ToolStripButton = DirectCast(sender, ToolStripButton)
    Dim flt As String = btn.Name.Substring(3, 1)
    Select Case e.Button
      Case MouseButtons.Left
        Strategy_Code("FV" & flt, Proc_Name_Get())
      Case MouseButtons.Right
        Strategy_Code("FC" & flt, Proc_Name_Get())
    End Select
    Pénalités("Stratégie " & Plcy_Strg & " " & Stg_Get(Plcy_Strg).Texte)
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
        Case "GCx" : Plcy_Strg = "GCx" : Strategy_GCx(U_temp)
        Case "XCy" : Plcy_Strg = "XCy" : Strategy_XCy(U_temp)
        Case "XRp" : Plcy_Strg = "XRp" : Strategy_XRp(U_temp)
        Case "XNl" : Plcy_Strg = "XNl" : Strategy_XNl(U_temp)
        Case "WgX" : Plcy_Strg = "WgX" : Strategy_WgX(U_temp)
        Case "WgY" : Plcy_Strg = "WgY" : Strategy_WgY(U_temp)
        Case "WgZ" : Plcy_Strg = "WgZ" : Strategy_WgZ(U_temp)
        Case "WgW" : Plcy_Strg = "WgW" : Strategy_WgW(U_temp)
        Case Else : Jrn_Add(, {"Mnu_Name inconnu : " & Mnu_Name, "Erreur"})
      End Select
      B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()

    Else
      Jrn_Add(, {"Sender inconnu : " & Sender.ToString(), "Erreur"})
    End If
  End Sub
  Private Sub Mnu04n_RésoudreUneCellule_Click(sender As Object, e As EventArgs)
    Cell_Slv_Interactif("S", "Résoudre une Cellule")
  End Sub
  '--------------04---------------------------------------------------------------
  '--------------05---------------------------------------------------------------
  Private Sub Mnu05_APropos_Click(sender As Object, e As EventArgs) Handles Mnu05_APropos.Click
    Frm_About.Show()
  End Sub
  Private Sub Mnu05_Préférences_Click(sender As Object, e As EventArgs) Handles Mnu05_Préférences.Click
    Frm_Préférences.Show()
  End Sub
  Private Sub Mnu05_FichierDesMessages_Click(sender As Object, e As EventArgs) Handles Mnu05_FichierDesMessages.Click
    Processing_Start(File_SDKMsg)
  End Sub
  Private Sub Mnu05_Documentation_Click(sender As Object, e As EventArgs) Handles Mnu05_Documentation.Click
    'Fichier de Documentation
    Processing_Start(File_SDKDoc)
  End Sub
  Private Sub Mnu05_Maintenance_Click(sender As Object, e As EventArgs) Handles Mnu05_Maintenance.Click
    Nsd_i = Shell($"Notepad {Path_SDK}SuDoKu\SuDoKu\Apriori\aMaintenance.txt", AppWinStyle.MaximizedFocus)
    SendKeys.Send("^{END}")
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
    Processing_Start(File)
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
    Game_New_Game(Gnrl:="Nrm", Strg:="   ", Nom:=Nom, Prb:=Prb, Jeu:=Jeu, Sol:=Sol, Cdd729:=StrDup(729, " "), Frc:=Frc, Proc_Name_Get())
    Mnu01_Saisir.Enabled = True
    Mnu01_Commencer.Enabled = True
  End Sub
  Private Sub Mnu06_SudokuPCA_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuPCA.Click
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "PCA_Sudoku\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuAngusJohnson_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuAngusJohnson.Click
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "SimpleSudoku\simplesudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuDarrenColes_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuDarrenColes.Click
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "WinSudoku\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuPatriceHenrion_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuPatriceHenrion.Click
    'Ligne de commande /p lance le jeu en petite taille
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "SudokuPH\Sudoku.exe /p"
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
    Processing_Start(Shell_St)
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
    Dim Shell_St As String = "C:\Users\berna\AppData\Local\Programs\Python\Python313\pythonw.exe " & Path_SDK_Autres_Jeux & "Python\FedynaK\Sudoku_FedynaK\Sudoku_FedynaK.py " & Grille
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
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "SudoCue\SudoCue.exe"

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
    Processing.StartInfo.Arguments = "/c """ & Path_SDK_Autres_Jeux & "DiufSudoku\gui.bat"""
    'gui.bat comporte java -Xms512m -Xmx2G -jar dist\DiufSudoku.jar SudokuExplainer.log
    Processing.StartInfo.WorkingDirectory = Path_SDK_Autres_Jeux & "DiufSudoku"
    Processing.StartInfo.CreateNoWindow = False ' True = cacher la fenêtre noire
    Processing.StartInfo.UseShellExecute = False
    Processing.Start()
  End Sub

  Private Sub Mnu06_MUDancingLink_Click(sender As Object, e As EventArgs) Handles Mnu06_MUDancingLink.Click
    ClipBoard_Copier_New("1") ' Copier les valeurs initiales de SDK

    Dim Shell_St As String = Path_SDK_Autres_Jeux & "Dancing_Link\Dancing_Links_Library_100\Dancing_Links_Library\Dancing_Links_Library\bin\Debug\Dancing_Links_Library.exe"
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
    Dim Shell_St As String = Msg_Read("MNU_06IE0")
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
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "CP_Sudoku\SudokuSolver_src\SudokuSolver src\bin\Debug\SudokuSolver.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_SudokuSolver_Click(sender As Object, e As EventArgs) Handles Mnu06_SudokuSolver.Click
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "Planete\VB_Net_Sud1916447222005\Sudoku\bin\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_MicrosoftSudoku_Click(sender As Object, e As EventArgs) Handles Mnu06_MicrosoftSudoku.Click
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "Sudoku_VisualBasic\bin\Debug\Sudoku.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  Private Sub Mnu06_HHSudokuGame_Click(sender As Object, e As EventArgs) Handles Mnu06_HHSudokuGame.Click
    Dim Shell_St As String = Path_SDK_Autres_Jeux & "SudokuGame\SudokuGame\bin\Debug\SudokuPuzzle.exe"
    Jrn_Add(, {Shell_St})
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub
  '--------------07---------------------------------------------------------------
  Private Sub Mnu07_Outils(sender As Object, e As EventArgs) Handles Mnu07n_Gbl.Click, Mnu07n_XRp.Click, Mnu07n_XNl.Click, Mnu07n_XCy.Click, Mnu07n_WgZ.Click, Mnu07n_WgY.Click, Mnu07n_WgX.Click, Mnu07n_WgW.Click, Mnu07n_GCx.Click, Mnu07n_GCs.Click, Mnu07n_Gbv.Click
    If TypeOf sender Is ToolStripMenuItem Then
      Dim Mnu As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
      Dim Mnu_Name As String = Mnu.Name
      Dim Stratégie_XW As String = Mnu_Name.Substring(7, 3)

      Dim File_Name As String = File_SDKEtd & Stratégie_XW & ".txt"
      Jrn_Add(, {"Stratégie Test fichier : " & File_Name})
      If Plcy_Open_Display Then Processing_Start(File_Name)
      Pzzl_Open(File_Name)
      Mnu04n_Stratégie_XW_Click(sender, e)
    End If
  End Sub

  '--------------08---------------------------------------------------------------
  Private Sub Mnu08_Jouer_Click(sender As Object, e As EventArgs) Handles Mnu08_Jouer.Click
    Mnu08J_F.Text = "Facile" & " (" & CStr(File_Nb("SDK_F")) & ") "
    Mnu08J_M.Text = "Moyen" & " (" & CStr(File_Nb("SDK_M")) & ") "
    Mnu08J_D.Text = "Difficile" & " (" & CStr(File_Nb("SDK_D")) & ") "
    Mnu08J_E.Text = "Expert" & " (" & CStr(File_Nb("SDK_E")) & ") "
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
    Jrn_Add(, {Proc_Name_Get() & " Difficulté: " & Difficulté})
    'Extension / Jouer un Sudoku FMDE
    Dim s, l As Integer
    ' Classique "shuffle LINQ" (mélange)
    Dim rnd As New Random()
    Dim Files As IEnumerable(Of String) =
        IO.Directory.GetFiles(Path_Batch).Where(Function(f) f.Contains("SDK_" & Difficulté)).OrderBy(Function(f) rnd.Next())
    If Files.Count > 0 Then
      For Each File As String In Files
        Pzzl_Open(File)        'Ouverture et lecture du premier fichier
        s = InStrRev(File.ToString(), "\")
        l = File.ToString().Length
        Dim Nom_Physique As String = Mid$(File.ToString(), s + 1, l - s)         ' Extension comprise
        'Copie du fichier
        '05/06/2024 Une fois la partie lancée, elle est copiée dans Batch_Poubelle et pourra ainsi être contrôlée
        Dim Source As String = File
        Dim Destination As String = Path_Batch_Poubelle & Nom_Physique
        Try
          My.Computer.FileSystem.CopyFile(Source, Destination)
        Catch ex As Exception
          Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
          Jrn_Add("ERR_00000", {"Erreur CopyFile"}, "Erreur")
          Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
          Return
        End Try

        Try
          My.Computer.FileSystem.DeleteFile(File,
        Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
        Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently)
        Catch ex As Exception
          Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
          Jrn_Add("ERR_00000", {"Erreur DeleteFile"}, "Erreur")
          Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
        End Try

        'Affichage du fichier si Plcy_Open_Display
        'L'affichage permet de lire les commentaires et d'en ajouter  
        If Plcy_Open_Display Then Processing_Start(Destination)
        Exit For
      Next File
    Else
      Dim MsgTit As String = Proc_Name_Get() & " " & Difficulté & " " & Application.ProductName & " " & SDK_Version
      Nsd_i = MsgBox("Il n'a pas été trouvé de parties à jouer.",, MsgTit)
    End If
  End Sub
  Private Sub Mnu08_JouerAutrement_Click(sender As Object, e As EventArgs) Handles Mnu08_JouerAutrement.Click
    Jrn_Add(, {Proc_Name_Get()})
    Pzzl_Prd_DL()
  End Sub
  Private Sub Mnu08_Création_Click(sender As Object, e As EventArgs) Handles Mnu08_Création.Click
    Jrn_Add(, {Proc_Name_Get()})
    Pzzl_Prd_Interactif("P")
  End Sub
  Private Sub Mnu08_Résolution_Click(sender As Object, e As EventArgs) Handles Mnu08_Résolution.Click
    Jrn_Add(, {Proc_Name_Get()})
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
    Build_Fond_Valeur()
    Invalidate()
  End Sub
  Private Sub Mnu08_RésoudreEnForceBrute_Click(sender As Object, e As EventArgs) Handles Mnu08_RésoudreEnForceBrute.Click
    Jrn_Add(, {Proc_Name_Get()})
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
    Jrn_Add(, {Proc_Name_Get()})
    'La stratégie DancingLink ne fonctionne que si le grille est correcte
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Cursor.Current = Cursors.WaitCursor
    Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U)
    Jrn_Add(, {"Dancing Link        : " & CStr(DL.Nb_Solution)})
    Select Case DL.Nb_Solution
      Case -1 : Jrn_Add(, {"U_temp(,) est contrôlé Not U_Check.Check."})
      Case 0 : Jrn_Add(, {"U_temp(,) n'a pas de solution. "})
      Case 1
        Jrn_Add(, {"Dancing Link        : " & DL.DLCode})
        For i As Integer = 0 To 80
          U(i, 2) = DL.Solution(0).Substring(i, 1)
        Next i
        Build_Fond_Valeur()
        Invalidate()
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
  Private Sub Mnu08_RésoudreDenisBerthier_Click(sender As Object, e As EventArgs) Handles Mnu08_RésoudreDenisBerthier.Click
    Dim Grid As String = ClipBoard_Copier_New("1")
    Jrn_Add(, {Grid})
    Dim AllCandidates(728) As Candidate
    AllCandidates_Init(AllCandidates)
    SDK_AllCandidate(Grid, AllCandidates)
    PropagateSolvedCandidates(AllCandidates)
    Dim Solved As Boolean = DB_Solution(AllCandidates)
    If Solved Then Jrn_Add(, {"Les stratégies de Denis Berthier ont résolu la grille."}, "Red")
    AllCandidates_SDK(AllCandidates)
    Build_Fond_Valeur()
  End Sub
  Private Sub Mnu08_EditionDuProblème_Click(sender As Object, e As EventArgs) Handles Mnu08_EditionDuProblème.Click
    Jrn_Add(, {Proc_Name_Get()})
    'Public Swt_ModeEdition As Integer = -1 Position Initiale
    Swt_ModeEdition *= -1
    Select Case Swt_ModeEdition
      Case 1
        Plcy_Gnrl = "Edi"
        Mnu08_EditionDuProblème.Checked = True
        Jrn_Add("SDK_00321")
        ContextMenuStrip = Mnu_EDI
        B_Info.Text = Msg_Read("SDK_00320")
        B_Info.BackColor = Color.Gray
        ' TODO à voir comment afficher les candidats
        'Strategy_Dsp_Cdd()
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
    Jrn_Add(, {Proc_Name_Get()})
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
        B_Info.Text = Msg_Read("SDK_00330")
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
    Jrn_Add(, {Proc_Name_Get()})
    If Plcy_Solution_Existante = False Then Exit Sub
    For i As Integer = 0 To 80
      If U(i, 2) = " " Then U(i, 2) = U_Sol(i)
    Next i
    Build_Fond_Valeur()
    Invalidate()
  End Sub

#End Region

#Region "Menus Contextuels"
  Private Sub Mnu_Cel_Val_Insérer(sender As Object, e As EventArgs) Handles Mnu_Cel_Val_Ins_9.Click, Mnu_Cel_Val_Ins_8.Click, Mnu_Cel_Val_Ins_7.Click, Mnu_Cel_Val_Ins_6.Click, Mnu_Cel_Val_Ins_5.Click, Mnu_Cel_Val_Ins_4.Click, Mnu_Cel_Val_Ins_3.Click, Mnu_Cel_Val_Ins_2.Click, Mnu_Cel_Val_Ins_1.Click
    'Jrn_Add_Yellow(Proc_Name_Get() & " V " & sender.ToString(18) & " " & U_Coord(Pbl_Cell_Select) & " Mnu_Ctx_Ins")
    Cell_Val_Insert(sender.ToString(18), Pbl_Cell_Select, "Mnu_Ctx_Ins")
  End Sub
  Private Sub Mnu_Cel_Val_Effacer(sender As Object, e As EventArgs) Handles Mnu_Cel_Val_Eff_x.Click
    'Jrn_Add_Yellow(Proc_Name_Get() & " " & U_Coord(Pbl_Cell_Select) & " Mnu_Ctx_Eff")
    Cell_Val_Delete(Pbl_Cell_Select, "Mnu_Ctx_Eff")
  End Sub
  Private Sub Mnu_Cel_Cdd_Insérer(Sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Ins_9.Click, Mnu_Cel_Cdd_Ins_8.Click, Mnu_Cel_Cdd_Ins_7.Click, Mnu_Cel_Cdd_Ins_6.Click, Mnu_Cel_Cdd_Ins_5.Click, Mnu_Cel_Cdd_Ins_4.Click, Mnu_Cel_Cdd_Ins_3.Click, Mnu_Cel_Cdd_Ins_2.Click, Mnu_Cel_Cdd_Ins_1.Click
    'Jrn_Add_Yellow(Proc_Name_Get() & " V " & Sender.ToString(20) & " " & U_Coord(Pbl_Cell_Select) & " Mnu_Ctx")
    Cell_Cdd_Insert(Sender.ToString(20), Pbl_Cell_Select, "Mnu_Ctx")
  End Sub
  Private Sub Mnu_Cel_Cdd_Exclure(sender As Object, e As EventArgs) Handles Mnu_Cel_Cdd_Exc_9.Click, Mnu_Cel_Cdd_Exc_8.Click, Mnu_Cel_Cdd_Exc_7.Click, Mnu_Cel_Cdd_Exc_6.Click, Mnu_Cel_Cdd_Exc_5.Click, Mnu_Cel_Cdd_Exc_4.Click, Mnu_Cel_Cdd_Exc_3.Click, Mnu_Cel_Cdd_Exc_2.Click, Mnu_Cel_Cdd_Exc_1.Click
    'Jrn_Add_Yellow(Proc_Name_Get() & " V " & sender.ToString(20) & " " & U_Coord(Pbl_Cell_Select))
    Cell_Cdd_Exclude(sender.ToString(20), Pbl_Cell_Select)
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
    Dim Afficher As Boolean

    Select Case ClickedMenuItem.Name

      Case "A", "B", "C", "D"
        Obj_Symbol = ClickedMenuItem.Name
        Obj_Color = Color_BySymbol(Obj_Symbol)
        Couleurs_ResetAndColorizsation(C1Items, Color_List)
        Couleurs_Ckeck(C1Items, Obj_Symbol)
        Afficher = False

      Case "Cadre", "Carré", "Cercle", "Disque", "Croix"
        Obj_Forme = ClickedMenuItem.Name
        Objets_Reset(O1Items)
        Objets_CheckAndColor(O1Items, Obj_Forme, Color_BySymbol(Obj_Symbol))
        Afficher = False

      Case "Cel_Cdd"
        Dim Cellule As Integer = Pbl_Cell_Select
        Dim Candidat As Integer = Pbl_Cell_Candidat_Select
        Dim Cdd As Integer = 0
        ' Pour une cellule  Cellule est compris entre 0 et 80 et Cdd = 0
        ' Pour une candidat Cellule est compris entre 0 et 80 et Cdd est compris entre 1 et 9
        If U(Cellule, 3).Contains(CStr(Candidat)) Then Cdd = Candidat
        Objet_List.Add(New Objet_Cls With {.Symbol = Obj_Symbol, .Forme = Obj_Forme, .Cel_From = Cellule, .Cdd_From = Cdd, .Cel_To = -1, .Cdd_To = 0})
        Afficher = True

      Case "Flèche"
        ' Pour dessiner une flèche, il faut 2 candidats
        '                           1 candidat Origine et un candidat Destination
        ' Lorsque la flèche est dessinée, le candidat destination devient le candidat origine
        Obj_Forme = sender.ToString()
        Objets_Reset(O1Items)
        Objets_CheckAndColor(O1Items, Obj_Forme, Color_BySymbol(Obj_Symbol))
        Afficher = False

      Case "Origine"
        Flè_Cel_From = Pbl_Cell_Select
        Flè_Cdd_From = Pbl_Cell_Candidat_Select
        Flè_From = 0
        If U(Flè_Cel_From, 3).Contains(CStr(Flè_Cdd_From)) Then Flè_From = Flè_Cdd_From
        Afficher = False

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
        Afficher = True

      Case "Flèche_Supprimer"
        ' Pour supprimer la flèche, il faut connaître les 2 points de la flèche.
        'Candidat_Pt = Me.PointToClient(MousePosition) a été documenté à l'ouverture du menu contextuel
        '👉 Mnu_Obj_Opening
        ' Trouver l'objet correspondant
        Dim ligneàSupprimer As Objet_Cls = Objet_List.FirstOrDefault(Function(l) Get_Distance_Point_Flèche(Pbl_PtF, l.Point_From, l.Point_To) < 5)
        If ligneàSupprimer IsNot Nothing Then
          Objet_List.Remove(ligneàSupprimer) ' Supprimer l'objet trouvé
        End If
        Afficher = True

      Case "Enlever"
        Dim Cellule As Integer = Pbl_Cell_Select
        Dim Candidat As Integer = Pbl_Cell_Candidat_Select
        Dim Cdd As Integer = 0
        ' Pour une cellule  Cellule est compris entre 0 et 80 et Cdd = 0
        ' Pour une candidat Cellule est compris entre 0 et 80 et Cdd est compris entre 1 et 9
        If U(Cellule, 3).Contains(CStr(Candidat)) Then Cdd = Candidat
        Objet_List.RemoveAll(Function(Obj) Obj.Cel_From = Cellule And Obj.Cdd_From = Cdd And Obj.Cel_To = -1 And Obj.Cdd_To = 0)
        Afficher = True

      Case "Enlever_Tout"
        Objet_List.Clear()                   ' Vider Objet_List
        Afficher = True

      Case "Lister"
        Objet_List_Display()

      Case Else
        Jrn_Add(, {ClickedMenuItem.Name & " " & sender.ToString() & " en cours!"})
    End Select

    My.Settings.Obj_Symbol = Obj_Symbol
    My.Settings.Obj_Forme = Obj_Forme
    My.Settings.Save()

    If Afficher Then
      Invalidate()
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
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get() & " Lancement des Stratégies G "})
    Plcy_Strg = "   "
    B_Info.Text = Proc_Name_Get()
    Invalidate()
    Application.DoEvents()
    Dim U_temp(80, 3) As String

    For i As Integer = 0 To Stg_List_Link.Count - 1
      Plcy_Strg = Stg_List_Link(i)
      Jrn_Add("SDK_Space")
      Jrn_Add(, {"Strategie " & Plcy_Strg})
      Array.Copy(U, U_temp, UNbCopy)

      GRslt.Productivité = False
      XRslt.Productivité = False

      Select Case Plcy_Strg
        Case "Gbl" : Strategy_Gbl(U_temp)
        Case "Gbv" : Strategy_Gbv(U_temp)
        Case "GCs" : Strategy_GCs(U_temp)
        Case "GCx" : Strategy_GCx(U_temp)
        Case "XCy" : Strategy_XCy(U_temp)
        Case "XRp" : Strategy_XRp(U_temp)
        Case "XNl" : Strategy_XNl(U_temp)
        Case "WgX" : Strategy_WgX(U_temp)
        Case "WgY" : Strategy_WgY(U_temp)
        Case "WgZ" : Strategy_WgZ(U_temp)
        Case "WgW" : Strategy_WgW(U_temp)
      End Select
      B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()
      If GRslt.Productivité Or XRslt.Productivité Then Exit For
    Next i
    If Pzzl_Slv_UO(U_temp) Then Jrn_Add(, {"La grille est désormais résolvable en CdU_CdO."}, "Red")

  End Sub
  Private Sub Mnu0902_Click(sender As Object, e As EventArgs) Handles Mnu0902.Click
    ' Cette option permet de supprimer les candidats et rétablit l'affichage standard
    Jrn_Add(, {Proc_Name_Get() & " Suppression des Candidats Exclus "})
    Jrn_Add(, {"Stratégie en cours: " & Plcy_Strg})
    Select Case Plcy_Strg
      Case "Gbl", "Gbv", "GCs", "GCx"
        Cell_Cdd_Exclude_GRslt()
      Case "XCy", "XRp", "XNl", "WgX", "WgY", "WgZ", "WgW"
        For Each XCel As XCel_Excl_Cls In XRslt.CelExcl
          Cell_Cdd_Exclude(XCel.Cdd, XCel.Cel)
        Next XCel
    End Select
    Mnu0902.Enabled = False
    Strategy_Dsp_Standard()
    Jrn_Add(, {"Les candidats sont supprimés"})
  End Sub

  Private Sub Mnu09_Exec(sender As Object, e As EventArgs) Handles Mnu0965_WgW.Click, Mnu0960_WgZ.Click, Mnu0955_WgY.Click, Mnu0950_WgX.Click, Mnu0945_XNl.Click, Mnu0940_XRp.Click, Mnu0935_XCy.Click, Mnu0930_GCx.Click, Mnu0925_GCs.Click, Mnu0920_Gbv.Click, Mnu0915_Gbl.Click, Mnu0910_GLk.Click
    If TypeOf sender Is ToolStripMenuItem Then
      Dim Mnu As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
      Dim Mnu_Name As String = Mnu.Name

      Jrn_Add(, {Mnu_Name.Substring(8, 3)})
      Plcy_Strg = Mnu_Name.Substring(8, 3)



      Dim U_temp(80, 3) As String
      Array.Copy(U, U_temp, UNbCopy)

      Select Case Plcy_Strg
        Case "GLk" : Strategy_GLk(U_temp)
        Case "Gbl" : Strategy_Gbl(U_temp)
        Case "Gbv" : Strategy_Gbv(U_temp)
        Case "GCs" : Strategy_GCs(U_temp)
        Case "GCx" : Strategy_GCx(U_temp)
        Case "XCy" : Strategy_XCy(U_temp)
        Case "XRp" : Strategy_XRp(U_temp)
        Case "XNl" : Strategy_XNl(U_temp)
        Case "WgX" : Strategy_WgX(U_temp)
        Case "WgY" : Strategy_WgY(U_temp)
        Case "WgZ" : Strategy_WgZ(U_temp)
        Case "WgW" : Strategy_WgW(U_temp)
      End Select
      B_Famille.Text = Stg_Get(Plcy_Strg).Family.ToString()
      B_Info.Text = Stg_Get(Plcy_Strg).Texte
    End If

  End Sub


#End Region
End Class