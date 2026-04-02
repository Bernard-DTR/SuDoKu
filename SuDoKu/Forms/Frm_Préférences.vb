Public Class Frm_Préférences
  '-------------------------------------------------------------------------------
  'Le formulaire est constitué d'onglets
  'Il est ouvert à l'endroit où il a été fermé, comme il a été fermé, sauf à l'ouverture de l'application.
  'Le nom des contrôles est le suivant: ABCxx_yy
  '                                     ABC       Type d'objet CB CheckBox, Lbl Label,...
  '                                        xx     N° d'onglet
  '                                           yy  N° d'ordre de l'objet  
  'Préférences affiche l'ensemble des variables utilisées dans SDK
  '            Ces variables sont définies dans le module M00_Variables
  '            Elles sont initiées dans SDK.SDK_INI_Personnalisation_Dimension_Présentation_1&2()
  '            Elles sont modifiées dans chaque fonction du formulaire
  '            La modification concerne LA VARIABLE GéNéRALE ET LE FICHIER INI
  '-------------------------------------------------------------------------------
  ' Formulaire, Onglets, Généraux
  '-------------------------------------------------------------------------------
  ' V  28/03/2025 Mise en place du Nuancier Onglet_07
  '    le nuancier doit reprendre le traitement 
  Dim Mode_Load As Boolean = True
  Dim TrackBar01_02_Minimum As Integer
  Dim TrackBar01_02_Maximum As Integer

  Private Prf06_Pt_TopLeft As Point        ' Position HG de la grille 3x3
  Private Prf06_WH As Integer              ' Largeur Hauteur = 60
  Private Prf06_WH2 As Integer             ' 30
  Private Prf06_WH3 As Integer             ' 20
  Private Prf06_WH6 As Integer             ' 10
  Private Prf06_WH_Grid As Integer         ' Largeur hauteur de la,grille, traits compris
  Private Prf06_Trait_Pos_xy(4) As Integer ' Position des traits
  Private Prf06_Clr(9) As Color            ' Les couleurs
  Private Prf06_Clr_Txt(9) As String       ' le texte pour les couleurs
  Private Prf06_Clr_Name(9) As String      ' le nom des champs des couleurs
  Private Prf06_Sqr(9) As Rectangle        ' 9 Rectangles de la grille 3x3
  Private Prf06_Sqr_Cdd(89) As Rectangle   ' 90 rectangles des candidats

  Private Sub Préférences_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Mode_Load = True
    TrackBar01_02_Minimum = 40
    TrackBar01_02_Maximum = 100
    BackColor = Color_Frm_BackColor

    Left = Frm_SDK.Left + Frm_SDK.Journal.Left
    Top = Frm_SDK.Top + Frm_SDK.Journal.Top
    Icon = My.Resources.SuDoKu
    TopMost = True
    Text = Msg_Read("PRF_00000")

    'Onglet 01 Grille
    Onglet_01.Text = Msg_Read("PRF_01000")
    Onglet_01.ForeColor = Color.Red 'Sans effet
    Onglet_01.BackColor = Color_Fond_Typ_I
    Btn01_90.Text = "Reset " & Msg_Read("PRF_01000")
    TrackBar01_02.Value = WH
    Lbl01_01.Text = Msg_Read("PRF_01011", {CStr(WH), CStr(TrackBar01_02_Minimum), CStr(TrackBar01_02_Maximum)})
    TrackBar01_02.Minimum = TrackBar01_02_Minimum
    TrackBar01_02.Maximum = TrackBar01_02_Maximum

    Lbl01_06VI.Text = Msg_Read("PRF_01060")
    CB01_ColorComboboxVI.Cbb_Color_Selected = My.Settings.Prf_01C_Clr_ComboBoxVI
    CB01_ColorComboboxVCdd.Cbb_Color_Exclude = My.Settings.Prf_01C_Clr_ComboBoxVI
    Lbl01_06VCdd.Text = Msg_Read("PRF_01061")
    CB01_ColorComboboxVCdd.Cbb_Color_Selected = My.Settings.Prf_01C_Clr_ComboBoxVCdd
    CB01_ColorComboboxVI.Cbb_Color_Exclude = My.Settings.Prf_01C_Clr_ComboBoxVCdd

    ' CB01_ComboBoxFond.
    Lbl01_Fond.Text = Msg_Read("PRF_01070")

    Dim Répertoire As String = Path_SDK & "S10_Icônes\Fonds\"
    Dim Fond_Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(Répertoire)
                                               Where File.Contains("JPG") Or File.Contains("jpg")
                                               Order By File Descending
    CB01_ComboBoxFond.Items.Clear()
    Nsd_i = CB01_ComboBoxFond.Items.Add("*Standard")
    For Each File As String In Fond_Files
      Nsd_i = CB01_ComboBoxFond.Items.Add(File)
    Next File
    CB01_ComboBoxFond.SelectedIndex = Plcy_Fond_Grille
    CB01_11.Text = Msg_Read("PRF_01110")
    Select Case Plcy_MouseClick_Middle
      Case False : CB01_11.Checked = False
      Case True : CB01_11.Checked = True
    End Select

    Cb01_Format.Items.Clear()
    Nsd_i = Cb01_Format.Items.Add("Angles droits")          ' 0  D
    Nsd_i = Cb01_Format.Items.Add("4 Coins arrondis")       ' 1  A
    Nsd_i = Cb01_Format.Items.Add("4 Coins biseautés")      ' 2  B
    Nsd_i = Cb01_Format.Items.Add("Coins arrondis")         ' 3  A
    Nsd_i = Cb01_Format.Items.Add("Coins biseautés")        ' 4  B
    Nsd_i = Cb01_Format.Items.Add("Régions arrondies")      ' 5  A
    Nsd_i = Cb01_Format.Items.Add("Régions biseautées")     ' 6  B
    Nsd_i = Cb01_Format.Items.Add("... Au hasard")          ' 7
    Cb01_Format.SelectedIndex = My.Settings.Format_DAB
    Plcy_Format_DAB = My.Settings.Format_DAB
    If My.Settings.Format_DAB = 7 Then Plcy_Format_DAB = Rdc.Next(0, 6)

    CB01_Thèmes.Items.Clear()
    Nsd_i = CB01_Thèmes.Items.Add("Standard")               ' 0     
    Nsd_i = CB01_Thèmes.Items.Add("Océan")                  ' 1         
    Nsd_i = CB01_Thèmes.Items.Add("Contemporain")           ' 2
    CB01_Thèmes.SelectedIndex = My.Settings.Thème_Clr

    CB01_Police.Items.Clear()
    CB01_Police.Items.Add(Msg_Read("PRF_01201"))
    CB01_Police.Items.Add(Msg_Read("PRF_01202"))
    CB01_Police.Items.Add(Msg_Read("PRF_01203"))
    CB01_Police.Items.Add(Msg_Read("PRF_01204"))

    CB01_Police.SelectedIndex = 0

    'Onglet 02 Création
    Onglet_02.Text = Msg_Read("PRF_02000")
    Onglet_02.BackColor = Color_Fond_Typ_I

    Lbl02_01.Text = Msg_Read("PRF_02010")
    TB02_02.Text = Create_Nb_Cel_Demandées
    Btn02_90.Text = "Reset " & Msg_Read("PRF_02000")
    CB02_04.Text = Msg_Read("PRF_02020")
    Select Case Plcy_Generate_Batch
      Case False : CB02_04.Checked = False
      Case True : CB02_04.Checked = True
    End Select
    Lbl02_05.Text = Msg_Read("PRF_02030")
    TB02_06.Text = CStr(My.Settings.Prf_02C_Nb_Batch_Generate)
    Lbl02_10.Text = Msg_Read("PRF_02040")
    TB02_10.Text = Create_Nb_Tentatives
    CB02_10.Text = Msg_Read("PRF_02050")
    Select Case Create_Chat
      Case False : CB02_10.Checked = False
      Case True : CB02_10.Checked = True
    End Select
    CB02_Contraintes.Items.Clear()
    CB02_Contraintes.Items.Add(Msg_Read("PRF_02060"))
    CB02_Contraintes.Items.Add(Msg_Read("PRF_02061"))
    CB02_Contraintes.Items.Add(Msg_Read("PRF_02062"))
    CB02_Contraintes.Items.Add(Msg_Read("PRF_02063"))
    CB02_Contraintes.SelectedIndex = My.Settings.Prf_02C_Create_Contrainte
    Lbl02_07.Text = Msg_Read("PRF_02070")
    TB02_07.Text = CStr(My.Settings.Prf_02C_Nb_Max_Dl)

    'Onglet 03 Stratégie
    Onglet_03.Text = Msg_Read("PRF_03000")
    Onglet_03.BackColor = Color_Fond_Typ_I
    CLB03_Stratégies.Items.Clear()
    ' Seules les stratégies de production sont chargées
    For Each Stg As Stg_Cls In Stg_List
      If Stg.Prd = "O" Then Nsd_i = CLB03_Stratégies.Items.Add(Stg.Lettre & " " & Stg.Texte)
    Next Stg
    For i As Integer = 0 To Plcy_Stg_Clb.Count - 1
      Select Case Mid$(Plcy_Stg_Clb, i + 1, 1)
        Case "O" : CLB03_Stratégies.SetItemChecked(i, True)
        Case "N" : CLB03_Stratégies.SetItemChecked(i, False)
      End Select
    Next i

    'Onglet 05 Divers
    Onglet_05.Text = Msg_Read("PRF_05000")
    Onglet_05.BackColor = Color_Fond_Typ_I
    Btn05_90.Text = "Reset " & Msg_Read("PRF_05000")
    'CB05_08.Text = Msg_Read("PRF_05080")
    'Select Case Plcy_Gbl_Etendue
    '  Case True : CB05_08.Checked = True
    '  Case False : CB05_08.Checked = False
    'End Select
    CB05_10.Text = Msg_Read("PRF_05100")
    Select Case Plcy_Dancing_Link
      Case True : CB05_10.Checked = True
      Case False : CB05_10.Checked = False
    End Select
    CB05_11.Text = Msg_Read("PRF_05110")
    Select Case Plcy_Open_Display
      Case True : CB05_11.Checked = True
      Case False : CB05_11.Checked = False
    End Select
    CB05_12.Text = Msg_Read("PRF_05120")
    Select Case Plcy_AfficherDCdd_Bande
      Case True : CB05_12.Checked = True
      Case False : CB05_12.Checked = False
    End Select

    'Onglet 06 Couleurs
    Onglet_06.Text = Msg_Read("PRF_06000")
    Onglet_06.BackColor = Color_Fond_Typ_I
    'Le traitement Paint de Préférences diffère du traitement OnPaint de SDK
    AddHandler Onglet_06.Paint, AddressOf Me.Onglet_06_Paint
    Prf06_Clr_Compute()
    Prf06_Clr_Initialisation()

    Onglet_TC.SelectTab(My.Settings.Prf_00_Onglet_Sélecté)
    Mode_Load = False

    ' Onglet Nuancier
    Onglet_07.BackColor = Color_Fond_Typ_I
    With DGV07_Color
      .Location = New Point(5, 5)
      .ColumnHeadersVisible = False
      .RowHeadersVisible = False
      .ColumnCount = 2
      .RowCount = 4
      .Columns(0).Width = 2 * WH
      .Columns(1).Width = 3 * WH
      For r As Integer = 0 To 3
        .Rows(r).Height = WH \ 2
      Next r
      'AllowUserToAddRows = False ne doit pas être utilisé
      .AllowUserToDeleteRows = False
      .AllowUserToOrderColumns = False
      .AllowUserToResizeColumns = False
      .AllowUserToResizeRows = False

      .Width = 5 * WH
      .Height = 4 * (WH \ 2) + 5
      .ScrollBars = ScrollBars.None
      .BorderStyle = BorderStyle.Fixed3D

      .SelectionMode = DataGridViewSelectionMode.CellSelect
      .MultiSelect = False
      .DefaultCellStyle.Font = New Font("Tahoma", 10)
      .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      DGV_Color_Display()
    End With

    Onglet_08.BackColor = Color_Fond_Typ_I
    CB08_01.Text = Msg_Read("PRF_08010")
    Select Case Xap
      Case True : CB08_01.Checked = True
      Case False : CB08_01.Checked = False
    End Select
    CBB08_03.SelectedIndex = My.Settings.Prf_08H_Flèche

  End Sub

  Private Sub Onglet_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Onglet_TC.SelectedIndexChanged
    'Mémorise le dernier onglet sélectionné
    My.Settings.Prf_00_Onglet_Sélecté = Onglet_TC.SelectedIndex
  End Sub

  Private Sub Onglet_06_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs)
    ' Construction des contrôles de l'Onglet_06

    Using pen_1 As New Pen(Prf06_Clr(2), 1),
          pen_3 As New Pen(Prf06_Clr(2), 3),
          brsh_1 As New SolidBrush(Prf06_Clr(1)),
          brsh_3 As New SolidBrush(Prf06_Clr(3)),
          brsh_4 As New SolidBrush(Prf06_Clr(4)),
          brsh_5 As New SolidBrush(Prf06_Clr(5)),
          brsh_6 As New SolidBrush(Prf06_Clr(6)),
          brsh_7 As New SolidBrush(Prf06_Clr(7)),
          brsh_8 As New SolidBrush(Prf06_Clr(8)),
          font_30 As New Font(Font_Name_ValCdd, 30),
          font_10 As New Font(Font_Name_ValCdd, 10)

      ' 1 Efface l'intégralité de l'emplacement de la grille avec un carré unique
      e.Graphics.FillRectangle(brsh_1,
                    Prf06_Pt_TopLeft.X - 10, Prf06_Pt_TopLeft.Y - 10, Prf06_WH_Grid + 20, Prf06_WH_Grid + 20)

      ' 2 Traçage des traits et de l'entourage
      Dim Pt_H1 As Point, Pt_H2 As Point, Pt_V1 As Point, Pt_V2 As Point
      For i As Integer = 0 To 4
        Pt_H1 = New Point(x:=Prf06_Pt_TopLeft.X,
                          y:=Prf06_Pt_TopLeft.Y + Prf06_Trait_Pos_xy(i))
        Pt_H2 = New Point(x:=Pt_H1.X + Prf06_WH_Grid - 1,
                          y:=Pt_H1.Y)
        Pt_V1 = New Point(x:=Prf06_Pt_TopLeft.X + Prf06_Trait_Pos_xy(i),
                          y:=Prf06_Pt_TopLeft.Y)
        Pt_V2 = New Point(x:=Pt_V1.X,
                          y:=Pt_V1.Y + Prf06_WH_Grid - 1)
        Select Case i
          Case 0, 3        '  Trait de 3
            e.Graphics.DrawLine(pen_3, Pt_H1, Pt_H2)
            e.Graphics.DrawLine(pen_3, Pt_V1, Pt_V2)
          Case 1, 2 '  Trait de 1
            e.Graphics.DrawLine(pen_1, Pt_H1, Pt_H2)
            e.Graphics.DrawLine(pen_1, Pt_V1, Pt_V2)
        End Select
      Next i

      ' 3 Fonds des 9 cellules Les couleurs sont arbitraires
      e.Graphics.FillRectangle(brsh_3, Prf06_Sqr(0))
      e.Graphics.FillRectangle(brsh_5, Prf06_Sqr(1))
      e.Graphics.FillRectangle(brsh_5, Prf06_Sqr(2))
      e.Graphics.FillRectangle(brsh_5, Prf06_Sqr(3))
      e.Graphics.FillRectangle(brsh_3, Prf06_Sqr(4))
      e.Graphics.FillRectangle(brsh_5, Prf06_Sqr(5))
      e.Graphics.FillRectangle(brsh_5, Prf06_Sqr(6))
      e.Graphics.FillRectangle(brsh_5, Prf06_Sqr(7))
      e.Graphics.FillRectangle(brsh_3, Prf06_Sqr(8))

      ' 4 Valeurs dans Cellules Initiales et dans Cellules Remplies 
      e.Graphics.DrawString("9", font_30, brsh_4, Prf06_Sqr(0).X + Prf06_WH2, Prf06_Sqr(0).Y + Prf06_WH2, Format_Center)
      e.Graphics.DrawString("1", font_30, brsh_4, Prf06_Sqr(4).X + Prf06_WH2, Prf06_Sqr(4).Y + Prf06_WH2, Format_Center)
      e.Graphics.DrawString("7", font_30, brsh_4, Prf06_Sqr(8).X + Prf06_WH2, Prf06_Sqr(8).Y + Prf06_WH2, Format_Center)
      e.Graphics.DrawString("2", font_30, brsh_6, Prf06_Sqr(2).X + Prf06_WH2, Prf06_Sqr(2).Y + Prf06_WH2, Format_Center)
      e.Graphics.DrawString("5", font_30, brsh_6, Prf06_Sqr(5).X + Prf06_WH2, Prf06_Sqr(5).Y + Prf06_WH2, Format_Center)

      ' 5 Candidats
      e.Graphics.DrawString("3", font_10, brsh_6, Prf06_Sqr_Cdd(63).X + Prf06_WH6, Prf06_Sqr_Cdd(63).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("4", font_10, brsh_6, Prf06_Sqr_Cdd(64).X + Prf06_WH6, Prf06_Sqr_Cdd(64).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("6", font_10, brsh_6, Prf06_Sqr_Cdd(66).X + Prf06_WH6, Prf06_Sqr_Cdd(66).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("8", font_10, brsh_6, Prf06_Sqr_Cdd(68).X + Prf06_WH6, Prf06_Sqr_Cdd(68).Y + Prf06_WH6, Format_Center)

      ' 6 Aide Graphique
      '   Uniquement pour les cellules 0, 1, 2
      e.Graphics.FillRectangle(brsh_7, Prf06_Sqr(0).X, Prf06_Sqr(0).Y + Prf06_WH3, (Prf06_WH * 3) + 4, Prf06_WH3)

      ' 7 la Sélection    
      e.Graphics.DrawString("1", font_10, brsh_6, Prf06_Sqr_Cdd(71).X + Prf06_WH6, Prf06_Sqr_Cdd(71).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("2", font_10, brsh_6, Prf06_Sqr_Cdd(72).X + Prf06_WH6, Prf06_Sqr_Cdd(72).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("3", font_10, brsh_6, Prf06_Sqr_Cdd(73).X + Prf06_WH6, Prf06_Sqr_Cdd(73).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("4", font_10, brsh_6, Prf06_Sqr_Cdd(74).X + Prf06_WH6, Prf06_Sqr_Cdd(74).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("5", font_10, brsh_6, Prf06_Sqr_Cdd(75).X + Prf06_WH6, Prf06_Sqr_Cdd(75).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("6", font_10, brsh_6, Prf06_Sqr_Cdd(76).X + Prf06_WH6, Prf06_Sqr_Cdd(76).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("7", font_10, brsh_6, Prf06_Sqr_Cdd(77).X + Prf06_WH6, Prf06_Sqr_Cdd(77).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("8", font_10, brsh_6, Prf06_Sqr_Cdd(78).X + Prf06_WH6, Prf06_Sqr_Cdd(78).Y + Prf06_WH6, Format_Center)
      e.Graphics.DrawString("9", font_10, brsh_6, Prf06_Sqr_Cdd(79).X + Prf06_WH6, Prf06_Sqr_Cdd(79).Y + Prf06_WH6, Format_Center)
      e.Graphics.FillRectangle(brsh_8, Prf06_Sqr(7))
    End Using
  End Sub
  '-------------------------------------------------------------------------------
  ' Contrôles par onglet et ordre de contrôle
  '-------------------------------------------------------------------------------
  Private Sub TrackBar01_02_ValueChanged(sender As Object, e As EventArgs) Handles TrackBar01_02.ValueChanged
    'Modification de la taille de la grille de TrackBar01_02_Minimum à TrackBar01_02_Maximum
    'Saisie numérique dans la TrackBar de la taille WH
    WH = TrackBar01_02.Value
    Lbl01_01.Text = Msg_Read("PRF_01011", {CStr(WH), CStr(TrackBar01_02_Minimum), CStr(TrackBar01_02_Maximum)})
    My.Settings.Prf_01C_Taille_Cellule = WH
    If Not Mode_Load Then
      OC_Présentation()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub
  Private Sub CB01_ColorComboboxVI_SelectedColorChanged(SelectedColor As Color, sender As Object) Handles CB01_ColorComboboxVI.Cbb_Color_Changed_Selected
    CB01_ColorComboboxVCdd.Cbb_Color_Exclude = CB01_ColorComboboxVI.Cbb_Color_Selected
    My.Settings.Prf_01C_Clr_ComboBoxVI = CB01_ColorComboboxVI.Cbb_Color_Selected
    Color_VI = Color.FromName(CB01_ColorComboboxVI.Cbb_Color_Selected)
    If Not Mode_Load Then
      'Pour prendre en compte la nouvelle couleur
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub
  Private Sub CB01_ColorComboboxVCdd_SelectedColorChanged(SelectedColor As Color, sender As Object) Handles CB01_ColorComboboxVCdd.Cbb_Color_Changed_Selected
    CB01_ColorComboboxVI.Cbb_Color_Exclude = CB01_ColorComboboxVCdd.Cbb_Color_Selected
    My.Settings.Prf_01C_Clr_ComboBoxVCdd = CB01_ColorComboboxVCdd.Cbb_Color_Selected
    Color_VCdd = Color.FromName(CB01_ColorComboboxVCdd.Cbb_Color_Selected)
    If Not Mode_Load Then
      'Pour prendre en compte la nouvelle couleur
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub
  Private Sub CB01_ComboBoxFond_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB01_ComboBoxFond.SelectedIndexChanged
    My.Settings.Prf_01C_Fond_Grille = CB01_ComboBoxFond.SelectedIndex
    Plcy_Fond_Grille = CB01_ComboBoxFond.SelectedIndex
    If Not Mode_Load Then
      ' Recalculer Sqr_Img et Font_Style
      OC_Présentation()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub

  Private Sub CB01_11_CheckedChanged(sender As Object, e As EventArgs) Handles CB01_11.CheckedChanged
    'Utilisation du Bouton central de la souris
    Select Case CB01_11.CheckState
      Case CheckState.Unchecked  '0   
        Plcy_MouseClick_Middle = False
        My.Settings.Prf_01C_MouseClick_Middle = False
      Case CheckState.Checked '1      
        Plcy_MouseClick_Middle = True
        My.Settings.Prf_01C_MouseClick_Middle = True
    End Select
  End Sub
  Private Sub CB01_Format_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cb01_Format.SelectedIndexChanged
    If Not Mode_Load Then
      My.Settings.Format_DAB = Cb01_Format.SelectedIndex
      Plcy_Format_DAB = My.Settings.Format_DAB
      If My.Settings.Format_DAB = 7 Then Plcy_Format_DAB = Rdc.Next(0, 6)
      OC_Présentation()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub

  Private Sub CB01_Thèmes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB01_Thèmes.SelectedIndexChanged
    If Not Mode_Load Then
      My.Settings.Thème_Clr = CB01_Thèmes.SelectedIndex
      OC_Présentation()
      Event_OnPaint = "Total_Frm_Préférences"
      Frm_SDK.Invalidate()
      OC_Thèmes_Couleurs(My.Settings.Thème_Clr)
      BackColor = Color_Frm_BackColor
      For Each Page As TabPage In Onglet_TC.TabPages
        Page.BackColor = Color_Fond_Typ_I
      Next Page
      Frm_SDK.Journal.BackColor = Color_Fond_Typ_I
    End If
  End Sub

  Private Sub CB01_Polices_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB01_Police.SelectedIndexChanged
    If Not Mode_Load Then
      Plcy_Fantasy_Name = CB01_Police.SelectedItem.ToString()
      Select Case Plcy_Fantasy_Name
        Case "Arial"
          Plcy_Fantasy = False
          Font_Name_ValCdd = "Arial"

        Case "MS Outlook"
          Plcy_Fantasy = True
          Font_Name_ValCdd = "MS Outlook"

        Case "Webdings"
          Plcy_Fantasy = True
          Font_Name_ValCdd = "Webdings"

        Case "Wingdings"
          Plcy_Fantasy = True
          Font_Name_ValCdd = "Wingdings"

      End Select
      OC_Présentation()
      Mnu_Mngt_Barre_Outils_Filtres()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub
  Private Sub TB02_02_Validated(sender As Object, e As EventArgs) Handles TB02_02.Validated
    'Crt Nombre de Cellules demandées, ce nombre ne peut excéder 81
    If TB02_02.Text >= "81" Then
      Jrn_Add("PRF_02009", Nothing, "Erreur")
      TB02_02.Text = "81"
    End If
    Create_Nb_Cel_Demandées = TB02_02.Text
    My.Settings.Prf_02C_Create_Nb_Cel_Demandées = CInt(TB02_02.Text)
  End Sub

  Private Sub TB02_10_Validated(sender As Object, e As EventArgs) Handles TB02_10.Validated
    'Nombre de Tentatives
    If CInt(TB02_10.Text) >= 25 Then
      Jrn_Add("PRF_02049", Nothing, "Erreur")
      TB02_10.Text = "5"
    End If
    Create_Nb_Tentatives = TB02_10.Text
    My.Settings.Prf_02C_Create_Nb_Tentatives = CInt(TB02_10.Text)
  End Sub

  Private Sub CB02_Contraintes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB02_Contraintes.SelectedIndexChanged
    Create_Contrainte_Originale = CB02_Contraintes.SelectedIndex
    My.Settings.Prf_02C_Create_Contrainte = CB02_Contraintes.SelectedIndex
  End Sub

  Private Sub CB02_04_CheckedChanged(sender As Object, e As EventArgs) Handles CB02_04.CheckedChanged
    Select Case CB02_04.CheckState
      Case CheckState.Unchecked '0  'Non
        Plcy_Generate_Batch = False
        My.Settings.Prf_02C_Plcy_Batch_Generate = False
      Case CheckState.Checked '1  'Oui
        Plcy_Generate_Batch = True
        My.Settings.Prf_02C_Plcy_Batch_Generate = True
    End Select
  End Sub
  Private Sub CB02_10_CheckedChanged(sender As Object, e As EventArgs) Handles CB02_10.CheckedChanged
    'Mode Bavard-Silencieux
    Select Case CB02_10.CheckState
      Case CheckState.Unchecked '0  'Non
        Create_Chat = False
        My.Settings.Prf_02C_Create_Chat = False
      Case CheckState.Checked '1  'Oui
        Create_Chat = True
        My.Settings.Prf_02C_Create_Chat = True
    End Select
  End Sub
  Private Sub TB02_06_TextChanged(sender As Object, e As EventArgs) Handles TB02_06.TextChanged
    My.Settings.Prf_02C_Nb_Batch_Generate = CInt(TB02_06.Text)
  End Sub
  Private Sub TB02_07_TextChanged(sender As Object, e As EventArgs) Handles TB02_07.TextChanged
    My.Settings.Prf_02C_Nb_Max_Dl = CInt(TB02_07.Text)
  End Sub

  Private Sub CB05_10_CheckedChanged(sender As Object, e As EventArgs) Handles CB05_10.CheckedChanged
    'Exécution systématiqhe de la Vérification Dancing_Link
    Select Case CB05_10.CheckState
      Case CheckState.Unchecked '0  'Non
        Plcy_Dancing_Link = False
        My.Settings.Prf_05D_Plcy_Dancing_Link = False
      Case CheckState.Checked '1  'Oui
        Plcy_Dancing_Link = True
        My.Settings.Prf_05D_Plcy_Dancing_Link = True
    End Select
    If Not Mode_Load Then
      OC_Présentation()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub

  Private Sub CB05_11_CheckedChanged(sender As Object, e As EventArgs) Handles CB05_11.CheckedChanged
    ' Affichage d'un fichier Puzzle lors de son ouverture
    Select Case CB05_11.CheckState
      Case CheckState.Unchecked '0  'Non
        Plcy_Open_Display = False
        My.Settings.Prf_05D_Plcy_Open_Display = False
      Case CheckState.Checked '1  'Oui
        Plcy_Open_Display = True
        My.Settings.Prf_05D_Plcy_Open_Display = True
    End Select
    If Not Mode_Load Then
      OC_Présentation()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub

  Private Sub CB05_12_CheckedChanged(sender As Object, e As EventArgs) Handles CB05_12.CheckedChanged
    ' Affichage du Dernier Candidat dans la bande de la Cellule
    Select Case CB05_12.CheckState
      Case CheckState.Unchecked '0  'Non
        Plcy_AfficherDCdd_Bande = False
        My.Settings.Prf_05D_Plcy_AfficherDCdd_Bande = False
      Case CheckState.Checked '1  'Oui
        Plcy_AfficherDCdd_Bande = True
        My.Settings.Prf_05D_Plcy_AfficherDCdd_Bande = True
    End Select
    If Not Mode_Load Then
      OC_Présentation()
      Event_OnPaint = "Global"
      Frm_SDK.Invalidate()
      Application.DoEvents()
    End If
  End Sub
  Private Sub Btn03_91_Click(sender As Object, e As EventArgs) Handles Btn03_91.Click
    'Enlever toutes les sélections
    For i As Integer = 0 To Plcy_Stg_Clb.Count - 1
      CLB03_Stratégies.SetItemChecked(i, False)
    Next i
  End Sub
  Private Sub Btn03_92_Click(sender As Object, e As EventArgs) Handles Btn03_92.Click
    'Tout sélectionner
    For i As Integer = 0 To Plcy_Stg_Clb.Count - 1
      CLB03_Stratégies.SetItemChecked(i, True)
    Next i
  End Sub
  Private Sub Btn03_93_Click(sender As Object, e As EventArgs) Handles Btn03_93.Click
    Plcy_Stg_Clb = StrDup(CLB03_Stratégies.Items.Count, "_")
    For i As Integer = 0 To Plcy_Stg_Clb.Count - 1
      Select Case CLB03_Stratégies.GetItemCheckState(i).ToString()
        Case "Checked" : Mid$(Plcy_Stg_Clb, i + 1, 1) = "O"
        Case "Unchecked" : Mid$(Plcy_Stg_Clb, i + 1, 1) = "N"
      End Select
    Next i
    My.Settings.Prf_03R_Plcy_Stg_Clb = Plcy_Stg_Clb
    OC_Plcy_Stg_UOBTXYSJZKQ()
    ' Affichage de contrôle
    Jrn_Add("Prl_00000", {"Stratégies         : "})
    Jrn_Add("Prl_00000", {"Plcy_Stg_Clb       : " & Plcy_Stg_Clb})
    Jrn_Add("Prl_00000", {"Stg_Profondeur     : " & Stg_Profondeur})

    Me.Close()
  End Sub
  Private Sub Reset(sender As Object, e As EventArgs) Handles Btn05_90.Click, Btn02_90.Click, Btn01_90.Click
    Select Case Onglet_TC.SelectedIndex ' Quel est l'onglet affecté ?
      Case 0 ' Grille
        '           Quelle est la taille de la grille ? 
        WH = CInt(Ini_Read("VU", "WH"))
        My.Settings.Prf_01C_Taille_Cellule = WH
        OC_Présentation()
        '           Quelle est la couleur des valeurs et des candidats ? Noire
        '           Quelle est la couleur des valeurs initiales        ? Rouge
        My.Settings.Prf_01C_Clr_ComboBoxVCdd = "Black"
        My.Settings.Prf_01C_Clr_ComboBoxVI = "Red"
        Color_VI = Color.FromName("Red")
        Color_VCdd = Color.FromName("Black")
        '           Affichage d'un fond de Grille ? Non
        My.Settings.Prf_01C_Fond_Grille = 0
        Plcy_Fond_Grille = 0
        Plcy_MouseClick_Middle = False
      Case 1 ' Création
        My.Settings.Prf_02C_Create_Chat = False
        My.Settings.Prf_02C_Create_Nb_Cel_Demandées = 25
        My.Settings.Prf_02C_Create_Nb_Tentatives = 10
        Create_Chat = My.Settings.Prf_02C_Create_Chat
        Create_Nb_Cel_Demandées = CStr(My.Settings.Prf_02C_Create_Nb_Cel_Demandées)
        Create_Nb_Tentatives = CStr(My.Settings.Prf_02C_Create_Nb_Tentatives)
        My.Settings.Prf_02C_Nb_Max_Dl = 15
      Case 2 ' Résolution
      Case 3 ' Simulation
      Case 4 ' Divers
        My.Settings.Prf_05D_Plcy_Dancing_Link = True

      Case Else
    End Select
    Préférences_Load(sender, e)
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub
  'Saisie numérique dans les textBox
  Private Sub TB02_02_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TB02_02.KeyPress
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub
  Private Sub TB05_03_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub
  Private Sub TB03_04_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub
  Private Sub TB04_05_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub
  Private Sub TB02_06_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub
  Private Sub TB04_07_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub
  Private Sub TB04_10_KeyPress(sender As Object, e As KeyPressEventArgs)
    If Char.IsNumber(e.KeyChar) = False Then e.Handled = True
  End Sub

  Public Sub Prf06_Clr_Compute()
    ' 1 Définition
    Prf06_Pt_TopLeft.X = 15
    Prf06_Pt_TopLeft.Y = 15
    Prf06_WH = 60
    Prf06_WH2 = 30
    Prf06_WH3 = 20
    Prf06_WH6 = 10
    Prf06_WH_Grid = (Prf06_WH * 3) + Bld_Trait_3 + Bld_Trait_1 + Bld_Trait_1 + Bld_Trait_3

    ' 2 Définition des traits Prf06_Trait_Pos_xy()
    Dim Ep_2 As Integer = 2 ' = 2   car le trait a une épaisseur de 3
    Dim Ep_1 As Integer = 1 ' = 1   car le trait a une épaisseur de 1

    Prf06_Trait_Pos_xy(0) = Ep_1
    Prf06_Trait_Pos_xy(1) = Ep_2 + Prf06_WH + Prf06_Trait_Pos_xy(0)
    Prf06_Trait_Pos_xy(2) = Ep_1 + Prf06_WH + Prf06_Trait_Pos_xy(1)
    Prf06_Trait_Pos_xy(3) = Ep_2 + Prf06_WH + Prf06_Trait_Pos_xy(2)

    ' 3 Calcul des 9 cellules
    Dim v, h, k As Integer
    For j As Integer = 0 To 2      ' Axe des y, vertical
      Select Case j
        Case 0
          v = 2
        Case 1, 2
          v = 1
      End Select
      For i As Integer = 0 To 2  ' Axe des x, horizontal
        Select Case i
          Case 0
            h = 2
          Case 1, 2
            h = 1
        End Select
        Prf06_Sqr(k) = New Rectangle(x:=Prf06_Pt_TopLeft.X + Prf06_Trait_Pos_xy(i) + h,
                                     y:=Prf06_Pt_TopLeft.Y + Prf06_Trait_Pos_xy(j) + v, width:=Prf06_WH, height:=Prf06_WH)
        k += 1
      Next i
    Next j

    ' 4 Tableau des Prf06_Clr
    Prf06_Clr(0) = Color_Frm_BackColor
    Prf06_Clr(1) = Color_Frm_BackColor
    Prf06_Clr(2) = Color_Trait
    Prf06_Clr(3) = Color_Fond_Typ_I
    Prf06_Clr(4) = Color_VI
    Prf06_Clr(5) = Color_Fond_Typ_RV
    Prf06_Clr(6) = Color_VCdd
    Prf06_Clr(7) = Color_Stratégique
    Prf06_Clr(8) = Color_Cell_Select

    ' 5 Calcul des 81 cellules des candidats
    k = 0
    For j As Integer = 0 To 2      ' Axe des y, vertical
      For i As Integer = 0 To 2  ' Axe des x, horizontal
        Prf06_Sqr_Cdd((k * 10) + 1) = New Rectangle(x:=Prf06_Sqr(k).X + (0 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (0 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 1
        Prf06_Sqr_Cdd((k * 10) + 2) = New Rectangle(x:=Prf06_Sqr(k).X + (1 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (0 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 2
        Prf06_Sqr_Cdd((k * 10) + 3) = New Rectangle(x:=Prf06_Sqr(k).X + (2 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (0 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 3
        Prf06_Sqr_Cdd((k * 10) + 4) = New Rectangle(x:=Prf06_Sqr(k).X + (0 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (1 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 4
        Prf06_Sqr_Cdd((k * 10) + 5) = New Rectangle(x:=Prf06_Sqr(k).X + (1 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (1 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 5
        Prf06_Sqr_Cdd((k * 10) + 6) = New Rectangle(x:=Prf06_Sqr(k).X + (2 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (1 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 6
        Prf06_Sqr_Cdd((k * 10) + 7) = New Rectangle(x:=Prf06_Sqr(k).X + (0 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (2 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 7
        Prf06_Sqr_Cdd((k * 10) + 8) = New Rectangle(x:=Prf06_Sqr(k).X + (1 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (2 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 8
        Prf06_Sqr_Cdd((k * 10) + 9) = New Rectangle(x:=Prf06_Sqr(k).X + (2 * Prf06_WH3), y:=Prf06_Sqr(k).Y + (2 * Prf06_WH3), width:=Prf06_WH3, height:=Prf06_WH3)     ' Cdd 9
        k += 1
      Next i
    Next j

  End Sub
  Public Sub Prf06_Clr_Initialisation()
    'Le texte des boutons et des tooltiptext 
    Prf06_Clr_Txt(1) = Msg_Read("PRF_06011")
    Prf06_Clr_Txt(2) = Msg_Read("PRF_06012")
    Prf06_Clr_Txt(3) = Msg_Read("PRF_06013")
    Prf06_Clr_Txt(4) = Msg_Read("PRF_06014")
    Prf06_Clr_Txt(5) = Msg_Read("PRF_06015")
    Prf06_Clr_Txt(6) = Msg_Read("PRF_06016")
    Prf06_Clr_Txt(7) = Msg_Read("PRF_06017")
    Prf06_Clr_Txt(8) = Msg_Read("PRF_06018")

    Btn06_01.Text = Prf06_Clr_Txt(1)
    Btn06_02.Text = Prf06_Clr_Txt(2)
    Btn06_03.Text = Prf06_Clr_Txt(3)
    Btn06_04.Text = Prf06_Clr_Txt(4)
    Btn06_05.Text = Prf06_Clr_Txt(5)
    Btn06_06.Text = Prf06_Clr_Txt(6)
    Btn06_07.Text = Prf06_Clr_Txt(7)
    Btn06_08.Text = Prf06_Clr_Txt(8)
    'Le ToolTip affiche le texte, quand celui-ci est trop foncé
    ToolTip.SetToolTip(Btn06_01, Prf06_Clr_Txt(1))
    ToolTip.SetToolTip(Btn06_02, Prf06_Clr_Txt(2))
    ToolTip.SetToolTip(Btn06_03, Prf06_Clr_Txt(3))
    ToolTip.SetToolTip(Btn06_04, Prf06_Clr_Txt(4))
    ToolTip.SetToolTip(Btn06_05, Prf06_Clr_Txt(5))
    ToolTip.SetToolTip(Btn06_06, Prf06_Clr_Txt(6))
    ToolTip.SetToolTip(Btn06_07, Prf06_Clr_Txt(7))
    ToolTip.SetToolTip(Btn06_08, Prf06_Clr_Txt(8))
    'Le backColor des boutons
    Btn06_01.BackColor = Prf06_Clr(1)
    Btn06_02.BackColor = Prf06_Clr(2)
    Btn06_03.BackColor = Prf06_Clr(3)
    Btn06_04.BackColor = Prf06_Clr(4)
    Btn06_05.BackColor = Prf06_Clr(5)
    Btn06_06.BackColor = Prf06_Clr(6)
    Btn06_07.BackColor = Prf06_Clr(7)
    Btn06_08.BackColor = Prf06_Clr(8)

  End Sub

  Private Sub Btn06_Color(sender As Object, e As EventArgs) Handles Btn06_08.Click, Btn06_07.Click, Btn06_06.Click, Btn06_05.Click, Btn06_04.Click, Btn06_03.Click, Btn06_02.Click, Btn06_01.Click
    'Gestion des couleurs
    'La couleur du bouton est affichée
    Dim Clr As Integer = CInt(sender.ToString().Substring(35, 1))
    Prf06_Color_Show(Clr)
  End Sub

  Private Sub Prf06_Color_Show(Clr As Integer)
    Dim Prf_Clr_Int(9) As Integer
    'L'utilisation de SDK_ColorDialog hérité de ColorDialog permet de modifier le titre
    Dim SDK_Clr As New SDK_ColorDialog
    With SDK_Clr
      .Title = "Couleurs: " & Prf06_Clr_Txt(Clr)
      .Color = Prf06_Clr(Clr)
      .ShowHelp = False
      .SolidColorOnly = False
      .AnyColor = False
      .AllowFullOpen = True
      .FullOpen = True

      'Toutes les couleurs utilisées sont affichées
      For i As Integer = 0 To 8
        Prf_Clr_Int(i) = ColorTranslator.ToWin32(Prf06_Clr(i))
      Next i
      .CustomColors = Prf_Clr_Int

      'Affichage de la ColorDialog
      If .ShowDialog() = DialogResult.OK Then
        Dim A As Integer = 255
        If Clr = 8 Then A = 128
        Prf06_Clr(Clr) = Color.FromArgb(A, .Color.R, .Color.G, .Color.B)
      End If
    End With
    Prf06_Clr_Initialisation()
    Onglet_06.Refresh()
  End Sub

  Private Sub Btn06_Appliquer_Click(sender As Object, e As EventArgs) Handles Btn_Appliquer.Click
    Color_Frm_BackColor = Prf06_Clr(0)
    Color_Frm_BackColor = Prf06_Clr(1)
    Color_Trait = Prf06_Clr(2)
    Color_Fond_Typ_I = Prf06_Clr(3)
    Color_VI = Prf06_Clr(4)
    Color_Fond_Typ_RV = Prf06_Clr(5)
    Color_VCdd = Prf06_Clr(6)
    Color_Stratégique = Prf06_Clr(7)
    Color_Cell_Select = Prf06_Clr(8)

    Prf06_Clr_Name(1) = "Color_Frm_BackColor"
    Prf06_Clr_Name(2) = "Color_Trait"
    Prf06_Clr_Name(3) = "Color_Fond_Typ_I "
    Prf06_Clr_Name(4) = "Color_VI"
    Prf06_Clr_Name(5) = "Color_Fond_Typ_RV"
    Prf06_Clr_Name(6) = "Color_VCdd"
    Prf06_Clr_Name(7) = "Color_Stratégique"
    Prf06_Clr_Name(8) = "Color_Cell_Select"

    Jrn_Add(, {"Thèmes : " & CB01_Thèmes.SelectedIndex})
    For i As Integer = 1 To 8
      Jrn_Add(, {Prf06_Clr_Txt(i).PadRight(30) &
                 Prf06_Clr_Name(i).PadRight(30) &
         " A " & Prf06_Clr(i).A.ToString().PadLeft(3) &
         " R " & Prf06_Clr(i).R.ToString().PadLeft(3) &
         " G " & Prf06_Clr(i).G.ToString().PadLeft(3) &
         " B " & Prf06_Clr(i).B.ToString().PadLeft(3)})
    Next i
    OC_Présentation()
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()
    Application.DoEvents()
  End Sub

  Private Sub Onglet_06_MouseClick(sender As Object, e As MouseEventArgs) Handles Onglet_06.MouseClick
    Dim Pt As New Point(x:=e.X, y:=e.Y)
    If Prf06_Sqr(0).Contains(Pt) Then Prf06_Color_Show(7)
    If Prf06_Sqr(1).Contains(Pt) Then Prf06_Color_Show(7)
    If Prf06_Sqr(2).Contains(Pt) Then Prf06_Color_Show(3)
    If Prf06_Sqr(3).Contains(Pt) Then Prf06_Color_Show(5)
    If Prf06_Sqr(4).Contains(Pt) Then Prf06_Color_Show(4)
    If Prf06_Sqr(5).Contains(Pt) Then Prf06_Color_Show(6)
    If Prf06_Sqr(6).Contains(Pt) Then Prf06_Color_Show(6)
    If Prf06_Sqr(7).Contains(Pt) Then Prf06_Color_Show(8)
    If Prf06_Sqr(8).Contains(Pt) Then Prf06_Color_Show(4)
  End Sub

  Public Sub DGV_Color_Display()
    With DGV07_Color
      .Item(columnIndex:=0, rowIndex:=0).Value = Color_List(0).Symbol
      .Item(columnIndex:=0, rowIndex:=1).Value = Color_List(1).Symbol
      .Item(columnIndex:=0, rowIndex:=2).Value = Color_List(2).Symbol
      .Item(columnIndex:=0, rowIndex:=3).Value = Color_List(3).Symbol
      .Item(columnIndex:=1, rowIndex:=0).Style.BackColor = Color_List(0).Color
      .Item(columnIndex:=1, rowIndex:=1).Style.BackColor = Color_List(1).Color
      .Item(columnIndex:=1, rowIndex:=2).Style.BackColor = Color_List(2).Color
      .Item(columnIndex:=1, rowIndex:=3).Style.BackColor = Color_List(3).Color
    End With
    DGV07_Color.Refresh()
  End Sub

  Private Sub Btn1_Nuancier_Click(sender As Object, e As EventArgs) Handles Btn07_01.Click
    ' Modification des couleurs de Coloriage 
    Dim Lcg As Integer = DGV07_Color.CurrentCell.RowIndex ' Ligne de la couleur à modifier
    Dim ClrDlg As New SDK_ColorDialog
    With ClrDlg
      .Title = "Modification de la Couleur: " & Color_List(Lcg).Symbol
      .Color = Color_List(Lcg).Color
      .ShowHelp = False
      .SolidColorOnly = False
      .AnyColor = False
      .AllowFullOpen = True
      .FullOpen = True
      'Toutes les couleurs utilisées sont affichées
      Dim Clr_Int(3) As Integer
      For i As Integer = 0 To 3 : Clr_Int(i) = ColorTranslator.ToWin32(Color_List(i).Color) : Next i
      .CustomColors = Clr_Int

      If .ShowDialog() = DialogResult.OK Then          ' Réponse OK à l'Affichage de la ColorDialog
        Color_List(Lcg).Color = .Color                 ' Modification de la couleur
        DGV07_Color.Item(columnIndex:=1, rowIndex:=Lcg).Style.BackColor = Color_List(Lcg).Color
        Obj_Colors_Save()                              ' Entregistrement des couleurs
      End If
    End With
  End Sub

  Private Sub Btn2_Nuancier_Click(sender As Object, e As EventArgs) Handles Btn07_02.Click
    ' Restauration des couleurs d'origine de Coloriage 
    Color_List = Color_Originale_List
    DGV_Color_Display()
  End Sub

  Private Sub Btn3_Nuancier_Click(sender As Object, e As EventArgs) Handles Btn07_03.Click
    ' Lister des couleurs
    Color_List_Display()
  End Sub

  Private Sub CB08_01_CheckedChanged(sender As Object, e As EventArgs) Handles CB08_01.CheckedChanged
    Select Case CB08_01.CheckState
      Case CheckState.Unchecked : Xap = False  '0  'Non
      Case CheckState.Checked : Xap = True     '1  'Oui
    End Select
  End Sub

  Private Sub CBB08_03_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBB08_03.SelectedIndexChanged
    Lettre_Flèche = CBB08_03.SelectedIndex
    My.Settings.Prf_08H_Flèche = CBB08_03.SelectedIndex
    Select Case Lettre_Flèche
      Case 0 : Lettre_Flèche_ChrW = 64
      Case 1 : Lettre_Flèche_ChrW = 96
      Case 2 : Lettre_Flèche_ChrW = 944
      Case 3 : Lettre_Flèche_ChrW = 1039
      Case Else : Lettre_Flèche_ChrW = 64
    End Select
  End Sub

  Private Sub Btn08_99_Click(sender As Object, e As EventArgs) Handles Btn08_99.Click
    Close()
  End Sub
End Class