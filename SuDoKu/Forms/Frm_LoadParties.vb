Option Strict On
Option Explicit On

Public Class Frm_LoadParties
  '-------------------------------------------------------------------------------
  '
  '  Formulaire de Gestion des Parties.
  '    Ce module permet de choisir une grille à jouer à partir d'une liste des noms.
  '-------------------------------------------------------------------------------
  Private Sub SDK_Frm_Load_Activated(sender As Object, e As EventArgs) Handles MyBase.Activated
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    With LV_Parties
      .View = View.Details
      .GridLines = True
      .FullRowSelect = True
      .Clear()
      ' Création de 3 Colonnes Nom_de_la_Partie, Force et Indication de Solution
      Nsd_CH = .Columns.Add("Nom", 115, HorizontalAlignment.Left)
      Nsd_CH = .Columns.Add("F", 22, HorizontalAlignment.Left)
      Nsd_CH = .Columns.Add("S", 22, HorizontalAlignment.Left)
      For i As Integer = 0 To Pzzl_List.Count - 1
        ' Création des Items et des SubItems avec une ListViewItem
        Dim LV_Item As New ListViewItem(Pzzl_List.Item(i).Nom)             ' Nom
        Nsd_LV = LV_Item.SubItems.Add(Pzzl_List.Item(i).Force)             ' Force
        Dim Sol As String = "#"
        'Si la solution comporte un 0, alors la solution n'est pas complète
        If Pzzl_List.Item(i).Solution.Contains("0") = True Then Sol = " " Else Sol = "S"  ' S ou Blanc
        Nsd_LV = LV_Item.SubItems.Add(Sol)
        LV_Parties.Items.Add(LV_Item)
      Next i
      Me.Text = CStr(.Items.Count) & " puzzles."
    End With
  End Sub
  Private Sub SDK_Frm_Load_Load(sender As Object, e As EventArgs) Handles Me.Load
    'La fenêtre apparaît dans l'ordre de plan le plus haut.
    TopMost = True
    Left = Frm_SDK.Left + 20
    Top = Frm_SDK.Top + 20
    Dim Last_Idx As Integer = My.Settings.LP_Numéro
    LV_Parties.Items(Last_Idx).Focused = True
    LV_Parties.Items(Last_Idx).Selected = True
    'Garantit que l'élément spécifié reste toujours visible dans le contrôle, en faisant défiler le contenu du contrôle si nécessaire.
    LV_Parties.Items(Last_Idx).EnsureVisible()
    Me.BackColor = Color_Frm_BackColor
    Me.Icon = My.Resources.SuDoKu
  End Sub
  Private Sub LV_Parties_DoubleClick(sender As Object, e As EventArgs) Handles LV_Parties.DoubleClick
    Btn_Charger_Click(Me, e)
  End Sub
  Private Sub Btn_Charger_Click(sender As Object, e As EventArgs) Handles Btn_Charger.Click
    ' Chargement d'une partie
    Dim LV_Selected As ListView.SelectedListViewItemCollection = Me.LV_Parties.SelectedItems
    My.Settings.LP_Numéro = LV_Selected.Item(0).Index
    Pzzl_Load(LV_Selected.Item(0).SubItems(0).Text)
    Me.Close()
  End Sub
  Private Sub Btn_Fermer_Click(sender As Object, e As EventArgs) Handles Btn_Fermer.Click
    Close()
  End Sub
End Class