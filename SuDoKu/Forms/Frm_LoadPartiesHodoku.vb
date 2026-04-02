Imports SuDoKu.DancingLink

Public Class Frm_LoadPartiesHodoku
  '-------------------------------------------------------------------------------
  '
  '  Formulaire de Gestion des Parties de test Hodoku, identique à la Gestion des Parties
  '    Ce module permet de choisir une grille à jouer à partir
  '       d'une liste des noms.
  '    La liste des parties de test a été chargée dès le lancement du jeu dans une List
  '    Les parties sont issues d'une librairie du jeu Hodoku
  '    Le fichier comporte des noms de familles de jeu
  '-------------------------------------------------------------------------------

  Private Sub Frm_LoadPartiesHodoku_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    'La fenêtre apparaît dans l'ordre de plan le plus haut.
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    TopMost = True
    Left = Frm_SDK.Left + 20
    Top = Frm_SDK.Top + 20
    'Chargement de LV_Parties et positionnement sur le dernier choisi
    With LV_Parties
      .View = View.Details
      .GridLines = True
      .FullRowSelect = True
      .Clear()
      Nsd_CH = .Columns.Add("Nom", 180, HorizontalAlignment.Left)
      Nsd_CH = .Columns.Add("Jeu", 30, HorizontalAlignment.Left)
      For i As Integer = 0 To Pzzl_Hodoku_List.Count - 1
        ' Création des Items et des SubItems avec une ListViewItem
        Dim LV_Item As New ListViewItem(Pzzl_Hodoku_List.Item(i).Nom)   ' Nom
        Nsd_LV = LV_Item.SubItems.Add(Pzzl_Hodoku_List.Item(i).Jeu)              ' Jeu
        LV_Parties.Items.Add(LV_Item)
      Next i
      Me.Text = CStr(.Items.Count + 1) & " puzzles."
    End With
    Dim Last_Idx As Integer = My.Settings.LP_Numéro_Hodoku
    LV_Parties.Items(Last_Idx).Focused = True
    LV_Parties.Items(Last_Idx).Selected = True
    LV_Parties.Items(Last_Idx).EnsureVisible()
    BackColor = Color_Frm_BackColor
    Me.Icon = My.Resources.SuDoKu
  End Sub
  Private Sub LV_Parties_DoubleClick(sender As Object, e As EventArgs) Handles LV_Parties.DoubleClick
    Btn_Charger_Click(Me, e)
  End Sub
  Private Sub Btn_Charger_Click(sender As Object, e As EventArgs) Handles Btn_Charger.Click
    ' Chargement d'une partie de test
    Dim LV_Selected As ListView.SelectedListViewItemCollection = Me.LV_Parties.SelectedItems
    'La partie sélectionnée doit être un jeu et non un commentaire
    If LV_Selected.Item(0).SubItems(1).Text = "Oui" Then
      Dim idx As Integer = LV_Selected.Item(0).Index
      Dim Nom As String = "Hodoku_" & LV_Selected.Item(0).SubItems(0).Text & "_" & CStr(idx) & "_hdk"
      Dim Prb As String = Pzzl_Hodoku_List(idx).Problème
      Prb = Prb.Replace(".", " ")
      Prb = Prb.PadRight(81)
      If Prb.Length > 81 Then Prb = Prb.Substring(0, 81)
      Dim Sol As String = StrDup(81, " ")
      Dim Frc As String = "5"
      My.Settings.LP_Numéro_Hodoku = LV_Selected.Item(0).Index
      Close()
      Jrn_Add(, {"Chargement de " & Nom & " , issue de la bibliothèque de tests Hodoku 220"})
      ' Calcul systématique d'une solution pour les parties de test Hodoku 220
      ' Attention: Data_Links utilise U_Temp, comme U_Checking
      Dim U_temp(80, 3) As String
      For i As Integer = 0 To 80
        U_temp(i, 1) = Prb(i)                       ' Valeur Initiale équivalent à = Prb.Substring(i, 1)
        U_temp(i, 2) = Prb(i)                       ' Valeur  
        U_temp(i, 3) = Cnddts                       ' Candidats
        U_Sol(i) = " "                              ' Solution
      Next i

      Dim DL_Solution As String = StrDup(81, " ")
      Dim DL As DL_Solve_Struct = A_Copyright.DL_Solve(U_temp)
      If DL.DLCode = "Dlu" Then
        DL_Solution = DL.Solution(0)
        Jrn_Add(, {"Puzzle Hodoku 220 sans solution, une solution a été calculée."})
      End If
      Game_New_Game(Gnrl:=Plcy_Gnrl, "   ", Nom:=Nom, Prb:=Prb, Jeu:=Prb, Sol:=DL_Solution, Cdd729:=StrDup(729, " "), Frc:=Frc, Proc_Name_Get())
      'Les 3 variables LP_ ont été documentées. 
      Jrn_Add(, {"PR_Nom=" & Nom})
      Jrn_Add(, {"PR_Ini=" & LP_Prb.Replace(" ", ".")})
      Jrn_Add(, {"PR_Val=" & LP_Jeu.Replace(" ", ".")})
      Jrn_Add(, {"PR_Sol=" & LP_Sol.Replace(" ", ".")})
    End If
    Try
      Close()
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Btn_Fermer_Click(sender As Object, e As EventArgs) Handles Btn_Fermer.Click
    Close()
  End Sub
End Class