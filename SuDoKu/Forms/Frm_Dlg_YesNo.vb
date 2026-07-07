Public Class Frm_Dlg_YesNo
  Private Sub Frm_Dlg_YesNo_Load(sender As Object, e As EventArgs)
    Left = Frm_SDK.Left + Frm_SDK.Journal.Left
    Top = Frm_SDK.Top + Frm_SDK.Journal.Top

    Text = Frm_Dlg_Titre
    TB.Text = Frm_Dlg_Texte
  End Sub

  Private Sub Btn_Yes_Click(sender As Object, e As EventArgs) Handles Btn_Yes.Click
    Frm_Dlg_Reponse = "Yes"
    Close()
  End Sub

  Private Sub Btn_No_Click(sender As Object, e As EventArgs) Handles Btn_No.Click
    Frm_Dlg_Reponse = "No"
    Close()
  End Sub
End Class