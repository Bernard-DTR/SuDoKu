Public Class Frm_Dlg_YesNo
  Public Property Frm_Dlg_Titre As String
  Public Property Frm_Dlg_Texte As String
  Public Property Frm_Dlg_Reponse As String
  Private Sub Frm_Dlg_YesNo_Load(sender As Object, e As EventArgs)
    ' Conversion du point Top-Left de Journal en ccorodonnées écran
    ' en tenant compte du menu et de la barre d'outils
    Dim p As Point = Frm_SDK.Journal.PointToScreen(New Point(0, 0))
    p.Offset(10, 10)
    Me.StartPosition = FormStartPosition.Manual
    Me.Location = p

    Text = Frm_Dlg_Titre
    TB.Text = Frm_Dlg_Texte
  End Sub
  Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
    'Il est possible d'utiliser les touches Entrée (Yes) et Echap (No) pour répondre à la boîte de dialogue
    If keyData = Keys.Enter Then
      Btn_Yes.PerformClick()
      Return True
    End If
    If keyData = Keys.Escape Then
      Btn_No.PerformClick()
      Return True
    End If
    Return MyBase.ProcessCmdKey(msg, keyData)
  End Function
  Private Sub Btn_Yes_Click(sender As Object, e As EventArgs) Handles Btn_Yes.Click
    Frm_Dlg_Reponse = "Yes"
    Close()
  End Sub
  Private Sub Btn_No_Click(sender As Object, e As EventArgs) Handles Btn_No.Click
    Frm_Dlg_Reponse = "No"
    Close()
  End Sub
End Class