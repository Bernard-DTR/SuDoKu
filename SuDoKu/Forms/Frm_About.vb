Option Strict On
Option Explicit On
Imports System.Reflection
Public NotInheritable Class Frm_About

  '-------------------------------------------------------------------------------
  ' Menu About
  '-------------------------------------------------------------------------------
  Private Sub Frm_About_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font

    Left = Frm_SDK.Left + 20
    Top = Frm_SDK.Top + 20
    Lbl_01.Text = "BD_Euring"
    Dim Label2 As String = ""
    Label2 &= Application.ProductName & " " & SDK_Version & vbCrLf
    Label2 &= Application.ExecutablePath & vbCrLf
    Label2 &= "Version du CLR/.NET : " & Environment.Version.ToString() & vbCrLf
    Dim vbCompiler As Assembly = GetType(Microsoft.VisualBasic.VBCodeProvider).Assembly
    Label2 &= "Version du compilateur VB : " & vbCompiler.GetName().Version.ToString() & vbCrLf
    Lbl_02.Text = Label2
    BackColor = Color_Frm_BackColor
    Icon = My.Resources.SuDoKu
  End Sub
  Private Sub Btn_Fermer_Click(sender As Object, e As EventArgs) Handles Btn_Fermer.Click
    Me.Close()
  End Sub
  Private Sub Abt_Mnu_0099_Click(sender As Object, e As EventArgs) Handles Abt_Mnu_0099.Click
    Me.Close()
  End Sub

End Class