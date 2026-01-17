Option Strict On
Option Explicit On
Imports System.Reflection
Public NotInheritable Class Frm_About

  '-------------------------------------------------------------------------------
  ' Menu About
  '-------------------------------------------------------------------------------
  Private Sub Frm_About_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    AutoScaleMode = Me_AutoScaleMode_Standard
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
  'Imports System
  'Imports System.Reflection

  'Module VersionInfo

  '  Sub Main()
  '    ' Version de l'application compilée
  '    Console.WriteLine("Version de l'application : " &
  '                      My.Application.Info.Version.ToString())

  '    ' Version du CLR (runtime .NET)
  '    Console.WriteLine("Version du CLR/.NET : " &
  '                      Environment.Version.ToString())

  '    ' Version du compilateur VB utilisé
  '    Dim vbCompiler = GetType(Microsoft.VisualBasic.VBCodeProvider).Assembly
  '    Console.WriteLine("Version du compilateur VB : " &
  '                      vbCompiler.GetName().Version.ToString())

  '    ' Version de Visual Studio (approximative, via l'IDE)
  '    ' Attention : Visual Studio n'est pas exposé directement par code.
  '    ' On peut récupérer l'info en lisant le manifest ou via "À propos" dans l'IDE.
  '    Console.WriteLine("Pour la version exacte de Visual Studio : menu Aide > À propos")
  '  End Sub

  'End Module
End Class