Option Strict On
Option Explicit On

'-------------------------------------------------------------------------------
'Création : Me 01/02/2023
'https://learn.microsoft.com/en-us/dotnet/visual-basic/developing-apps/programming/app-settings/how-to-create-property-grids-for-user-settings
'
'-------------------------------------------------------------------------------




Public Class Frm_MySettings
  Private Sub Frm_MySettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    PG.SelectedObject = My.Settings
    BackColor = Color_Frm_BackColor
    PG.CommandsBackColor = Color_Fond_Typ_I
    PG.ViewBackColor = Color_Fond_Typ_RV
  End Sub
End Class