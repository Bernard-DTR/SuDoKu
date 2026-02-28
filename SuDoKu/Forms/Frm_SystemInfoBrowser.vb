Option Strict On
Option Explicit On
'-------------------------------------------------------------------------------
'Création: 31/01/2023
'https://learn.microsoft.com/en-us/dotnet/api/system.windows.forms.systeminformation.mousewheelscrolllines?view=windowsdesktop-7.0
'
'-------------------------------------------------------------------------------
Imports System.Reflection ' Nécessaire pour Property Info

Public Class Frm_SystemInfoBrowser

  Private LB1 As System.Windows.Forms.ListBox
  Private LB2 As System.Windows.Forms.ListBox
  Private TB As System.Windows.Forms.TextBox
  Private l As Integer
  Public Sub New()
    InitializeComponent()
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    SuspendLayout()
    LB1 = New System.Windows.Forms.ListBox With {
        .Anchor = CType(System.Windows.Forms.AnchorStyles.Top _
                     Or System.Windows.Forms.AnchorStyles.Bottom _
                     Or System.Windows.Forms.AnchorStyles.Left _
                     Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles),
        .Location = New System.Drawing.Point(8, 16),
        .Size = New System.Drawing.Size(272, 60),
        .BackColor = Color_Fond_Typ_I,
        .Font = New Drawing.Font("Courier New", 9.75!, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CType(0, Byte)),
        .TabIndex = 0
    }
    Nsd_i = LB1.Items.Add("Application")
    Nsd_i = LB1.Items.Add("Environment")
    Nsd_i = LB1.Items.Add("SystemInformation")

    Me.LB2 = New System.Windows.Forms.ListBox With {
        .Anchor = CType(System.Windows.Forms.AnchorStyles.Top _
                     Or System.Windows.Forms.AnchorStyles.Bottom _
                     Or System.Windows.Forms.AnchorStyles.Left _
                     Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles),
        .Location = New System.Drawing.Point(8, 86),
        .Size = New System.Drawing.Size(272, 426),
        .BackColor = Color_Fond_Typ_I,
        .Font = New Drawing.Font("Courier New", 9.75!, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CType(0, Byte)),
        .TabIndex = 0
    }
    Me.TB = New System.Windows.Forms.TextBox With {
        .Anchor = CType(System.Windows.Forms.AnchorStyles.Top _
                     Or System.Windows.Forms.AnchorStyles.Bottom _
                     Or System.Windows.Forms.AnchorStyles.Left _
                     Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles),
        .Location = New System.Drawing.Point(288, 16),
        .Multiline = True,
        .ScrollBars = System.Windows.Forms.ScrollBars.Vertical,
        .Size = New System.Drawing.Size(520, 496),
        .BackColor = Color_Fond_Typ_RV,
        .Font = New Drawing.Font("Courier New", 12.0!, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CType(0, Byte)),
       .TabIndex = 1
    }
    Me.ClientSize = New System.Drawing.Size(816, 525)
    Me.Controls.Add(Me.LB1)
    Me.Controls.Add(Me.LB2)
    Me.Controls.Add(Me.TB)
    Me.BackColor = Color_Frm_BackColor
    Me.FormBorderStyle = FormBorderStyle.FixedSingle
    Me.MaximizeBox = False
    TB.Text = "Choisir une classe"
    Me.Text = "Liste des Propriétés"

    AddHandler LB1.SelectedIndexChanged, AddressOf LB1_SelectedIndexChanged
    AddHandler LB2.SelectedIndexChanged, AddressOf LB2_SelectedIndexChanged
    Me.ResumeLayout(False)
  End Sub

  Private Sub LB1_SelectedIndexChanged(sender As Object, e As EventArgs)
    If LB1.SelectedIndex = -1 Then
      Return
    End If
    TB.Text = ""
    LB2.Items.Clear()

    Dim T As Type = Nothing
    Try
      Select Case LB1.SelectedItem.ToString()
        Case "Application" : T = GetType(Application)
        Case "Environment" : T = GetType(Environment)
        Case "SystemInformation" : T = GetType(SystemInformation)
      End Select
      Dim Info() As PropertyInfo = T.GetProperties()
      l = Info.Length - 1
      Dim Tab(l) As String
      For i As Integer = 0 To l : Tab(i) = Info(i).Name : Next i
      Array.Sort(Tab)
      For i As Integer = 0 To l : Nsd_i = LB2.Items.Add(Tab(i)) : Next i
      Me.Text = LB1.SelectedItem.ToString() + " : " + Info.Length.ToString() + " Properties."
    Catch ex As Exception
      TB.Text += ex.Message & vbCrLf
    End Try
    TB.Text += "Choisir une propriété"
  End Sub

  Private Sub LB2_SelectedIndexChanged(sender As Object, e As EventArgs)
    If LB2.SelectedIndex = -1 Then
      Return
    End If

    Try
      TB.Text = ""
      Dim T As Type = Nothing

      Select Case LB1.SelectedItem.ToString()
        Case "Application" : T = GetType(Application)
        Case "Environment" : T = GetType(Environment)
        Case "SystemInformation" : T = GetType(SystemInformation)
      End Select

      Dim Info_T As PropertyInfo() = T.GetProperties()
      Dim Info As PropertyInfo = Nothing
      l = Info_T.Length - 1
      For i As Integer = 0 To l
        If Info_T(i).Name = LB2.SelectedItem.ToString() Then
          Info = Info_T(i)
          Exit For
        End If
      Next i
      Dim Property_Value As Object = Info.GetValue(Nothing, Nothing)
      TB.Text = "Valeur de : " + vbCrLf + vbCrLf +
                LB1.SelectedItem.ToString() + vbCrLf +
                LB2.SelectedItem.ToString() + vbCrLf + vbCrLf +
                Property_Value.ToString()
    Catch ex As Exception
      TB.Text += ex.Message & vbCrLf
    End Try
  End Sub
End Class