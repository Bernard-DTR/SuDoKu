<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_LoadParties
  Inherits System.Windows.Forms.Form

  'Form remplace la méthode Dispose pour nettoyer la liste des composants.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    Try
      If disposing AndAlso components IsNot Nothing Then
        components.Dispose()
      End If
    Finally
      MyBase.Dispose(disposing)
    End Try
  End Sub

  'Requise par le Concepteur Windows Form
  Private components As System.ComponentModel.IContainer

  'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
  'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
  'Ne la modifiez pas à l'aide de l'éditeur de code.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
        Me.Btn_Charger = New System.Windows.Forms.Button()
        Me.Btn_Fermer = New System.Windows.Forms.Button()
        Me.Lbl_02 = New System.Windows.Forms.Label()
        Me.LV_Parties = New System.Windows.Forms.ListView()
        Me.Nom = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Force = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Solution = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'Btn_Charger
        '
        Me.Btn_Charger.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Btn_Charger.Location = New System.Drawing.Point(10, 469)
        Me.Btn_Charger.Margin = New System.Windows.Forms.Padding(2)
        Me.Btn_Charger.Name = "Btn_Charger"
        Me.Btn_Charger.Size = New System.Drawing.Size(183, 25)
        Me.Btn_Charger.TabIndex = 2
        Me.Btn_Charger.Text = "&Charger"
        Me.Btn_Charger.UseVisualStyleBackColor = True
        '
        'Btn_Fermer
        '
        Me.Btn_Fermer.Location = New System.Drawing.Point(10, 502)
        Me.Btn_Fermer.Margin = New System.Windows.Forms.Padding(2)
        Me.Btn_Fermer.Name = "Btn_Fermer"
        Me.Btn_Fermer.Size = New System.Drawing.Size(183, 25)
        Me.Btn_Fermer.TabIndex = 10
        Me.Btn_Fermer.Text = "&Fermer"
        Me.Btn_Fermer.UseVisualStyleBackColor = True
        '
        'Lbl_02
        '
        Me.Lbl_02.Location = New System.Drawing.Point(12, 12)
        Me.Lbl_02.Name = "Lbl_02"
        Me.Lbl_02.Size = New System.Drawing.Size(181, 20)
        Me.Lbl_02.TabIndex = 2
        Me.Lbl_02.Text = "Choisir un Puzzle :"
        '
        'LV_Parties
        '
        Me.LV_Parties.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.LV_Parties.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Nom, Me.Force, Me.Solution})
        Me.LV_Parties.HideSelection = False
        Me.LV_Parties.Location = New System.Drawing.Point(11, 35)
        Me.LV_Parties.MultiSelect = False
        Me.LV_Parties.Name = "LV_Parties"
        Me.LV_Parties.Size = New System.Drawing.Size(182, 424)
        Me.LV_Parties.TabIndex = 1
        Me.LV_Parties.UseCompatibleStateImageBehavior = False
        '
        'Nom
        '
        Me.Nom.Text = "Nom"
        '
        'Force
        '
        Me.Force.Text = "F"
        Me.Force.Width = 30
        '
        'Solution
        '
        Me.Solution.Text = "Sol"
        Me.Solution.Width = 30
        '
        'Frm_LoadParties
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(204, 535)
        Me.Controls.Add(Me.LV_Parties)
        Me.Controls.Add(Me.Lbl_02)
        Me.Controls.Add(Me.Btn_Fermer)
        Me.Controls.Add(Me.Btn_Charger)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.KeyPreview = True
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_LoadParties"
        Me.Text = "Ouvrir"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Btn_Charger As System.Windows.Forms.Button
  Friend WithEvents Btn_Fermer As System.Windows.Forms.Button
  Friend WithEvents Lbl_02 As System.Windows.Forms.Label
  Friend WithEvents LV_Parties As System.Windows.Forms.ListView
  Friend WithEvents Nom As System.Windows.Forms.ColumnHeader
  Friend WithEvents Force As System.Windows.Forms.ColumnHeader
  Friend WithEvents Solution As System.Windows.Forms.ColumnHeader
End Class
