<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Dlg_YesNo
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
    Me.Btn_Yes = New System.Windows.Forms.Button()
    Me.Btn_No = New System.Windows.Forms.Button()
    Me.TB = New System.Windows.Forms.TextBox()
    Me.SuspendLayout()
    '
    'Btn_Yes
    '
    Me.Btn_Yes.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Btn_Yes.Location = New System.Drawing.Point(20, 306)
    Me.Btn_Yes.Name = "Btn_Yes"
    Me.Btn_Yes.Size = New System.Drawing.Size(194, 53)
    Me.Btn_Yes.TabIndex = 0
    Me.Btn_Yes.Text = "Oui"
    Me.Btn_Yes.UseVisualStyleBackColor = True
    AddHandler Me.Btn_Yes.Click, AddressOf Me.Btn_Yes_Click
    '
    'Btn_No
    '
    Me.Btn_No.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Btn_No.Location = New System.Drawing.Point(220, 308)
    Me.Btn_No.Name = "Btn_No"
    Me.Btn_No.Size = New System.Drawing.Size(194, 53)
    Me.Btn_No.TabIndex = 1
    Me.Btn_No.Text = "Non"
    Me.Btn_No.UseVisualStyleBackColor = True
    AddHandler Me.Btn_No.Click, AddressOf Me.Btn_No_Click
    '
    'TB
    '
    Me.TB.Font = New System.Drawing.Font("Courier New", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.TB.Location = New System.Drawing.Point(20, 27)
    Me.TB.Multiline = True
    Me.TB.Name = "TB"
    Me.TB.Size = New System.Drawing.Size(394, 266)
    Me.TB.TabIndex = 2
    '
    'Frm_Dlg_YesNo
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(437, 380)
    Me.Controls.Add(Me.TB)
    Me.Controls.Add(Me.Btn_No)
    Me.Controls.Add(Me.Btn_Yes)
    Me.HelpButton = True
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "Frm_Dlg_YesNo"
    Me.Text = "Frm_Dlg_YesNo"
    AddHandler Load, AddressOf Me.Frm_Dlg_YesNo_Load
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents Btn_Yes As Button
    Friend WithEvents Btn_No As Button
    Friend WithEvents TB As TextBox
End Class
