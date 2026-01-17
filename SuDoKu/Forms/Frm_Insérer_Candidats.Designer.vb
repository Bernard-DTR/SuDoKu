<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Insérer_Candidats
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Btn_Insérer = New System.Windows.Forms.Button()
        Me.TB_Candidats = New System.Windows.Forms.TextBox()
        Me.Btn_TTT = New System.Windows.Forms.Button()
        Me.PB_G = New System.Windows.Forms.PictureBox()
        Me.CB01 = New System.Windows.Forms.CheckBox()
        Me.PB_D = New System.Windows.Forms.PictureBox()
        Me.PB_BG = New System.Windows.Forms.PictureBox()
        Me.PB_BD = New System.Windows.Forms.PictureBox()
        Me.TT = New System.Windows.Forms.ToolTip(Me.components)
        Me.PB_A = New System.Windows.Forms.PictureBox()
        CType(Me.PB_G, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB_D, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB_BG, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB_BD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PB_A, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_Insérer
        '
        Me.Btn_Insérer.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Btn_Insérer.Location = New System.Drawing.Point(95, 35)
        Me.Btn_Insérer.Name = "Btn_Insérer"
        Me.Btn_Insérer.Size = New System.Drawing.Size(119, 23)
        Me.Btn_Insérer.TabIndex = 1
        Me.Btn_Insérer.Text = "Insérer"
        Me.Btn_Insérer.UseVisualStyleBackColor = True
        '
        'TB_Candidats
        '
        Me.TB_Candidats.Location = New System.Drawing.Point(30, 5)
        Me.TB_Candidats.Name = "TB_Candidats"
        Me.TB_Candidats.Size = New System.Drawing.Size(60, 20)
        Me.TB_Candidats.TabIndex = 4
        Me.TT.SetToolTip(Me.TB_Candidats, " ")
        '
        'Btn_TTT
        '
        Me.Btn_TTT.Location = New System.Drawing.Point(95, 5)
        Me.Btn_TTT.Name = "Btn_TTT"
        Me.Btn_TTT.Size = New System.Drawing.Size(120, 23)
        Me.Btn_TTT.TabIndex = 6
        Me.Btn_TTT.Text = "Copier le ToolTipText"
        Me.Btn_TTT.UseVisualStyleBackColor = True
        '
        'PB_G
        '
        Me.PB_G.BackColor = System.Drawing.Color.Beige
        Me.PB_G.Image = Global.SuDoKu.My.Resources.Resources.Aimant_Gauche
        Me.PB_G.Location = New System.Drawing.Point(5, 5)
        Me.PB_G.Name = "PB_G"
        Me.PB_G.Size = New System.Drawing.Size(20, 55)
        Me.PB_G.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_G.TabIndex = 10
        Me.PB_G.TabStop = False
        Me.TT.SetToolTip(Me.PB_G, " ")
        '
        'CB01
        '
        Me.CB01.AutoSize = True
        Me.CB01.Location = New System.Drawing.Point(31, 34)
        Me.CB01.Name = "CB01"
        Me.CB01.Size = New System.Drawing.Size(49, 17)
        Me.CB01.TabIndex = 11
        Me.CB01.Text = "Ins_I"
        Me.TT.SetToolTip(Me.CB01, " ")
        Me.CB01.UseVisualStyleBackColor = True
        '
        'PB_D
        '
        Me.PB_D.BackColor = System.Drawing.Color.Beige
        Me.PB_D.Image = Global.SuDoKu.My.Resources.Resources.Aimant_Droit
        Me.PB_D.Location = New System.Drawing.Point(220, 5)
        Me.PB_D.Name = "PB_D"
        Me.PB_D.Size = New System.Drawing.Size(20, 55)
        Me.PB_D.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_D.TabIndex = 12
        Me.PB_D.TabStop = False
        Me.TT.SetToolTip(Me.PB_D, " ")
        '
        'PB_BG
        '
        Me.PB_BG.BackColor = System.Drawing.Color.Beige
        Me.PB_BG.Image = Global.SuDoKu.My.Resources.Resources.Aimant_Bas
        Me.PB_BG.Location = New System.Drawing.Point(5, 66)
        Me.PB_BG.Name = "PB_BG"
        Me.PB_BG.Size = New System.Drawing.Size(85, 20)
        Me.PB_BG.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_BG.TabIndex = 13
        Me.PB_BG.TabStop = False
        Me.TT.SetToolTip(Me.PB_BG, " ")
        '
        'PB_BD
        '
        Me.PB_BD.BackColor = System.Drawing.Color.Beige
        Me.PB_BD.Image = Global.SuDoKu.My.Resources.Resources.Aimant_Bas
        Me.PB_BD.Location = New System.Drawing.Point(156, 66)
        Me.PB_BD.Name = "PB_BD"
        Me.PB_BD.Size = New System.Drawing.Size(85, 20)
        Me.PB_BD.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_BD.TabIndex = 14
        Me.PB_BD.TabStop = False
        Me.TT.SetToolTip(Me.PB_BD, " ")
        '
        'PB_A
        '
        Me.PB_A.BackColor = System.Drawing.Color.Beige
        Me.PB_A.Image = Global.SuDoKu.My.Resources.Resources.Annuler
        Me.PB_A.Location = New System.Drawing.Point(96, 66)
        Me.PB_A.Name = "PB_A"
        Me.PB_A.Size = New System.Drawing.Size(54, 20)
        Me.PB_A.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PB_A.TabIndex = 15
        Me.PB_A.TabStop = False
        Me.TT.SetToolTip(Me.PB_A, " ")
        '
        'Frm_Insérer_Candidats
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(247, 92)
        Me.Controls.Add(Me.PB_A)
        Me.Controls.Add(Me.PB_BD)
        Me.Controls.Add(Me.PB_BG)
        Me.Controls.Add(Me.PB_D)
        Me.Controls.Add(Me.CB01)
        Me.Controls.Add(Me.PB_G)
        Me.Controls.Add(Me.Btn_TTT)
        Me.Controls.Add(Me.TB_Candidats)
        Me.Controls.Add(Me.Btn_Insérer)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_Insérer_Candidats"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.TopMost = True
        CType(Me.PB_G, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB_D, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB_BG, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB_BD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PB_A, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Btn_Insérer As Button
    Friend WithEvents TB_Candidats As TextBox
    Friend WithEvents Btn_TTT As Button
    Friend WithEvents PB_G As PictureBox
    Friend WithEvents CB01 As CheckBox
    Friend WithEvents PB_D As PictureBox
    Friend WithEvents PB_BG As PictureBox
    Friend WithEvents PB_BD As PictureBox
    Friend WithEvents TT As ToolTip
    Friend WithEvents PB_A As PictureBox
End Class
