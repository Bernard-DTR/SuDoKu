<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_About
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Frm_About))
        Me.Btn_Fermer = New System.Windows.Forms.Button()
        Me.Lbl_01 = New System.Windows.Forms.Label()
        Me.Lbl_02 = New System.Windows.Forms.Label()
        Me.Abt_Mnu = New System.Windows.Forms.MenuStrip()
        Me.Abt_Mnu_00 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Abt_Mnu_0099 = New System.Windows.Forms.ToolStripMenuItem()
        Me.PictureBox = New System.Windows.Forms.PictureBox()
        Me.Abt_Mnu.SuspendLayout()
        CType(Me.PictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Btn_Fermer
        '
        Me.Btn_Fermer.Location = New System.Drawing.Point(15, 189)
        Me.Btn_Fermer.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_Fermer.Name = "Btn_Fermer"
        Me.Btn_Fermer.Size = New System.Drawing.Size(809, 35)
        Me.Btn_Fermer.TabIndex = 0
        Me.Btn_Fermer.Text = "OK"
        Me.Btn_Fermer.UseVisualStyleBackColor = True
        '
        'Lbl_01
        '
        Me.Lbl_01.AutoSize = True
        Me.Lbl_01.Location = New System.Drawing.Point(114, 62)
        Me.Lbl_01.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_01.Name = "Lbl_01"
        Me.Lbl_01.Size = New System.Drawing.Size(57, 20)
        Me.Lbl_01.TabIndex = 1
        Me.Lbl_01.Text = "Auteur"
        '
        'Lbl_02
        '
        Me.Lbl_02.AutoSize = True
        Me.Lbl_02.Location = New System.Drawing.Point(114, 117)
        Me.Lbl_02.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_02.Name = "Lbl_02"
        Me.Lbl_02.Size = New System.Drawing.Size(63, 20)
        Me.Lbl_02.TabIndex = 2
        Me.Lbl_02.Text = "Version"
        '
        'Abt_Mnu
        '
        Me.Abt_Mnu.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.Abt_Mnu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.Abt_Mnu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Abt_Mnu_00})
        Me.Abt_Mnu.Location = New System.Drawing.Point(0, 0)
        Me.Abt_Mnu.Name = "Abt_Mnu"
        Me.Abt_Mnu.Size = New System.Drawing.Size(831, 33)
        Me.Abt_Mnu.TabIndex = 4
        Me.Abt_Mnu.Text = "MenuStrip1"
        '
        'Abt_Mnu_00
        '
        Me.Abt_Mnu_00.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Abt_Mnu_0099})
        Me.Abt_Mnu_00.Name = "Abt_Mnu_00"
        Me.Abt_Mnu_00.Size = New System.Drawing.Size(78, 29)
        Me.Abt_Mnu_00.Text = "Fichier"
        '
        'Abt_Mnu_0099
        '
        Me.Abt_Mnu_0099.Name = "Abt_Mnu_0099"
        Me.Abt_Mnu_0099.ShortcutKeys = CType((System.Windows.Forms.Keys.Alt Or System.Windows.Forms.Keys.F4), System.Windows.Forms.Keys)
        Me.Abt_Mnu_0099.Size = New System.Drawing.Size(234, 34)
        Me.Abt_Mnu_0099.Text = "&Fermer"
        '
        'PictureBox
        '
        Me.PictureBox.Image = CType(resources.GetObject("PictureBox.Image"), System.Drawing.Image)
        Me.PictureBox.Location = New System.Drawing.Point(20, 62)
        Me.PictureBox.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox.Name = "PictureBox"
        Me.PictureBox.Size = New System.Drawing.Size(70, 72)
        Me.PictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox.TabIndex = 5
        Me.PictureBox.TabStop = False
        '
        'Frm_About
        '
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(831, 240)
        Me.Controls.Add(Me.PictureBox)
        Me.Controls.Add(Me.Lbl_02)
        Me.Controls.Add(Me.Lbl_01)
        Me.Controls.Add(Me.Btn_Fermer)
        Me.Controls.Add(Me.Abt_Mnu)
        Me.Cursor = System.Windows.Forms.Cursors.Hand
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MainMenuStrip = Me.Abt_Mnu
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Frm_About"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "A Propos..."
        Me.TopMost = True
        Me.Abt_Mnu.ResumeLayout(False)
        Me.Abt_Mnu.PerformLayout()
        CType(Me.PictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Btn_Fermer As System.Windows.Forms.Button
    Friend WithEvents Lbl_01 As System.Windows.Forms.Label
    Friend WithEvents Lbl_02 As System.Windows.Forms.Label
    Friend WithEvents Abt_Mnu As MenuStrip
    Friend WithEvents Abt_Mnu_00 As ToolStripMenuItem
    Friend WithEvents Abt_Mnu_0099 As ToolStripMenuItem
    Friend WithEvents PictureBox As PictureBox
End Class