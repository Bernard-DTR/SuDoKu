<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Frm_Police
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
        Me.DGV = New System.Windows.Forms.DataGridView()
        Me.CB_Police_Installed = New System.Windows.Forms.ComboBox()
        Me.Btn_Fermer = New System.Windows.Forms.Button()
        Me.CB_Size = New System.Windows.Forms.ComboBox()
        Me.Btn_Charmap = New System.Windows.Forms.Button()
        Me.CB_Police_Private = New System.Windows.Forms.ComboBox()
        Me.Lbl_Police_Installed = New System.Windows.Forms.Label()
        Me.Lbl_Police_Private = New System.Windows.Forms.Label()
        Me.Lbl_Size = New System.Windows.Forms.Label()
        Me.CB_Style = New System.Windows.Forms.ComboBox()
        Me.Lbl_Style = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DGV
        '
        Me.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DGV.Location = New System.Drawing.Point(18, 29)
        Me.DGV.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV.Name = "DGV"
        Me.DGV.RowHeadersWidth = 62
        Me.DGV.Size = New System.Drawing.Size(1440, 743)
        Me.DGV.TabIndex = 0
        '
        'CB_Police_Installed
        '
        Me.CB_Police_Installed.FormattingEnabled = True
        Me.CB_Police_Installed.Location = New System.Drawing.Point(1514, 48)
        Me.CB_Police_Installed.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.CB_Police_Installed.Name = "CB_Police_Installed"
        Me.CB_Police_Installed.Size = New System.Drawing.Size(180, 28)
        Me.CB_Police_Installed.TabIndex = 27
        '
        'Btn_Fermer
        '
        Me.Btn_Fermer.Location = New System.Drawing.Point(1514, 737)
        Me.Btn_Fermer.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_Fermer.Name = "Btn_Fermer"
        Me.Btn_Fermer.Size = New System.Drawing.Size(182, 35)
        Me.Btn_Fermer.TabIndex = 28
        Me.Btn_Fermer.Text = "Fermer"
        Me.Btn_Fermer.UseVisualStyleBackColor = True
        '
        'CB_Size
        '
        Me.CB_Size.FormattingEnabled = True
        Me.CB_Size.Items.AddRange(New Object() {"10", "12", "15", "20", "30", "40", "50", "70"})
        Me.CB_Size.Location = New System.Drawing.Point(1514, 203)
        Me.CB_Size.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.CB_Size.Name = "CB_Size"
        Me.CB_Size.Size = New System.Drawing.Size(180, 28)
        Me.CB_Size.TabIndex = 29
        '
        'Btn_Charmap
        '
        Me.Btn_Charmap.Location = New System.Drawing.Point(1514, 675)
        Me.Btn_Charmap.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_Charmap.Name = "Btn_Charmap"
        Me.Btn_Charmap.Size = New System.Drawing.Size(182, 35)
        Me.Btn_Charmap.TabIndex = 30
        Me.Btn_Charmap.Text = "CharMap"
        Me.Btn_Charmap.UseVisualStyleBackColor = True
        '
        'CB_Police_Private
        '
        Me.CB_Police_Private.FormattingEnabled = True
        Me.CB_Police_Private.Location = New System.Drawing.Point(1514, 128)
        Me.CB_Police_Private.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.CB_Police_Private.Name = "CB_Police_Private"
        Me.CB_Police_Private.Size = New System.Drawing.Size(180, 28)
        Me.CB_Police_Private.TabIndex = 31
        '
        'Lbl_Police_Installed
        '
        Me.Lbl_Police_Installed.AutoSize = True
        Me.Lbl_Police_Installed.Location = New System.Drawing.Point(1514, 18)
        Me.Lbl_Police_Installed.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Police_Installed.Name = "Lbl_Police_Installed"
        Me.Lbl_Police_Installed.Size = New System.Drawing.Size(131, 20)
        Me.Lbl_Police_Installed.TabIndex = 32
        Me.Lbl_Police_Installed.Text = "Polices Installées"
        '
        'Lbl_Police_Private
        '
        Me.Lbl_Police_Private.AutoSize = True
        Me.Lbl_Police_Private.Location = New System.Drawing.Point(1509, 103)
        Me.Lbl_Police_Private.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Police_Private.Name = "Lbl_Police_Private"
        Me.Lbl_Police_Private.Size = New System.Drawing.Size(114, 20)
        Me.Lbl_Police_Private.TabIndex = 33
        Me.Lbl_Police_Private.Text = "Polices Privées"
        '
        'Lbl_Size
        '
        Me.Lbl_Size.AutoSize = True
        Me.Lbl_Size.Location = New System.Drawing.Point(1514, 178)
        Me.Lbl_Size.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Size.Name = "Lbl_Size"
        Me.Lbl_Size.Size = New System.Drawing.Size(40, 20)
        Me.Lbl_Size.TabIndex = 34
        Me.Lbl_Size.Text = "Size"
        '
        'CB_Style
        '
        Me.CB_Style.FormattingEnabled = True
        Me.CB_Style.Items.AddRange(New Object() {"Gras", "Italique", "Normal", "Barré", "Souligné"})
        Me.CB_Style.Location = New System.Drawing.Point(1514, 286)
        Me.CB_Style.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.CB_Style.Name = "CB_Style"
        Me.CB_Style.Size = New System.Drawing.Size(180, 28)
        Me.CB_Style.TabIndex = 35
        '
        'Lbl_Style
        '
        Me.Lbl_Style.AutoSize = True
        Me.Lbl_Style.Location = New System.Drawing.Point(1514, 262)
        Me.Lbl_Style.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Lbl_Style.Name = "Lbl_Style"
        Me.Lbl_Style.Size = New System.Drawing.Size(44, 20)
        Me.Lbl_Style.TabIndex = 36
        Me.Lbl_Style.Text = "Style"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1720, 24)
        Me.MenuStrip1.TabIndex = 37
        Me.MenuStrip1.Text = "Menu           Strip1    "
        '
        'Frm_Police
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1720, 798)
        Me.Controls.Add(Me.Lbl_Style)
        Me.Controls.Add(Me.CB_Style)
        Me.Controls.Add(Me.Lbl_Size)
        Me.Controls.Add(Me.Lbl_Police_Private)
        Me.Controls.Add(Me.Lbl_Police_Installed)
        Me.Controls.Add(Me.CB_Police_Private)
        Me.Controls.Add(Me.Btn_Charmap)
        Me.Controls.Add(Me.CB_Size)
        Me.Controls.Add(Me.Btn_Fermer)
        Me.Controls.Add(Me.CB_Police_Installed)
        Me.Controls.Add(Me.DGV)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Frm_Police"
        Me.Text = "Frm_Police"
        CType(Me.DGV, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DGV As DataGridView
    Friend WithEvents CB_Police_Installed As ComboBox
    Friend WithEvents Btn_Fermer As Button
    Friend WithEvents CB_Size As ComboBox
    Friend WithEvents Btn_Charmap As Button
    Friend WithEvents CB_Police_Private As ComboBox
    Friend WithEvents Lbl_Police_Installed As Label
    Friend WithEvents Lbl_Police_Private As Label
    Friend WithEvents Lbl_Size As Label
    Friend WithEvents CB_Style As ComboBox
    Friend WithEvents Lbl_Style As Label
    Friend WithEvents MenuStrip1 As MenuStrip
End Class
