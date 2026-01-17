<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_Dictionnaire
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Mnu_0190 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_01 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0101 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0103 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu = New System.Windows.Forms.MenuStrip()
        Me.Mnu_02 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0201 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0202 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0203 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0204 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0205 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Mnu_0206 = New System.Windows.Forms.ToolStripMenuItem()
        Me.TB_Information = New System.Windows.Forms.TextBox()
        Me.BDS = New System.Windows.Forms.BindingSource(Me.components)
        Me.P6 = New System.Windows.Forms.TabPage()
        Me.Btn_MotTRéservés_Load = New System.Windows.Forms.Button()
        Me.DGV_MotTRéservés = New System.Windows.Forms.DataGridView()
        Me.Btn_MotTRéservés_MAJ = New System.Windows.Forms.Button()
        Me.P5 = New System.Windows.Forms.TabPage()
        Me.Btn_MsgT_Reload = New System.Windows.Forms.Button()
        Me.TB_MsgT_Input = New System.Windows.Forms.TextBox()
        Me.DGV_MsgT = New System.Windows.Forms.DataGridView()
        Me.P4 = New System.Windows.Forms.TabPage()
        Me.Btn_PrcT_Reload = New System.Windows.Forms.Button()
        Me.TB_PrcT_Input = New System.Windows.Forms.TextBox()
        Me.DGV_PrcT = New System.Windows.Forms.DataGridView()
        Me.P3 = New System.Windows.Forms.TabPage()
        Me.Btn_MotT_Reload = New System.Windows.Forms.Button()
        Me.TB_MotT_Input = New System.Windows.Forms.TextBox()
        Me.DGV_MotT = New System.Windows.Forms.DataGridView()
        Me.P2 = New System.Windows.Forms.TabPage()
        Me.Btn_CodT_Reload = New System.Windows.Forms.Button()
        Me.TB_CodT_Input = New System.Windows.Forms.TextBox()
        Me.DGV_CodT = New System.Windows.Forms.DataGridView()
        Me.P1 = New System.Windows.Forms.TabPage()
        Me.Btn_Dictionnaire_Load = New System.Windows.Forms.Button()
        Me.DGV_Dictionnaire = New System.Windows.Forms.DataGridView()
        Me.Btn_Dictionnaire_MAJ = New System.Windows.Forms.Button()
        Me.TB = New System.Windows.Forms.TabControl()
        Me.P7 = New System.Windows.Forms.TabPage()
        Me.DGV_PrmT = New System.Windows.Forms.DataGridView()
        Me.Mnu.SuspendLayout()
        CType(Me.BDS, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.P6.SuspendLayout()
        CType(Me.DGV_MotTRéservés, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.P5.SuspendLayout()
        CType(Me.DGV_MsgT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.P4.SuspendLayout()
        CType(Me.DGV_PrcT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.P3.SuspendLayout()
        CType(Me.DGV_MotT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.P2.SuspendLayout()
        CType(Me.DGV_CodT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.P1.SuspendLayout()
        CType(Me.DGV_Dictionnaire, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TB.SuspendLayout()
        Me.P7.SuspendLayout()
        CType(Me.DGV_PrmT, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Mnu_0190
        '
        Me.Mnu_0190.Name = "Mnu_0190"
        Me.Mnu_0190.Size = New System.Drawing.Size(210, 34)
        Me.Mnu_0190.Text = "Fermer"
        '
        'Mnu_01
        '
        Me.Mnu_01.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Mnu_0101, Me.Mnu_0103, Me.Mnu_0190})
        Me.Mnu_01.Name = "Mnu_01"
        Me.Mnu_01.Size = New System.Drawing.Size(78, 29)
        Me.Mnu_01.Text = "Fichier"
        '
        'Mnu_0101
        '
        Me.Mnu_0101.Name = "Mnu_0101"
        Me.Mnu_0101.Size = New System.Drawing.Size(210, 34)
        Me.Mnu_0101.Text = "System Info"
        '
        'Mnu_0103
        '
        Me.Mnu_0103.Name = "Mnu_0103"
        Me.Mnu_0103.Size = New System.Drawing.Size(210, 34)
        Me.Mnu_0103.Text = "My_Settings"
        '
        'Mnu
        '
        Me.Mnu.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.Mnu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.Mnu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Mnu_01, Me.Mnu_02})
        Me.Mnu.Location = New System.Drawing.Point(0, 0)
        Me.Mnu.Name = "Mnu"
        Me.Mnu.Size = New System.Drawing.Size(1694, 33)
        Me.Mnu.TabIndex = 8
        Me.Mnu.Text = "Mnu"
        '
        'Mnu_02
        '
        Me.Mnu_02.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.Mnu_0201, Me.Mnu_0202, Me.Mnu_0203, Me.Mnu_0204, Me.Mnu_0205, Me.Mnu_0206})
        Me.Mnu_02.Name = "Mnu_02"
        Me.Mnu_02.Size = New System.Drawing.Size(74, 29)
        Me.Mnu_02.Text = "Outils"
        '
        'Mnu_0201
        '
        Me.Mnu_0201.Name = "Mnu_0201"
        Me.Mnu_0201.Size = New System.Drawing.Size(325, 34)
        Me.Mnu_0201.Text = "Nombre de lignes de code"
        '
        'Mnu_0202
        '
        Me.Mnu_0202.Name = "Mnu_0202"
        Me.Mnu_0202.Size = New System.Drawing.Size(325, 34)
        Me.Mnu_0202.Text = "Editer M9x_Correction"
        '
        'Mnu_0203
        '
        Me.Mnu_0203.Name = "Mnu_0203"
        Me.Mnu_0203.Size = New System.Drawing.Size(325, 34)
        Me.Mnu_0203.Text = "Formulaire des Fonts"
        '
        'Mnu_0204
        '
        Me.Mnu_0204.Name = "Mnu_0204"
        Me.Mnu_0204.Size = New System.Drawing.Size(325, 34)
        Me.Mnu_0204.Text = "Liste des Processus"
        '
        'Mnu_0205
        '
        Me.Mnu_0205.Name = "Mnu_0205"
        Me.Mnu_0205.Size = New System.Drawing.Size(325, 34)
        Me.Mnu_0205.Text = "Liste des Stratégies"
        '
        'Mnu_0206
        '
        Me.Mnu_0206.Name = "Mnu_0206"
        Me.Mnu_0206.Size = New System.Drawing.Size(325, 34)
        Me.Mnu_0206.Text = "SQL_Sudoku"
        '
        'TB_Information
        '
        Me.TB_Information.Location = New System.Drawing.Point(18, 1092)
        Me.TB_Information.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TB_Information.Name = "TB_Information"
        Me.TB_Information.Size = New System.Drawing.Size(1663, 26)
        Me.TB_Information.TabIndex = 10
        '
        'P6
        '
        Me.P6.Controls.Add(Me.Btn_MotTRéservés_Load)
        Me.P6.Controls.Add(Me.DGV_MotTRéservés)
        Me.P6.Controls.Add(Me.Btn_MotTRéservés_MAJ)
        Me.P6.Location = New System.Drawing.Point(4, 29)
        Me.P6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P6.Name = "P6"
        Me.P6.Size = New System.Drawing.Size(1657, 1009)
        Me.P6.TabIndex = 5
        Me.P6.Text = "MotTRéservés"
        Me.P6.UseVisualStyleBackColor = True
        '
        'Btn_MotTRéservés_Load
        '
        Me.Btn_MotTRéservés_Load.Location = New System.Drawing.Point(142, 954)
        Me.Btn_MotTRéservés_Load.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_MotTRéservés_Load.Name = "Btn_MotTRéservés_Load"
        Me.Btn_MotTRéservés_Load.Size = New System.Drawing.Size(112, 35)
        Me.Btn_MotTRéservés_Load.TabIndex = 13
        Me.Btn_MotTRéservés_Load.Text = "Chargement"
        Me.Btn_MotTRéservés_Load.UseVisualStyleBackColor = True
        '
        'DGV_MotTRéservés
        '
        Me.DGV_MotTRéservés.AllowUserToOrderColumns = True
        Me.DGV_MotTRéservés.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_MotTRéservés.DefaultCellStyle = DataGridViewCellStyle1
        Me.DGV_MotTRéservés.Location = New System.Drawing.Point(4, 5)
        Me.DGV_MotTRéservés.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_MotTRéservés.Name = "DGV_MotTRéservés"
        Me.DGV_MotTRéservés.RowHeadersWidth = 62
        Me.DGV_MotTRéservés.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV_MotTRéservés.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_MotTRéservés.TabIndex = 12
        '
        'Btn_MotTRéservés_MAJ
        '
        Me.Btn_MotTRéservés_MAJ.Location = New System.Drawing.Point(4, 954)
        Me.Btn_MotTRéservés_MAJ.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_MotTRéservés_MAJ.Name = "Btn_MotTRéservés_MAJ"
        Me.Btn_MotTRéservés_MAJ.Size = New System.Drawing.Size(112, 35)
        Me.Btn_MotTRéservés_MAJ.TabIndex = 11
        Me.Btn_MotTRéservés_MAJ.Text = "Mise à Jour"
        Me.Btn_MotTRéservés_MAJ.UseVisualStyleBackColor = True
        '
        'P5
        '
        Me.P5.Controls.Add(Me.Btn_MsgT_Reload)
        Me.P5.Controls.Add(Me.TB_MsgT_Input)
        Me.P5.Controls.Add(Me.DGV_MsgT)
        Me.P5.Location = New System.Drawing.Point(4, 29)
        Me.P5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P5.Name = "P5"
        Me.P5.Size = New System.Drawing.Size(1657, 1009)
        Me.P5.TabIndex = 4
        Me.P5.Text = "MsgT"
        Me.P5.UseVisualStyleBackColor = True
        '
        'Btn_MsgT_Reload
        '
        Me.Btn_MsgT_Reload.Location = New System.Drawing.Point(164, 946)
        Me.Btn_MsgT_Reload.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_MsgT_Reload.Name = "Btn_MsgT_Reload"
        Me.Btn_MsgT_Reload.Size = New System.Drawing.Size(112, 35)
        Me.Btn_MsgT_Reload.TabIndex = 17
        Me.Btn_MsgT_Reload.Text = "Recherche"
        Me.Btn_MsgT_Reload.UseVisualStyleBackColor = True
        '
        'TB_MsgT_Input
        '
        Me.TB_MsgT_Input.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.TB_MsgT_Input.Location = New System.Drawing.Point(4, 946)
        Me.TB_MsgT_Input.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TB_MsgT_Input.Name = "TB_MsgT_Input"
        Me.TB_MsgT_Input.Size = New System.Drawing.Size(148, 30)
        Me.TB_MsgT_Input.TabIndex = 16
        Me.TB_MsgT_Input.Text = "qsm"
        Me.TB_MsgT_Input.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TB_MsgT_Input.UseWaitCursor = True
        '
        'DGV_MsgT
        '
        Me.DGV_MsgT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_MsgT.DefaultCellStyle = DataGridViewCellStyle2
        Me.DGV_MsgT.Location = New System.Drawing.Point(4, 5)
        Me.DGV_MsgT.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_MsgT.Name = "DGV_MsgT"
        Me.DGV_MsgT.RowHeadersWidth = 62
        Me.DGV_MsgT.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_MsgT.TabIndex = 15
        '
        'P4
        '
        Me.P4.Controls.Add(Me.Btn_PrcT_Reload)
        Me.P4.Controls.Add(Me.TB_PrcT_Input)
        Me.P4.Controls.Add(Me.DGV_PrcT)
        Me.P4.Location = New System.Drawing.Point(4, 29)
        Me.P4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P4.Name = "P4"
        Me.P4.Size = New System.Drawing.Size(1657, 1009)
        Me.P4.TabIndex = 3
        Me.P4.Text = "PrcT"
        Me.P4.UseVisualStyleBackColor = True
        '
        'Btn_PrcT_Reload
        '
        Me.Btn_PrcT_Reload.Location = New System.Drawing.Point(164, 946)
        Me.Btn_PrcT_Reload.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_PrcT_Reload.Name = "Btn_PrcT_Reload"
        Me.Btn_PrcT_Reload.Size = New System.Drawing.Size(112, 35)
        Me.Btn_PrcT_Reload.TabIndex = 17
        Me.Btn_PrcT_Reload.Text = "Recherche"
        Me.Btn_PrcT_Reload.UseVisualStyleBackColor = True
        '
        'TB_PrcT_Input
        '
        Me.TB_PrcT_Input.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.TB_PrcT_Input.Location = New System.Drawing.Point(4, 946)
        Me.TB_PrcT_Input.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TB_PrcT_Input.Name = "TB_PrcT_Input"
        Me.TB_PrcT_Input.Size = New System.Drawing.Size(148, 30)
        Me.TB_PrcT_Input.TabIndex = 16
        Me.TB_PrcT_Input.Text = "qsm"
        Me.TB_PrcT_Input.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TB_PrcT_Input.UseWaitCursor = True
        '
        'DGV_PrcT
        '
        Me.DGV_PrcT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_PrcT.DefaultCellStyle = DataGridViewCellStyle3
        Me.DGV_PrcT.Location = New System.Drawing.Point(4, 5)
        Me.DGV_PrcT.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_PrcT.Name = "DGV_PrcT"
        Me.DGV_PrcT.RowHeadersWidth = 62
        Me.DGV_PrcT.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_PrcT.TabIndex = 15
        '
        'P3
        '
        Me.P3.Controls.Add(Me.Btn_MotT_Reload)
        Me.P3.Controls.Add(Me.TB_MotT_Input)
        Me.P3.Controls.Add(Me.DGV_MotT)
        Me.P3.Location = New System.Drawing.Point(4, 29)
        Me.P3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P3.Name = "P3"
        Me.P3.Size = New System.Drawing.Size(1657, 1009)
        Me.P3.TabIndex = 2
        Me.P3.Text = "MotT"
        Me.P3.UseVisualStyleBackColor = True
        '
        'Btn_MotT_Reload
        '
        Me.Btn_MotT_Reload.Location = New System.Drawing.Point(164, 946)
        Me.Btn_MotT_Reload.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_MotT_Reload.Name = "Btn_MotT_Reload"
        Me.Btn_MotT_Reload.Size = New System.Drawing.Size(112, 35)
        Me.Btn_MotT_Reload.TabIndex = 14
        Me.Btn_MotT_Reload.Text = "Recherche"
        Me.Btn_MotT_Reload.UseVisualStyleBackColor = True
        '
        'TB_MotT_Input
        '
        Me.TB_MotT_Input.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!)
        Me.TB_MotT_Input.Location = New System.Drawing.Point(4, 946)
        Me.TB_MotT_Input.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TB_MotT_Input.Name = "TB_MotT_Input"
        Me.TB_MotT_Input.Size = New System.Drawing.Size(148, 30)
        Me.TB_MotT_Input.TabIndex = 13
        Me.TB_MotT_Input.Text = "qsm"
        Me.TB_MotT_Input.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TB_MotT_Input.UseWaitCursor = True
        '
        'DGV_MotT
        '
        Me.DGV_MotT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_MotT.DefaultCellStyle = DataGridViewCellStyle4
        Me.DGV_MotT.Location = New System.Drawing.Point(4, 5)
        Me.DGV_MotT.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_MotT.Name = "DGV_MotT"
        Me.DGV_MotT.RowHeadersWidth = 62
        Me.DGV_MotT.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_MotT.TabIndex = 12
        '
        'P2
        '
        Me.P2.Controls.Add(Me.Btn_CodT_Reload)
        Me.P2.Controls.Add(Me.TB_CodT_Input)
        Me.P2.Controls.Add(Me.DGV_CodT)
        Me.P2.Location = New System.Drawing.Point(4, 29)
        Me.P2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P2.Name = "P2"
        Me.P2.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P2.Size = New System.Drawing.Size(1657, 1009)
        Me.P2.TabIndex = 1
        Me.P2.Text = "CodT"
        Me.P2.UseVisualStyleBackColor = True
        '
        'Btn_CodT_Reload
        '
        Me.Btn_CodT_Reload.Location = New System.Drawing.Point(164, 946)
        Me.Btn_CodT_Reload.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_CodT_Reload.Name = "Btn_CodT_Reload"
        Me.Btn_CodT_Reload.Size = New System.Drawing.Size(112, 35)
        Me.Btn_CodT_Reload.TabIndex = 11
        Me.Btn_CodT_Reload.Text = "Recherche"
        Me.Btn_CodT_Reload.UseVisualStyleBackColor = True
        '
        'TB_CodT_Input
        '
        Me.TB_CodT_Input.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TB_CodT_Input.Location = New System.Drawing.Point(4, 946)
        Me.TB_CodT_Input.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TB_CodT_Input.Name = "TB_CodT_Input"
        Me.TB_CodT_Input.Size = New System.Drawing.Size(148, 30)
        Me.TB_CodT_Input.TabIndex = 5
        Me.TB_CodT_Input.Text = "qsm"
        Me.TB_CodT_Input.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.TB_CodT_Input.UseWaitCursor = True
        '
        'DGV_CodT
        '
        Me.DGV_CodT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_CodT.DefaultCellStyle = DataGridViewCellStyle5
        Me.DGV_CodT.Location = New System.Drawing.Point(4, 5)
        Me.DGV_CodT.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_CodT.Name = "DGV_CodT"
        Me.DGV_CodT.RowHeadersWidth = 62
        Me.DGV_CodT.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_CodT.TabIndex = 0
        '
        'P1
        '
        Me.P1.Controls.Add(Me.Btn_Dictionnaire_Load)
        Me.P1.Controls.Add(Me.DGV_Dictionnaire)
        Me.P1.Controls.Add(Me.Btn_Dictionnaire_MAJ)
        Me.P1.Location = New System.Drawing.Point(4, 29)
        Me.P1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P1.Name = "P1"
        Me.P1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P1.Size = New System.Drawing.Size(1657, 1009)
        Me.P1.TabIndex = 0
        Me.P1.Text = "Dictionnaire"
        Me.P1.UseVisualStyleBackColor = True
        '
        'Btn_Dictionnaire_Load
        '
        Me.Btn_Dictionnaire_Load.Location = New System.Drawing.Point(142, 946)
        Me.Btn_Dictionnaire_Load.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_Dictionnaire_Load.Name = "Btn_Dictionnaire_Load"
        Me.Btn_Dictionnaire_Load.Size = New System.Drawing.Size(112, 35)
        Me.Btn_Dictionnaire_Load.TabIndex = 10
        Me.Btn_Dictionnaire_Load.Text = "Chargement"
        Me.Btn_Dictionnaire_Load.UseVisualStyleBackColor = True
        '
        'DGV_Dictionnaire
        '
        Me.DGV_Dictionnaire.AllowUserToOrderColumns = True
        Me.DGV_Dictionnaire.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle6.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle6.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle6.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_Dictionnaire.DefaultCellStyle = DataGridViewCellStyle6
        Me.DGV_Dictionnaire.Location = New System.Drawing.Point(4, 5)
        Me.DGV_Dictionnaire.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_Dictionnaire.Name = "DGV_Dictionnaire"
        Me.DGV_Dictionnaire.RowHeadersWidth = 62
        Me.DGV_Dictionnaire.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV_Dictionnaire.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_Dictionnaire.TabIndex = 8
        '
        'Btn_Dictionnaire_MAJ
        '
        Me.Btn_Dictionnaire_MAJ.Location = New System.Drawing.Point(4, 946)
        Me.Btn_Dictionnaire_MAJ.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Btn_Dictionnaire_MAJ.Name = "Btn_Dictionnaire_MAJ"
        Me.Btn_Dictionnaire_MAJ.Size = New System.Drawing.Size(112, 35)
        Me.Btn_Dictionnaire_MAJ.TabIndex = 7
        Me.Btn_Dictionnaire_MAJ.Text = "Mise à Jour"
        Me.Btn_Dictionnaire_MAJ.UseVisualStyleBackColor = True
        '
        'TB
        '
        Me.TB.Controls.Add(Me.P1)
        Me.TB.Controls.Add(Me.P2)
        Me.TB.Controls.Add(Me.P3)
        Me.TB.Controls.Add(Me.P4)
        Me.TB.Controls.Add(Me.P5)
        Me.TB.Controls.Add(Me.P6)
        Me.TB.Controls.Add(Me.P7)
        Me.TB.Location = New System.Drawing.Point(18, 42)
        Me.TB.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TB.Name = "TB"
        Me.TB.SelectedIndex = 0
        Me.TB.Size = New System.Drawing.Size(1665, 1042)
        Me.TB.TabIndex = 9
        '
        'P7
        '
        Me.P7.Controls.Add(Me.DGV_PrmT)
        Me.P7.Location = New System.Drawing.Point(4, 29)
        Me.P7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P7.Name = "P7"
        Me.P7.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.P7.Size = New System.Drawing.Size(1657, 1009)
        Me.P7.TabIndex = 6
        Me.P7.Text = "PrmT"
        Me.P7.UseVisualStyleBackColor = True
        '
        'DGV_PrmT
        '
        Me.DGV_PrmT.AllowUserToOrderColumns = True
        Me.DGV_PrmT.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DGV_PrmT.DefaultCellStyle = DataGridViewCellStyle7
        Me.DGV_PrmT.Location = New System.Drawing.Point(4, 5)
        Me.DGV_PrmT.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DGV_PrmT.Name = "DGV_PrmT"
        Me.DGV_PrmT.RowHeadersWidth = 62
        Me.DGV_PrmT.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DGV_PrmT.Size = New System.Drawing.Size(1644, 923)
        Me.DGV_PrmT.TabIndex = 13
        '
        'Frm_Dictionnaire
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1694, 1111)
        Me.Controls.Add(Me.TB_Information)
        Me.Controls.Add(Me.TB)
        Me.Controls.Add(Me.Mnu)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "Frm_Dictionnaire"
        Me.Text = "Frm_Dictionnaire"
        Me.Mnu.ResumeLayout(False)
        Me.Mnu.PerformLayout()
        CType(Me.BDS, System.ComponentModel.ISupportInitialize).EndInit()
        Me.P6.ResumeLayout(False)
        CType(Me.DGV_MotTRéservés, System.ComponentModel.ISupportInitialize).EndInit()
        Me.P5.ResumeLayout(False)
        Me.P5.PerformLayout()
        CType(Me.DGV_MsgT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.P4.ResumeLayout(False)
        Me.P4.PerformLayout()
        CType(Me.DGV_PrcT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.P3.ResumeLayout(False)
        Me.P3.PerformLayout()
        CType(Me.DGV_MotT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.P2.ResumeLayout(False)
        Me.P2.PerformLayout()
        CType(Me.DGV_CodT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.P1.ResumeLayout(False)
        CType(Me.DGV_Dictionnaire, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TB.ResumeLayout(False)
        Me.P7.ResumeLayout(False)
        CType(Me.DGV_PrmT, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Mnu_0190 As ToolStripMenuItem
    Friend WithEvents Mnu_01 As ToolStripMenuItem
    Friend WithEvents Mnu As MenuStrip
    Friend WithEvents TB_Information As TextBox
    Friend WithEvents BDS As BindingSource
    Friend WithEvents P6 As TabPage
    Friend WithEvents Btn_MotTRéservés_Load As Button
    Friend WithEvents DGV_MotTRéservés As DataGridView
    Friend WithEvents Btn_MotTRéservés_MAJ As Button
    Friend WithEvents P5 As TabPage
    Friend WithEvents Btn_MsgT_Reload As Button
    Friend WithEvents TB_MsgT_Input As TextBox
    Friend WithEvents DGV_MsgT As DataGridView
    Friend WithEvents P4 As TabPage
    Friend WithEvents Btn_PrcT_Reload As Button
    Friend WithEvents TB_PrcT_Input As TextBox
    Friend WithEvents DGV_PrcT As DataGridView
    Friend WithEvents P3 As TabPage
    Friend WithEvents Btn_MotT_Reload As Button
    Friend WithEvents TB_MotT_Input As TextBox
    Friend WithEvents DGV_MotT As DataGridView
    Friend WithEvents P2 As TabPage
    Friend WithEvents Btn_CodT_Reload As Button
    Friend WithEvents TB_CodT_Input As TextBox
    Friend WithEvents DGV_CodT As DataGridView
    Friend WithEvents P1 As TabPage
    Friend WithEvents Btn_Dictionnaire_Load As Button
    Friend WithEvents DGV_Dictionnaire As DataGridView
    Friend WithEvents Btn_Dictionnaire_MAJ As Button
    Friend WithEvents TB As TabControl
    Friend WithEvents P7 As TabPage
    Friend WithEvents DGV_PrmT As DataGridView
    Friend WithEvents Mnu_0101 As ToolStripMenuItem
    Friend WithEvents Mnu_0103 As ToolStripMenuItem
    Friend WithEvents Mnu_02 As ToolStripMenuItem
    Friend WithEvents Mnu_0201 As ToolStripMenuItem
    Friend WithEvents Mnu_0202 As ToolStripMenuItem
    Friend WithEvents Mnu_0203 As ToolStripMenuItem
    Friend WithEvents Mnu_0204 As ToolStripMenuItem
    Friend WithEvents Mnu_0206 As ToolStripMenuItem
    Friend WithEvents Mnu_0205 As ToolStripMenuItem
End Class
