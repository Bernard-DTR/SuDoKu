Imports System.Data
Imports System.Data.SqlClient

' DGV           DataGridView
' BDS           BindingSource
'               Composant conçu pour simplifier la liaison du DGV à une source de données 
Public Class Frm_Dictionnaire
  Private ReadOnly SQL_Connect As String = My.Settings.Connect
  Private SQL_DataAdaptateur As New SqlDataAdapter()
  Private ReadOnly Sql_Qt As String = "'"

  Private Sub Frm_Dictionnaire_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.BackColor = Color_Frm_BackColor
    Me.Text = Application.ProductName & "_" & My.Settings.Date_Extraction
    P1_Click(sender, e)
    TB_CodT_Input.Text = My.Settings.Sql_TB_CodT_Input
    TB_Information.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
    AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    ' La taille précisée dans le Design de Frm_Dictionnaire est 1716; 1167
    ' c'est-à-dire 1144 + 1144/2 = 1716
    '               778 +  778/2 = 1167
    Width = 1144
    Height = 778
  End Sub
  Private Sub TB_Click(sender As Object, e As EventArgs) Handles TB.Click
    Select Case TB.SelectedIndex
      Case 0 : P1_Click(sender, e)
      Case 1 : P2_Click(sender, e)
      Case 2 : P3_Click(sender, e)
      Case 3 : P4_Click(sender, e)
      Case 4 : P5_Click(sender, e)
      Case 5 : P6_Click(sender, e)
      Case 6 : P7_Click(sender, e)
    End Select
  End Sub
  Private Sub P1_Click(sender As Object, e As EventArgs) Handles P1.Click
    'Concerne le Dictionnaire
    Dim W_Col() As String = Split(My.Settings.DGV_Dictionnaire_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_Dictionnaire
      .AllowUserToAddRows = True
      .AllowUserToDeleteRows = True
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "Select * FROM Dictionnaire Order by Famille, Nom, Explication "
      Get_Data(Sql_Select)
      TB_Information.Text = "Dictionnaire: " & CStr(DGV_Dictionnaire.RowCount) & " enregistrements; " & Sql_Select
      'Il faut attendre que les colonnes soient retournées
      Get_DGV_Alignment(DGV_Dictionnaire)
      Get_DGV_ColWidth(DGV_Dictionnaire, W_Col)
    End With
  End Sub
  Private Sub P2_Click(sender As Object, e As EventArgs) Handles P2.Click
    'Concerne le Code
    Dim W_Col() As String = Split(My.Settings.DGV_CodT_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_CodT
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "SELECT Top 100 * FROM CodT Order By IdCode"
      Get_Data(Sql_Select)
      TB_Information.Text = "CodT: " & CStr(DGV_CodT.RowCount) & " enregistrements; " & Sql_Select
      Get_DGV_Alignment(DGV_CodT)
      Get_DGV_ColWidth(DGV_CodT, W_Col)
    End With
    TB_CodT_Input.Text = My.Settings.Sql_TB_CodT_Input
  End Sub
  Private Sub P3_Click(sender As Object, e As EventArgs) Handles P3.Click
    'Concerne les Mots
    Dim W_Col() As String = Split(My.Settings.DGV_MotT_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_MotT
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "SELECT * FROM MotV Order by Nombre Desc, Mot Asc "
      Get_Data(Sql_Select)
      TB_Information.Text = "MotV: " & CStr(DGV_MotT.RowCount) & " enregistrements; " & Sql_Select
      Get_DGV_Alignment(DGV_MotT)
      Get_DGV_ColWidth(DGV_MotT, W_Col)
    End With
    TB_MotT_Input.Text = My.Settings.Sql_TB_MotT_Input
  End Sub
  Private Sub P4_Click(sender As Object, e As EventArgs) Handles P3.Click
    'Concerne les Noms des Procédures
    Dim W_Col() As String = Split(My.Settings.DGV_PrcT_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_PrcT
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "SELECT * FROM PrcT Order by Nombre Desc , Procédure Asc "
      Get_Data(Sql_Select)
      TB_Information.Text = "PrcT: " & CStr(DGV_PrcT.RowCount) & " enregistrements; " & Sql_Select
      Get_DGV_Alignment(DGV_PrcT)
      Get_DGV_ColWidth(DGV_PrcT, W_Col)
    End With
    TB_PrcT_Input.Text = My.Settings.Sql_TB_PrcT_Input
  End Sub
  Private Sub P5_Click(sender As Object, e As EventArgs) Handles P3.Click
    'Concerne les Messages
    Dim W_Col() As String = Split(My.Settings.DGV_MsgT_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_MsgT
      .AllowUserToAddRows = False
      .AllowUserToDeleteRows = False
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "SELECT * FROM MsgT Order By MsgId Asc"
      Get_Data(Sql_Select)
      TB_Information.Text = "MsgT: " & CStr(DGV_MsgT.RowCount) & " enregistrements; " & Sql_Select
      Get_DGV_Alignment(DGV_MsgT)
      Get_DGV_ColWidth(DGV_MsgT, W_Col)
      TB_MsgT_Input.Text = My.Settings.Sql_TB_MsgT_Input
    End With
  End Sub
  Private Sub P6_Click(sender As Object, e As EventArgs) Handles P6.Click
    'Concerne les Mots Réservés
    Dim W_Col() As String = Split(My.Settings.DGV_MotTRéservés_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_MotTRéservés
      .AllowUserToAddRows = True
      .AllowUserToDeleteRows = True
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "Select * FROM MotTRéservés Order by MotR Asc "
      Get_Data(Sql_Select)
      TB_Information.Text = "MotTRéservés: " & CStr(DGV_MotTRéservés.RowCount) & " enregistrements; " & Sql_Select
      'Il faut attendre que les colonnes soient retournées
      Get_DGV_Alignment(DGV_MotTRéservés)
      Get_DGV_ColWidth(DGV_MotTRéservés, W_Col)
    End With
  End Sub
  Private Sub P7_Click(sender As Object, e As EventArgs) Handles P7.Click
    'Concerne les paramètres
    Dim W_Col() As String = Split(My.Settings.DGV_PrmT_Width, ";")
    Dim i As Integer = -1
    With Me.DGV_PrmT
      .AllowUserToAddRows = True
      .AllowUserToDeleteRows = True
      .AllowUserToOrderColumns = True
      .AllowUserToResizeColumns = True
      .AllowUserToResizeRows = False
      .AllowDrop = False
      .RowsDefaultCellStyle.BackColor = Color.Beige
      .AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(192, 255, 192)
      .DataSource = BDS
      Dim Sql_Select As String = "SELECT * FROM PrmT Order by Paramètre Asc"
      Get_Data(Sql_Select)
      TB_Information.Text = "Paramètres: " & CStr(DGV_PrmT.RowCount) & " enregistrements; " & Sql_Select
      'Il faut attendre que les colonnes soient retournées
      Get_DGV_Alignment(DGV_PrmT)
      Get_DGV_ColWidth(DGV_PrmT, W_Col)
    End With
  End Sub
  Private Sub Get_DGV_ColWidth(DGV As DataGridView, DGV_With() As String)
    'Chargement des largeurs de colonnes
    Dim i As Integer = -1
    Try
      For Each Col As Object In DGV.Columns
        Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
        i += 1
        If DGV_With(i) = "" Then DGV_With(i) = "100"
        Ccol.Width = CInt(DGV_With(i))
      Next Col
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Private Sub Get_DGV_Alignment(DGV As DataGridView)
    'Alignement des colonnes
    Dim i As Integer = -1
    For Each Col As Object In DGV.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      i += 1
      Select Case Ccol.ValueType
        Case GetType(Decimal), GetType(Int32)
          DGV.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Case GetType(String)
          DGV.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
        Case Else
          Jrn_Add("ERR_00000", {"Typologie du Champ non répertoriée : "}, "Erreur")
          Jrn_Add("ERR_00000", {DGV.Name & " " & Ccol.Name & " C " & Ccol.ValueType.ToString()}, "Erreur")
          Jrn_Add("ERR_00000", {CStr(i) & " " & DGV.Columns(i).Name}, "Erreur")
          DGV.Columns(i).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
      End Select
    Next Col
  End Sub
  Private Sub Get_Data(Cmd_Select As String)
    'Fonctionne pour l'Ensemble des Tables, lecture des données 
    Try
      SQL_DataAdaptateur = New SqlDataAdapter(Cmd_Select, SQL_Connect)
      Dim CMD_Builder As New SqlCommandBuilder(SQL_DataAdaptateur)
      Dim Table As New DataTable With {.Locale = Globalization.CultureInfo.InvariantCulture}
      Nsd_i = SQL_DataAdaptateur.Fill(Table)
      BDS.DataSource = Table
    Catch ex As SqlException
      Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      Nsd_i = MsgBox("Chaîne de connection incorrecte" & vbCrLf & SQL_Connect & vbCrLf & Cmd_Select,, MsgTit)
    End Try
  End Sub
  Private Sub DGV_Dictionnaire_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_Dictionnaire.ColumnWidthChanged
    'Enregistrement des largeurs de colonnes
    Dim W As String = ""
    For Each Col As Object In DGV_Dictionnaire.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_Dictionnaire_Width = W
  End Sub
  Private Sub Btn_Dictionnaire_Update_Click(sender As Object, e As EventArgs) Handles Btn_Dictionnaire_MAJ.Click
    ' Met à jour les valeurs de la base de données en exécutant les instructions
    ' INSERT, UPDATE ou DELETE respectives pour chaque ligne insérée, mise à jour ou supprimée
    Nsd_i = SQL_DataAdaptateur.Update(CType(BDS.DataSource, DataTable))
    TB_Information.Text = "Dictionnaire: Mise à jour effectuée."
  End Sub
  Private Sub Btn_Dictionnaire_Load_Click(sender As Object, e As EventArgs) Handles Btn_Dictionnaire_Load.Click
    ' Ré-affiche les données
    P1_Click(sender, e)
    TB_Information.Text = "Dictionnaire: Données rechargées." & SQL_DataAdaptateur.SelectCommand.CommandText
  End Sub
  Private Sub DGV_CodT_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_CodT.ColumnWidthChanged
    Dim W As String = ""
    For Each Col As Object In DGV_CodT.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_CodT_Width = W
  End Sub
  Private Sub Btn_CodT_Reload_Click(sender As Object, e As EventArgs) Handles Btn_CodT_Reload.Click
    Cursor.Current = Cursors.WaitCursor
    Dim Sql_Select As String = "SELECT * FROM CodT Where Ligne Like " & Sql_Qt & "%" & TB_CodT_Input.Text & "%" & Sql_Qt & " Order by Module, Procédure "
    Get_Data(Sql_Select)
    TB_Information.Text = "CodT : " & CStr(DGV_CodT.RowCount + 1) & " enregistrements; " & SQL_DataAdaptateur.SelectCommand.CommandText
    My.Settings.Sql_TB_CodT_Input = TB_CodT_Input.Text
    Cursor.Current = Cursors.Default
  End Sub
  Private Sub DGV_MotT_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_MotT.ColumnWidthChanged
    Dim W As String = ""
    For Each Col As Object In DGV_MotT.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_MotT_Width = W
  End Sub
  Private Sub Btn_MotT_Reload_Click(sender As Object, e As EventArgs) Handles Btn_MotT_Reload.Click
    Cursor.Current = Cursors.WaitCursor
    Dim Sql_Select As String = "SELECT * FROM MotV Where Mot Like " & Sql_Qt & "%" & TB_MotT_Input.Text & "%" & Sql_Qt & " Order by Nombre, Mot "
    Get_Data(Sql_Select)
    TB_Information.Text = "MotT : " & CStr(DGV_MotT.RowCount + 1) & " enregistrements; " & SQL_DataAdaptateur.SelectCommand.CommandText
    My.Settings.Sql_TB_MotT_Input = TB_MotT_Input.Text
    Cursor.Current = Cursors.Default
  End Sub
  Private Sub DGV_PrcT_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_PrcT.ColumnWidthChanged
    Dim W As String = ""
    For Each Col As Object In DGV_PrcT.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_PrcT_Width = W
  End Sub
  Private Sub Btn_PrcT_Reload_Click(sender As Object, e As EventArgs) Handles Btn_PrcT_Reload.Click
    Cursor.Current = Cursors.WaitCursor
    Dim Sql_Select As String = "SELECT * FROM PrcT Where Procédure Like " & Sql_Qt & "%" & TB_PrcT_Input.Text & "%" & Sql_Qt & " Order by Nombre, Procédure "
    Get_Data(Sql_Select)
    TB_Information.Text = "PrcT : " & CStr(DGV_PrcT.RowCount + 1) & " enregistrements; " & SQL_DataAdaptateur.SelectCommand.CommandText
    My.Settings.Sql_TB_PrcT_Input = TB_PrcT_Input.Text
    Cursor.Current = Cursors.Default
  End Sub
  Private Sub DGV_MsgT_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_MsgT.ColumnWidthChanged
    Dim W As String = ""
    For Each Col As Object In DGV_MsgT.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_MsgT_Width = W
  End Sub
  Private Sub Btn_MsgT_Reload_Click(sender As Object, e As EventArgs) Handles Btn_MsgT_Reload.Click
    Cursor.Current = Cursors.WaitCursor
    'SELECT IdMsg, MsgId, Msg, Nombre FROM MsgT
    Dim Sql_Select As String = "SELECT * FROM MsgT Where Msg Like " & Sql_Qt & "%" & TB_MsgT_Input.Text & "%" & Sql_Qt & " Order by Nombre, MsgId "
    Get_Data(Sql_Select)
    TB_Information.Text = "MsgT : " & CStr(DGV_MsgT.RowCount + 1) & " enregistrements; " & SQL_DataAdaptateur.SelectCommand.CommandText
    My.Settings.Sql_TB_MsgT_Input = TB_MsgT_Input.Text
    Cursor.Current = Cursors.Default
  End Sub
  Private Sub DGV_MotTRéservés_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_MotTRéservés.ColumnWidthChanged
    Dim W As String = ""
    For Each Col As Object In DGV_MotTRéservés.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_MotTRéservés_Width = W
  End Sub
  Private Sub Btn_MotTRéservés_MAJ_Click(sender As Object, e As EventArgs) Handles Btn_MotTRéservés_MAJ.Click
    ' Met à jour les valeurs de la base de données en exécutant les instructions
    ' INSERT, UPDATE ou DELETE respectives pour chaque ligne insérée, mise à jour ou supprimée
    Nsd_i = SQL_DataAdaptateur.Update(CType(BDS.DataSource, DataTable))
    TB_Information.Text = "MotTRéservés: Mise à jour effectuée."
  End Sub
  Private Sub Btn_MotTRéservés_Load_Click(sender As Object, e As EventArgs) Handles Btn_MotTRéservés_Load.Click
    ' Ré-affiche les données
    P6_Click(sender, e)
    TB_Information.Text = "MotTRéservés: Données rechargées." & SQL_DataAdaptateur.SelectCommand.CommandText
  End Sub
  Private Sub DGV_PrmT_ColumnWidthChanged(sender As Object, e As DataGridViewColumnEventArgs) Handles DGV_PrmT.ColumnWidthChanged
    Dim W As String = ""
    For Each Col As Object In DGV_PrmT.Columns
      Dim Ccol As DataGridViewColumn = CType(Col, DataGridViewColumn)
      W &= Ccol.Width & ";"
    Next Col
    My.Settings.DGV_PrmT_Width = W
  End Sub

#Region "Menu"
  Private Sub Mnu_0101_Click(sender As Object, e As EventArgs) Handles Mnu_0101.Click
    Frm_SystemInfoBrowser.Show()
  End Sub
  Private Sub Mnu_0103_Click(sender As Object, e As EventArgs) Handles Mnu_0103.Click
    Frm_MySettings.Show()
  End Sub
  Private Sub Mnu_0190_Click(sender As Object, e As EventArgs) Handles Mnu_0190.Click
    Me.Close()
  End Sub
  Private Sub Mnu_0201_Click(sender As Object, e As EventArgs) Handles Mnu_0201.Click
    Wh_Nb_Lignes_App(Path_SDK & "SuDoKu\SuDoKu")
    Jrn_Add("SDK_Space")
  End Sub
  Private Sub Mnu_0202_Click(sender As Object, e As EventArgs) Handles Mnu_0202.Click
    'Nsd_i = Shell("Notepad " & Path_SDK & "SuDoKu\SuDoKu\Apriori\aMaintenance.txt", AppWinStyle.MaximizedFocus)
    'Utilisation d'une chaîne interpolée pour éviter les & ou les +
    ' le symbole $ est utilisé en début de chaîne et les expressions sont placées {}
    Nsd_i = Shell($"Notepad {Path_SDK}SuDoKu\SuDoKu\Apriori\aMaintenance.txt", AppWinStyle.MaximizedFocus)
    SendKeys.Send("^{END}")
  End Sub
  Private Sub Mnu_0203_Click(sender As Object, e As EventArgs) Handles Mnu_0203.Click
    Frm_Police.Show()
  End Sub
  Private Sub Mnu_0204_Click(sender As Object, e As EventArgs) Handles Mnu_0204.Click
    Jrn_Add(, {Proc_Name_Get()})
    Dim Nb As Integer = 0
    Dim processes() As Process = Process.GetProcesses()
    For Each proc As Process In processes
      Nb += 1
      Try
        Dim processName As String = proc.ProcessName
        Dim executablePath As String = proc.MainModule.FileName
        Jrn_Add(, {Nb.ToString()})
        Jrn_Add(, {"Nom       : " & processName})
        Jrn_Add(, {"Programme : " & executablePath})
      Catch ex As Exception
        ' Certains processus système ne permettent pas l'accès à MainModule
        Jrn_Add(, {"Nom       : " & proc.ProcessName & " → Accès refusé"})
      End Try
    Next proc
  End Sub
  Private Sub Mnu_0205_Click(sender As Object, e As EventArgs) Handles Mnu_0205.Click
    Jrn_Add(, {Proc_Name_Get()})
    Stg_List_Display()
  End Sub
  Private Sub Mnu_0206_Click(sender As Object, e As EventArgs) Handles Mnu_0206.Click
    Jrn_Add(, {Proc_Name_Get()})
    Sql_SuDoKu()
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
#End Region
End Class