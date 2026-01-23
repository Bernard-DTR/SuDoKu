Option Strict On
Option Explicit On

'--------------------------------------------------------------------------------------
'Date de Création: Lundi 22/08/2022
'
' x ToolTip_Text sur les contrôles
' x Cohérence Copie des candidats avec Préférences
' x Aimantation_Click
' x Insertion intelligente
' x Choix des noms (Formulaire et Variables)
' x Insertion multiple 16/04/2023 
'-------------------------------------------------------------------------------------
Public Class Frm_Insérer_Candidats
  Private Sub Frm_Insérer_Candidats_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.AutoScaleMode = Me_AutoScaleMode_Standard
    Plcy_FIC_Frm_Insérer_Candidats = True
    Me.Text = Msg_Read_IA("SDK_10010")
    BackColor = Color_Frm_BackColor
    Select Case Plcy_Slm
      Case True
        Btn_Insérer.Text = Msg_Read_IA("SDK_10071", {U_Coord(Pbl_Cell_Select)})
        TT.SetToolTip(Btn_Insérer, Msg_Read_IA("SDK_10061"))
      Case False
        Btn_Insérer.Text = Msg_Read_IA("SDK_10070", {U_Coord(Pbl_Cell_Select)})
        TT.SetToolTip(Btn_Insérer, Msg_Read_IA("SDK_10060"))
    End Select
    CB01.Text = Msg_Read_IA("SDK_10040")
    PB_G.BackColor = Color.IndianRed
    Plcy_FIC_Zone_Aimantée = "G"
    TT.SetToolTip(PB_G, Msg_Read_IA("SDK_10020"))
    TT.SetToolTip(PB_BG, Msg_Read_IA("SDK_10021"))
    TT.SetToolTip(PB_BD, Msg_Read_IA("SDK_10022"))
    TT.SetToolTip(PB_D, Msg_Read_IA("SDK_10023"))
    TT.SetToolTip(PB_A, Msg_Read_IA("SDK_10024"))
    TT.SetToolTip(Btn_TTT, Msg_Read_IA("SDK_10050"))
    TT.SetToolTip(TB_Candidats, Msg_Read_IA("SDK_10030"))
    TT.SetToolTip(CB01, Msg_Read_IA("SDK_10041"))
    If Not Plcy_MouseClick_Middle Then Btn_TTT.Enabled = False
    TB_Candidats.Font = Font_Mnu_Cel
  End Sub

  Private Sub Frm_Insérer_Candidats_LocationChanged(sender As Object, e As EventArgs) Handles MyBase.LocationChanged
    Dim Frm_Insérer_Candidats_Rect As Rectangle
    Dim Frm_SDK_Rect As Rectangle
    Frm_Insérer_Candidats_Rect = New Rectangle(x:=Location.X, y:=Location.Y, width:=Width, height:=Height)
    Frm_SDK_Rect = New Rectangle(x:=Frm_SDK.Location.X, y:=Frm_SDK.Location.Y, width:=Frm_SDK.Width, height:=Frm_SDK.Height)
    Dim Pt_GH As New Point(x:=Location.X, y:=Location.Y)
    Dim Pt_DH As New Point(x:=Location.X + Width, y:=Location.Y)
    Dim Pt_GB As New Point(x:=Location.X, y:=Location.Y + Height)
    Dim Pt_DB As New Point(x:=Location.X + Width, y:=Location.Y + Height)
    Dim Q As Integer = 0
    If Not Frm_SDK_Rect.Contains(Pt_GH) Then Q += 1
    If Not Frm_SDK_Rect.Contains(Pt_DH) Then Q += 1
    If Not Frm_SDK_Rect.Contains(Pt_GB) Then Q += 1
    If Not Frm_SDK_Rect.Contains(Pt_DB) Then Q += 1
    If Q = 4 Then
      Dés_Aimantation()
    End If
    Select Case Plcy_Slm
      Case True
        Btn_Insérer.Text = Msg_Read_IA("SDK_10071", {U_Coord(Pbl_Cell_Select)})
        TT.SetToolTip(Btn_Insérer, Msg_Read_IA("SDK_10061"))
      Case False
        Btn_Insérer.Text = Msg_Read_IA("SDK_10070", {U_Coord(Pbl_Cell_Select)})
        TT.SetToolTip(Btn_Insérer, Msg_Read_IA("SDK_10060"))
    End Select
  End Sub
  Private Sub Insérer_Click(sender As Object, e As EventArgs) Handles Btn_Insérer.Click
    Dim L As Integer = TB_Candidats.TextLength
    If L > 9 Then L = 9
    For j As Integer = 0 To L - 1
      Dim Cdd As String = TB_Candidats.Text.Substring(j, 1)
      Select Case Plcy_Slm
        Case True
          For i As Integer = 0 To 80
            If U_Slm(i) Then
              Dim Cel As Integer = i
              If Cdd >= "1" And Cdd <= "9" Then
                Select Case CB01.Checked
                  Case True
                    If Cell_Coll_Val_Check(U, Cdd, Cel) = -1 Then
                      Cell_Cdd_Insert(Cdd, Cel, "_Cdd Mm")
                    End If
                  Case False
                    Cell_Cdd_Insert(Cdd, Cel, "_Cdd Mm")
                End Select
              End If
            End If
          Next i
        Case False
          Dim Cel As Integer = Pbl_Cell_Select
          If Cdd >= "1" And Cdd <= "9" Then
            Select Case CB01.Checked
              Case True
                If Cell_Coll_Val_Check(U, Cdd, Cel) = -1 Then
                  Cell_Cdd_Insert(Cdd, Cel, "_Cdd Mm")
                End If
              Case False
                Cell_Cdd_Insert(Cdd, Cel, "_Cdd Mm")
            End Select
          End If
      End Select
    Next j
  End Sub
  Private Sub Btn_TTT_Click(sender As Object, e As EventArgs) Handles Btn_TTT.Click
    TB_Candidats.Text = Plcy_FIC_TTT
  End Sub
  Private Sub Frm_Insérer_Candidats_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
    Plcy_FIC_Frm_Insérer_Candidats = False
  End Sub
  Private Sub Aimantation(sender As Object, e As MouseEventArgs) Handles PB_G.MouseClick, PB_BG.MouseClick, PB_D.MouseClick, PB_BD.MouseClick, PB_A.MouseClick
    Dim Cntl As Control = CType(sender, Control)
    PB_G.BackColor = Color.Beige
    PB_BG.BackColor = Color.Beige
    PB_BD.BackColor = Color.Beige
    PB_D.BackColor = Color.Beige
    PB_A.BackColor = Color.Beige
    Select Case Cntl.Name
      Case "PB_G" : PB_G.BackColor = Color.IndianRed : Plcy_FIC_Zone_Aimantée = "G"
      Case "PB_BG" : PB_BG.BackColor = Color.IndianRed : Plcy_FIC_Zone_Aimantée = "BG"
      Case "PB_BD" : PB_BD.BackColor = Color.IndianRed : Plcy_FIC_Zone_Aimantée = "BD"
      Case "PB_D" : PB_D.BackColor = Color.IndianRed : Plcy_FIC_Zone_Aimantée = "D"
      Case "PB_A" : PB_A.BackColor = Color.IndianRed : Plcy_FIC_Zone_Aimantée = "A"
    End Select
  End Sub
  Sub Dés_Aimantation()
    PB_G.BackColor = Color.Beige
    PB_BG.BackColor = Color.Beige
    PB_BD.BackColor = Color.Beige
    PB_D.BackColor = Color.Beige
    PB_A.BackColor = Color.IndianRed
    Plcy_FIC_Zone_Aimantée = "A"
  End Sub
End Class