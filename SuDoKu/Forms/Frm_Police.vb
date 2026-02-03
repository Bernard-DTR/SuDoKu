Option Strict On
Option Explicit On

Imports System.Drawing.Text

Public Class Frm_Police

  Private Colonne, Rangée As Integer
  Private Police_Font As Font
  Private Police_Name As String
  Private Police_Index As Integer
  Private Police_Size As Single = 12
  Private Police_Style As FontStyle = FontStyle.Regular

  Private Sub Frm_Police_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.AutoScaleMode = Me_AutoScaleMode_Standard
    DGV.Width = 62 * 16
    DGV.Height = 62 * 8

    DGV.ColumnHeadersVisible = False
    DGV.RowHeadersVisible = False

    For Colonne = 0 To 15
      Dim Col As New DataGridViewTextBoxColumn With {.Width = 60}
      DGV.Columns.Add(Col)
      DGV.Columns(Colonne).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
    Next Colonne

    For Rangée = 0 To 127
      DGV.Rows.Add()
      DGV.Rows(Rangée).Height = 60
    Next Rangée

    InstalledFontFamily = InstalledFontCollection.Families
    CB_Police_Installed.Items.Clear()
    Dim nb As Integer = InstalledFontFamily.Count
    Jrn_Add(, {"InstalledFontFamily"})
    Dim ji As Integer
    While ji < nb
      CB_Police_Installed.Items.Add(InstalledFontFamily(ji).Name)
      Jrn_Add(, {CStr(ji).PadLeft(5) & " " & InstalledFontFamily(ji).Name})
      ji += 1
    End While

    PrivateFontFamily = PrivateFontCollection.Families
    Dim jp As Integer
    nb = PrivateFontFamily.Length
    Jrn_Add(, {"PrivateFontFamily"})

    CB_Police_Private.Items.Clear()
    While jp < nb
      CB_Police_Private.Items.Add(PrivateFontFamily(jp).Name)
      Jrn_Add(, {CStr(jp).PadLeft(5) & " " & PrivateFontFamily(jp).Name})
      jp += 1
    End While

  End Sub

  Private Sub CB_Police_Installed_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Police_Installed.SelectedIndexChanged
    Police_Index = CB_Police_Installed.SelectedIndex
    Police_Name = CB_Police_Installed.SelectedItem.ToString()
    Police_Font = New Font(family:=InstalledFontFamily(Police_Index), emSize:=Police_Size, style:=Police_Style)
    Police_Display()
  End Sub

  Private Sub CB_Police_Private_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Police_Private.SelectedIndexChanged
    Police_Index = CB_Police_Private.SelectedIndex
    Police_Name = CB_Police_Private.SelectedItem.ToString()
    Police_Font = New Font(family:=PrivateFontFamily(Police_Index), emSize:=Police_Size, style:=Police_Style)
    Police_Display()
  End Sub

  Private Sub CB01_Size_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Size.SelectedIndexChanged
    Police_Size = CSng(CB_Size.SelectedItem.ToString())
    Police_Display()
  End Sub

  Private Sub CB_Style_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CB_Style.SelectedIndexChanged
    Select Case CB_Style.SelectedItem.ToString()
      Case "Gras" : Police_Style = FontStyle.Bold                  '1
      Case "Italique" : Police_Style = FontStyle.Italic            '2
      Case "Normal" : Police_Style = FontStyle.Regular             '0
      Case "Barré" : Police_Style = FontStyle.Strikeout            '8
      Case "Souligné" : Police_Style = FontStyle.Underline         '4
      Case Else : Police_Style = FontStyle.Regular
    End Select
    Police_Display()
  End Sub

  Private Sub Btn_Charmap_Click(sender As Object, e As EventArgs) Handles Btn_Charmap.Click
    Dim Shell_St As String = "Charmap.exe"
    Try
      Nsd_i = Shell(Shell_St, AppWinStyle.NormalFocus)
    Catch ex As Exception
      Nsd_i = MsgBox("L'application ci-dessous doit être installée." & vbCrLf & Shell_St)
    End Try
  End Sub

  Private Sub Btn_Fermer_Click(sender As Object, e As EventArgs) Handles Btn_Fermer.Click
    Me.Close()
  End Sub

  Private Sub Police_Display()
    DGV.Font = Police_Font
    Jrn_Add(, {Proc_Name_Get() & " " & DGV.Font.ToString()})
    Jrn_Add(, {"Index  : " & " " & Police_Index})
    Jrn_Add(, {"Nom    : " & " " & Police_Name})
    Jrn_Add(, {"Size   : " & " " & Police_Size})
    Jrn_Add(, {"Style  : " & " " & Police_Style})
    For i As Integer = 0 To 2048 '256
      'Chr(i) est limité à 256
      Rangée = (i \ 16)
      Colonne = (i - ((i \ 16) * 16))
      Try
        DGV(Colonne, Rangée).Value = CStr(ChrW(i))
        DGV(Colonne, Rangée).ToolTipText = CStr(i).PadRight(5) & " Numéro" & vbCrLf &
                                           CStr(Chr(i)).PadRight(5) & " Caractère" & vbCrLf &
                                           CStr(Hex(i)).PadRight(5) & " Hexadécimale" & vbCrLf &
                                           CStr(Asc(Chr(i))).PadRight(5) & " ASCII"
      Catch ex As Exception
        'Exit For
        Continue For
      End Try
    Next i
  End Sub
End Class