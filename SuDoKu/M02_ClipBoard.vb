Imports System.Drawing.Imaging
Friend Module M02_ClipBoard
  '-------------------------------------------------------------------------------
  ' Traitement des Copier-Coller
  '-------------------------------------------------------------------------------
  Public Sub ClipBoard_Coller()
    '-----------------------------------------------------------------------
    ' OK    Copier de Angus Johnson
    ' OK    N'existe pas pour Coles      
    ' OK    Copier les grilles de Mots-Croisés Internet
    ' OK    Copier les grilles de SudoCue 05/02/2025
    '-----------------------------------------------------------------------
    Dim Prb As String
    Prb = Clipboard.GetText(TextDataFormat.Text)
    Try
      If Mid$(Prb, 1, 14) = " *-----------*" Then
        'Origine Angus Johnson
        Prb = Prb.Replace(" *", "")
        Prb = Prb.Replace(" |", "")
        Prb = Prb.Replace("*", "")
        Prb = Prb.Replace("-", "")
        Prb = Prb.Replace("|", "")
        Prb = Prb.Replace("+", "")
        Prb = Prb.Replace(".", " ")
        Prb = Prb.Replace(vbCrLf, "")
      Else
        Prb = Prb.Replace(vbCrLf, "")
        Prb = Prb.Replace(vbTab, "")
        Prb = Prb.Replace("-", "")    ' SudoCue
        Prb = Prb.Replace("|", "")    ' SudoCue
        Prb = Prb.Replace("+", "")    ' SudoCue
        Prb = Prb.Replace(" ", "")    ' SudoCue
      End If

      Prb = Prb.Substring(0, 81)
      Game_New_Game(Gnrl:="Nrm", "   ", Nom:=Proc_Name_Get(), Prb:=Prb, Jeu:=Prb, Sol:=StrDup(81, " "), Cdd729:=StrDup(729, " "), Frc:="5", Proc_Name_Get())
    Catch Ex As Exception
      Dim Msg As String = "Le ClipBoard Coller n'est pas exploitable." & vbCrLf & Ex.ToString() & vbCrLf & Len(Ex.ToString())
      Nsd_i = MsgBox(Msg & vbCrLf & Prb,, Proc_Name_Get())
      Exit Sub
    End Try
  End Sub
  Public Sub ClipBoard_Coller_RTF()
    'Colle la grille sous forme graphique dans le journal RTF
    'Accessible par Ctrl + P, le PP n'est pas effacé 
    ' La propriété AutoScaleMode =  Font pour tous les Forms de SDK
    'HP Omen 16 a une résolution (recommandée) de  2560 x 1440 
    '           or la résolution indiquée est de   1707 x 960
    ' et une mise à l'échelle de 150% également recommandée
    'BENQ Les résolutions physiques et trouvées sont les mêmes
    Jrn_Add(, {Proc_Name_Get() & " " & LP_Nom})
    ' Sur quel écran suis-je ?
    Dim Screens As Screen() = Screen.AllScreens
    Dim Device_Number As Integer = 1 ' Concerne le BENQ
    'La position du TopLeft de Frm_SDK indique si l'application se trouve sur HPOmen16 ou BENQ
    If Frm_SDK.Left <= Screens(0).Bounds.Width Then Device_Number = 0 'Concerne HPOmen16
    Dim Device_Name As String = Screens(Device_Number).DeviceName

    Try
      'Gz_tl  est le point Top_Left dans le formulaire
      'Screen_TopLeft est le point Top_Left dans l'écran
      'PointToScreen convertie le point du formulaire en point sur l'écran
      Dim Screen_TopLeft As Point = Frm_SDK.PointToScreen(Gz_tl)
      Dim W_Grid As Integer
      Dim H_Grid As Integer
      ' La taill de l'image semble la même qqs WH que ce soit sur HPOmen16 ou BENQ
      Screen_TopLeft.X = CInt(Screen_TopLeft.X * Get_Scale(Device_Number).X)
      Screen_TopLeft.Y = CInt(Screen_TopLeft.Y * Get_Scale(Device_Number).Y)
      W_Grid = CInt(Bld_WH_Grid * Get_Scale(Device_Number).X)
      H_Grid = CInt(Bld_WH_Grid * Get_Scale(Device_Number).Y)
      ' Le format est de 32 bits par pixel. Les composants ARGB ont 8 bits chacun
      Try
        Using bmp As New Bitmap(W_Grid, H_Grid, PixelFormat.Format32bppArgb)
          Using g As Graphics = Graphics.FromImage(bmp)
            ' The source area is copied directly to the destination area.
            g.CopyFromScreen(Screen_TopLeft.X, Screen_TopLeft.Y, 0, 0, New Size(W_Grid, H_Grid), CopyPixelOperation.SourceCopy)
          End Using
          Clipboard.SetImage(bmp) ' Copiez le Bitmap dans le presse-papier
          'Remplace la sélection active de la zone de texte par le contenu du Presse-papiers.
          Frm_SDK.Journal.SelectionStart = Frm_SDK.Journal.Text.Length
          Frm_SDK.Journal.ReadOnly = False
          Frm_SDK.Journal.Paste()
          Frm_SDK.Journal.ReadOnly = True
        End Using
      Catch ex As Exception
        Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      End Try
      Jrn_Add("SDK_Space")
    Catch ex As System.Runtime.InteropServices.ExternalException
      Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      Nsd_i = MsgBox("Presse-Papier inaccessible. Please try again.",, MsgTit)
    End Try
  End Sub

  Public Function ClipBoard_Copier(Cel_Type As String) As String
    '-----------------------------------------------------------------------
    ' OK    Coller dans Angus Johnson 
    '              AJ accepte Espace ou point pour une valeur absente
    ' OK    Coller dans Coles à partir de Game/acTestGrid
    '-----------------------------------------------------------------------
    Dim sb As New System.Text.StringBuilder(900)

    Select Case Cel_Type
      Case "1"   ' Valeurs initiales
        For i As Integer = 0 To 80
          sb.Append(If(U(i, 1) = " ", ".", U(i, 1)))
        Next

      Case "2"   ' Valeurs trouvées
        For i As Integer = 0 To 80
          sb.Append(If(U(i, 2) = " ", ".", U(i, 2)))
        Next

      Case "3"   ' Candidats
        For i As Integer = 0 To 80
          sb.Append(U(i, 3).Replace(" ", ""))
          sb.Append(";"c)
        Next
    End Select

    Dim CC As String = sb.ToString()

    Clipboard.Clear()
    Clipboard.SetText(CC)

    Return Clipboard.GetText()
  End Function

End Module