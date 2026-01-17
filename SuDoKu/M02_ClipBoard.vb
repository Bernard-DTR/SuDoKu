Option Strict On
Option Explicit On
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
      Game_New_Game(Gnrl:="Nrm", Nom:=Procédure_Name_Get(), Prb:=Prb, Jeu:=Prb, Sol:=StrDup(81, " "), Cdd729:=StrDup(729, " "), Frc:="5")
    Catch Ex As Exception
      Dim Msg As String = "Le ClipBoard Coller n'est pas exploitable." & vbCrLf & Ex.ToString() & vbCrLf & Len(Ex.ToString())
      Nsd_i = MsgBox(Msg & vbCrLf & Prb,, Procédure_Name_Get())
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
    Jrn_Add(, {Procédure_Name_Get() & " " & LP_Nom})
    'Question N° 1 Sur quel écran suis-je ?
    Dim Screens As Screen() = Screen.AllScreens
    Dim Device_Number As Integer = 1 ' Concerne le BENQ
    'La position du TopLeft de Frm_SDK indique si l'application se trouve sur HPOmen16 ou BENQ
    If Frm_SDK.Left <= Screens(0).Bounds.Width Then Device_Number = 0 'Concerne HPOmen16
    Dim Device_Name As String = Screens(Device_Number).DeviceName

    Try
      'Gz_Pt_TopLeft  est le point Top_Left dans le formulaire
      'Screen_TopLeft est le point Top_Left dans l'écran
      'PointToScreen convertie le point du formulaire en point sur l'écran
      Dim Screen_TopLeft As Point = Frm_SDK.PointToScreen(Gz_Pt_TopLeft)
      Dim W_Grid As Integer = Bld_WH_Grid
      Dim H_Grid As Integer = Bld_WH_Grid

      Screen_TopLeft.X = CInt(Screen_TopLeft.X * Get_Scale_IA(Device_Number).X)
      Screen_TopLeft.Y = CInt(Screen_TopLeft.Y * Get_Scale_IA(Device_Number).Y)
      W_Grid = CInt(Bld_WH_Grid * Get_Scale_IA(Device_Number).X)
      H_Grid = CInt(Bld_WH_Grid * Get_Scale_IA(Device_Number).Y)
      ' Le format est de 32 bits par pixel. Les composants ARGB ont 8 bits chacun
      Try
        Dim Bmp As New Bitmap(W_Grid, H_Grid, PixelFormat.Format32bppArgb)
        ' Utilisez Graphics pour effectuer la capture d'écran
        Using g As Graphics = Graphics.FromImage(Bmp)
          ' The source area is copied directly to the destination area.
          g.CopyFromScreen(Screen_TopLeft.X, Screen_TopLeft.Y, 0, 0, New Size(W_Grid, H_Grid), CopyPixelOperation.SourceCopy)
        End Using
        ' Copiez le Bitmap dans le presse-papier
        Clipboard.SetImage(Bmp)
        'Remplace la sélection active de la zone de texte par le contenu du Presse-papiers.
        Frm_SDK.Journal.SelectionStart = Frm_SDK.Journal.Text.Length
        Frm_SDK.Journal.ReadOnly = False
        Frm_SDK.Journal.Paste()
        Frm_SDK.Journal.ReadOnly = True
      Catch ex As Exception
        Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      End Try
      Jrn_Add("SDK_Space")
    Catch ex As System.Runtime.InteropServices.ExternalException
      Dim MsgTit As String = Procédure_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      Nsd_i = MsgBox("Presse-Papier inaccessible. Please try again.",, MsgTit)
    End Try
  End Sub


  Public Function ClipBoard_Copier_New(Cel_Type As String) As String
    '-----------------------------------------------------------------------
    ' OK    Coller dans Angus Johnson 
    '              AJ accepte Espace ou point pour une valeur absente
    ' OK    Coller dans Coles à partir de Game/acTestGrid
    '-----------------------------------------------------------------------
    Dim Value As String = "#"
    Dim CC As String = ""
    'Une valeur absente est remplacée par un espace
    '                                 par un point  pour HODOKU
    Select Case Cel_Type
      Case "1"  ' Valeurs initiales
        For i As Integer = 0 To 80
          Select Case U(i, 1)
            Case " " : CC &= "."
            Case Else : CC &= U(i, 1)
          End Select
        Next i
        Jrn_Add("SDK_00000", {"VI        :" & CC})

      Case "2" ' Toutes les valeurs
        For i As Integer = 0 To 80
          Select Case U(i, 2)
            Case " " : CC &= "."
            Case Else : CC &= U(i, 2)
          End Select
        Next i
        Jrn_Add("SDK_00000", {"Valeurs   :" & CC})

      Case "3" ' Les candidats
        For i As Integer = 0 To 80
          CC &= U(i, 3).Replace(" ", "") & ";"
        Next i
        Jrn_Add("SDK_00000", {"Candidats :" & CC})
    End Select
    Clipboard.Clear()
    Clipboard.SetData(DataFormats.Text, CType(CC, Object))
    Value = Clipboard.GetText(TextDataFormat.Text)
    Return Value
  End Function

End Module