Option Strict On
Option Explicit On


Module M01_Enregistrement_Partie


  Public Function Pzzl_Save(Prd As Prd_Struct) As String
    ' Enregistrement d'un Puzzle
    ' 17/10/2025 Modification du nom du fichier FMDE
    Dim Code_FMDE As String = "#"
    Dim File_Save_Name As String = ""

    ' 1  La grille doit être DLu
    If Prd.Prd_DlCode <> "Dlu" Then
      Return "#_" & Prd.Prd_DlCode
    End If

    '    La grille est DLu
    ' 2  SDK a t-il résolu la grille ? SDK sait créer des grilles qu'il ne sait pas résoudre
    Dim Grille_Résolue As Boolean = True
    For i As Integer = 0 To 80
      If Prd.Prd_Val(i) = " " Then
        Grille_Résolue = False
        Exit For
      End If
    Next i
    ' La grille n'a pas été résolue pas SDK 
    If Not Grille_Résolue Then Code_FMDE = "E"

    If Code_FMDE = "#" Then
      If Prd.Slv_Strg_Nb(10) > 0 _
      Or Prd.Slv_Strg_Nb(9) > 0 _
      Or Prd.Slv_Strg_Nb(8) > 0 _
      Or Prd.Slv_Strg_Nb(7) > 0 _
      Or Prd.Slv_Strg_Nb(6) > 0 _
      Or Prd.Slv_Strg_Nb(5) > 0 _
      Or Prd.Slv_Strg_Nb(4) > 0 Then
        Code_FMDE = "D"
      End If
    End If

    If Code_FMDE = "#" Then
      If Prd.Slv_Strg_Nb(3) > 0 _
      Or Prd.Slv_Strg_Nb(2) > 0 Then
        Code_FMDE = "M"
      End If
    End If

    If Code_FMDE = "#" Then
      If Prd.Slv_Strg_Nb(1) > 0 _
      Or Prd.Slv_Strg_Nb(0) > 0 Then
        Code_FMDE = "F"
      End If
    End If

    ' 3  Enregistrement
    If Prd.Prd_BI = "I" Then File_Save_Name = Path_Save & "Interactif_Poubelle\"
    If Prd.Prd_BI = "B" Then File_Save_Name = Path_Save & "Batch\"

    Dim Prd_Number As Integer = My.Settings.Prd_Number
    Prd_Number += 1
    My.Settings.Prd_Number = Prd_Number
    Dim Prd_DateTime As String = Format(Now, "yyyy_MM_dd") & "_" & Format(Now, "HH_mm_ss_fff")

    File_Save_Name &= "SDK_" & Code_FMDE & "_" & CStr(Prd_Number).PadLeft(5, CChar("0")) & ".txt"
    ' 31 Création du Fichier
    Try
      Dim File As System.IO.FileStream = System.IO.File.Create(File_Save_Name)
      Dim Ligne As String
      Dim Grille_Ini As String = ""
      Dim Grille_Val As String = ""
      Dim Grille_Sol As String = ""
      ' 31 Enregistrement des lignes du Fichier
      Ligne = Msg_Read_IA("PRD_20030", {File_Save_Name})
      Pzzl_Record_Ligne(File, Ligne)
      Dim s As Integer = InStrRev(File_Save_Name, "\")
      Dim Name As String = Mid$(File_Save_Name, s + 1, File_Save_Name.Length - s - 4)
      Ligne = Msg_Read_IA("PRD_20010", {Name})
      Pzzl_Record_Ligne(File, Ligne)
      For i As Integer = 0 To 80
        Grille_Ini &= Prd.Prd_Ini(i)
        Grille_Val &= Prd.Prd_Ini(i)
        Grille_Sol &= Prd.Prd_Val(i)
      Next i
      Ligne = Msg_Read_IA("PRD_20020", {Grille_Ini.Replace(" ", ".")})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20021", {Grille_Val.Replace(" ", ".")})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20023", {Grille_Sol.Replace(" ", ".")})
      Pzzl_Record_Ligne(File, Ligne)

      ' Affichage de DLCode et DLSolution 
      Ligne = Msg_Read_IA("PRD_20025", {Prd.Prd_DlCode})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20026", {Prd.Prd_DlSolution})
      Pzzl_Record_Ligne(File, Ligne)

      Ligne = Msg_Read_IA("PRD_20030", {"Date_Heure de création       : " & Prd_DateTime})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20030", {"Version                      : " & SDK_Version})
      Pzzl_Record_Ligne(File, Ligne)

      ' 32 Enregistrement des Informations de Prd
      Ligne = Msg_Read_IA("PRD_20030", {"Plcy_Strg_Profondeur         : " & Prd.Prd_Plcy_Strg_Profondeur})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20030", {"Contrainte                   : " & Prd.Prd_Cnt_Type & CStr(Prd.Prd_Cnt_Valeur)})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20030", {"Nb_Cellules_Demandées        : " & CStr(Prd.Prd_Create_Nb_Cel_Demandées)})
      Pzzl_Record_Ligne(File, Ligne)

      Dim S1 As String = "Str: "
      Dim S2 As String = "Crt: "
      Dim S3 As String = "Slv: "
      For i As Integer = 0 To 10
        S1 &= CStr(i).PadLeft(1) & "-" & Stg_List_Code(i) & " "
        S2 &= CStr(Prd.Crt_Strg_Nb(i)).PadLeft(5) & " "
        S3 &= CStr(Prd.Slv_Strg_Nb(i)).PadLeft(5) & " "
      Next i
      Ligne = Msg_Read_IA("PRD_20030", {S1})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20030", {S2})
      Pzzl_Record_Ligne(File, Ligne)
      Ligne = Msg_Read_IA("PRD_20030", {S3})
      Pzzl_Record_Ligne(File, Ligne)

      If Prd.Prd_Ext_Triplet <> "#" Then
        Ligne = Msg_Read_IA("PRD_20030", {"Triplet                      : " & Prd.Prd_Ext_Triplet & " " & U_cr(Prd.Prd_Ext_Triplet_Cellule)})
        Pzzl_Record_Ligne(File, Ligne)
      End If
      If Prd.Prd_Ext_XWing <> "#" Then
        Ligne = Msg_Read_IA("PRD_20030", {"Xwing                        : " & Prd.Prd_Ext_XWing & " " & U_cr(Prd.Prd_Ext_XWing_Cellule)})
        Pzzl_Record_Ligne(File, Ligne)
      End If

      Ligne = Msg_Read_IA("PRD_20030", {"..."})
      Pzzl_Record_Ligne(File, Ligne)
      ' 33 Fermeture du Fichier
      File.Close()

    Catch ex As Exception
      ' Une erreur se produit lors de la création du fichier
      '28/05/2024 le message permet de comprendre l'arrêt anormal du traitement
      Dim MsgTit As String = Procédure_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      MsgBox(ex.ToString(),, MsgTit)
    End Try
    Return File_Save_Name
  End Function

  Public Sub Pzzl_Record_Ligne(File As System.IO.FileStream, Ligne As String)
    ' Enregistre une ligne
    Dim WLigne As Byte() = New System.Text.UTF8Encoding(True).GetBytes(Ligne & vbCrLf)
    File.Write(WLigne, 0, WLigne.Length)
  End Sub

  Sub Pzzl_Open(File_Name As String)
    ' Chargement_Ouverture d'une partie
    ' 14/11/2024 NotePad est ouvert si l'option est True et si le journal est affiché
    ' Ouverture de Notepad
    ' Liste dans le journal
    ' J 24/07/2025 Modification de la Profondeur de Stratégie
    Dim LL() As String
    Dim Nom As String = ""
    Dim Ini As String = ""
    Dim Val As String = ""
    Dim Sol As String = ""
    Dim Stg As String = ""
    Dim Cdd As String = ""

    Jrn_Add(, {Procédure_Name_Get() & " Modfification de PR_Stg"})
    Jrn_Add("SDK_00130", {File_Name})

    'Chargement de la partie dans U(i,123)
    Using File As New FileIO.TextFieldParser(File_Name)
      'LL comporte x postes séparés par le séparateur =
      'LL(0) est la première partie jusqu'au delimiter =
      'LL(1) est la seconde partie
      'Soit PR_xyz=12345...8945 (Prb/Jeu/Sol)
      'Soit des commentaires commençant par ' ou par n'importe quoi
      File.TextFieldType = FileIO.FieldType.Delimited
      File.SetDelimiters("=")
      While Not File.EndOfData
        Try
          'Autant de fois que de lignes dans le fichier
          LL = File.ReadFields()
          Select Case LL(0)
            Case "PR_Nom" : Nom = LL(1)
            Case "PR_Ini" : Ini = LL(1)
            Case "PR_Val" : Val = LL(1)
            Case "PR_Sol" : Sol = LL(1)
            Case "PR_Stg" : Stg = LL(1)
            Case "PR_Cdd" : Cdd = LL(1)
          End Select
          'Nom, Ini, Val et Sol sont affichés ET les commentaires
          'Le tableau LL est affiché, chaque partie réunie par le delimiter =
          Jrn_Add(, {Join(LL, "=")})
        Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
          Jrn_Add("ERR_00000", {Procédure_Name_Get()}, "Erreur")
          Jrn_Add("ERR_00000", {ex.Message})
          Jrn_Add("ERR_00000", {ex.ToString()})
        End Try
      End While
    End Using
    If Stg <> "" Then
      Jrn_Add("Prl_00000", {"Profondeur         : "})
      Jrn_Add("Prl_00000", {"Avant              : " & Stg_Profondeur})
      Jrn_Add("Prl_00000", {"... partie test    : " & Stg})
      Stg_Profondeur = Stg
      For i As Integer = 0 To Plcy_Stg_Clb.Count - 1
        Dim c As Char = CChar(Stg(i).ToString())
        If Char.IsUpper(c) Then
          Mid$(Plcy_Stg_Clb, i + 1, 1) = "O"
          Stg_Bll(i) = True
        ElseIf Char.IsLower(c) Then
          Mid$(Plcy_Stg_Clb, i + 1, 1) = "N"
          Stg_Bll(i) = False
        Else
          Mid$(Plcy_Stg_Clb, i + 1, 1) = "#"
          Stg_Bll(i) = False
        End If
      Next i
      My.Settings.Prf_03R_Plcy_Stg_Clb = Plcy_Stg_Clb
      ' Affichage de contrôle
      Jrn_Add("Prl_00000", {"Profondeur         : "})
      Jrn_Add("Prl_00000", {"Avant              : " & Stg_Profondeur})
      Jrn_Add("Prl_00000", {"... partie test    : " & Stg})
      Jrn_Add("Prl_00000", {"Stratégies         : "})
      Jrn_Add("Prl_00000", {"Plcy_Stg_Clb       : " & Plcy_Stg_Clb})
      Jrn_Add("Prl_00000", {"Stg_Profondeur     : " & Stg_Profondeur})
    End If

    '#440
    ' 3  Gestion des boutons des stratégies Enabled/Unelabled de la barre d'outils
    'Stg_List_Code: CdU , CdO, Cbl, Tpl, Xwg, XYw, Swf, Jly, XYZ, SKy, Unq
    For i As Integer = 0 To Stg_Bll.Count - 1
      Dim Btn_Name As String = "Btn_" & Stg_List_Code.Item(i)
      If Frm_SDK.BarreOutils.Items.ContainsKey(Btn_Name) Then
        If Stg_Bll(i) = True Then Frm_SDK.BarreOutils.Items(Btn_Name).Enabled = True
        If Stg_Bll(i) = False Then Frm_SDK.BarreOutils.Items(Btn_Name).Enabled = False
      End If
    Next

    '#443
    'Gestion des candidats
    Dim Cdd729 As String = ""
    If Cdd = "" Then
      Cdd729 = StrDup(729, " ")
    Else
      Dim Cdd81() As String = Cdd.Split(";"c)
      Dim Cdd81e(80) As String   ' 81 cases
      For i As Integer = 0 To 80
        Dim a As String = Cdd81(i)
        Dim b As String = StrDup(9, " ")  ' 9 blancs

        If a <> "" Then
          For Each ch As Char In a
            Dim d As Integer = CInt(ch.ToString())
            ' placer le chiffre à la bonne position (1→index0, 9→index8)
            Mid$(b, d, 1) = ch
          Next
        End If
        Cdd81e(i) = b
      Next
      Cdd729 = String.Join("", Cdd81e)
    End If
    Game_New_Game(Gnrl:="Nrm", Nom:=Nom, Prb:=Ini, Jeu:=Val, Sol:=Sol, Cdd729:=Cdd729, Frc:="5")
  End Sub

  Sub Pzzl_Load_Partie_Test()
    Dim s, l As Integer
    Dim File_Name As String = ""
    Jrn_Add(, {Procédure_Name_Get()})

    Dim Last_Paramètre As String = My.Settings.SDK_Partie_Test
    s = InStrRev(Last_Paramètre, "\")
    l = Last_Paramètre.Length
    Dim Last_Répertoire As String = Mid$(Last_Paramètre, 1, s)
    Dim Last_Partie As String = Mid$(Last_Paramètre, s + 1, l - s)

    '1 OpenFileDialog pour sélectionner le fichier
    Dim OFD As New OpenFileDialog With
        {
        .Title = "Choisir une Partie Test",
        .InitialDirectory = Last_Répertoire,
        .Filter = "(*.txt)|*.txt|All files (*.*)|*.*",
        .FileName = Last_Partie
        }
    Dim OFDStream As System.IO.Stream = Nothing
    If OFD.ShowDialog() = System.Windows.Forms.DialogResult.OK Then 'OK_La partie est choisie
      Try
        OFDStream = OFD.OpenFile()
        If (OFDStream IsNot Nothing) Then File_Name = OFD.FileName.ToString()
      Catch Ex As Exception
        Jrn_Add("ERR_00000", {Ex.Message}, "Erreur")
        Jrn_Add("ERR_00000", {Ex.ToString()}, "Erreur")
        Exit Sub
      Finally
        If (OFDStream IsNot Nothing) Then OFDStream.Close()
      End Try
    End If

    If File_Name = "" Then Exit Sub

    My.Settings.SDK_Partie_Test = File_Name
    Jrn_Add("SDK_00130", {File_Name})
    Pzzl_Open(File_Name)
  End Sub

  Sub Pzzl_Write_Partie_Test()
    'Enregistrement d'une Partie Test 
    'Structure conforme à l'enregistrement M50_Production.vb Pzzl_Save
    '20/11/2025 Enregistrement des candidats
    Dim Texte As String = "Nom de la Partie Test: "
    Dim Titre As String = "Enregistrement d'une Partie Test"
    Dim DftValue As String = My.Settings.SDK_Partie_Test
    Dim Réponse As String = InputBox(Texte, Titre, DftValue, Frm_SDK.Left, Frm_SDK.Top)
    If Réponse Is "" Then
      Jrn_Add("SDK_00121")
      Exit Sub 'Cancel enfoncé
    End If

    Dim File_Save_Name As String = Réponse.ToString()
    Dim Prd_DateTime As String = Format(Now, "yyyy_MM_dd") & "_" & Format(Now, "HH_mm_ss_fff")

    ' Delete the file if it exists. 
    If IO.File.Exists(File_Save_Name) Then
      Dim MsgTit As String = Procédure_Name_Get() & " " & Application.ProductName & " " & SDK_Version

      If MsgBox("Il existe déjà un fichier " & File_Save_Name & vbCrLf &
                "Voulez-vous le remplacer ? " & vbCrLf & "(Les anciens commentaires seront effacés.)",
                MsgBoxStyle.OkCancel, MsgTit) = vbCancel Then Exit Sub
      IO.File.Delete(File_Save_Name)
    End If

    Dim File As IO.FileStream = IO.File.Create(File_Save_Name)
    Dim Ligne As String
    Dim Grille_Ini As String = ""
    Dim Grille_Val As String = ""
    Dim Grille_Sol As String = ""
    Dim Grille_Cdd As String = ""

    Ligne = Msg_Read_IA("PRD_20030", {File_Save_Name})
    Pzzl_Record_Ligne(File, Ligne)
    Dim s As Integer = InStrRev(File_Save_Name, "\")
    Dim Name As String = Mid$(File_Save_Name, s + 1, File_Save_Name.Length - s - 4)
    Ligne = Msg_Read_IA("PRD_20010", {Name})
    Pzzl_Record_Ligne(File, Ligne)

    For i As Integer = 0 To 80
      Grille_Ini &= U(i, 1)                                                        'Problème
      If U(i, 2) <> " " Then Grille_Val &= U(i, 2) Else Grille_Val &= U(i, 1)      'Jeu
      Grille_Sol &= U_Sol(i)                                                       'Solution
      Grille_Cdd &= U(i, 3).Replace(" ", "") & ";"
    Next i

    Ligne = Msg_Read_IA("PRD_20020", {Grille_Ini.Replace(" ", ".")})
    Pzzl_Record_Ligne(File, Ligne)
    Ligne = Msg_Read_IA("PRD_20021", {Grille_Val.Replace(" ", ".")})
    Pzzl_Record_Ligne(File, Ligne)
    Ligne = Msg_Read_IA("PRD_20023", {Grille_Sol.Replace(" ", ".")})
    Pzzl_Record_Ligne(File, Ligne)
    Ligne = Msg_Read_IA("PRD_20027", {Grille_Cdd})
    Pzzl_Record_Ligne(File, Ligne)
    Ligne = Msg_Read_IA("PRD_20024", {Stg_Profondeur})
    Pzzl_Record_Ligne(File, Ligne)
    Ligne = Msg_Read_IA("PRD_20030", {"Date_Heure d'enregistrement  : " & Prd_DateTime})
    Pzzl_Record_Ligne(File, Ligne)

    Ligne = Msg_Read_IA("PRD_20030", {"..."})
    Pzzl_Record_Ligne(File, Ligne)

    File.Close()
    My.Settings.SDK_Partie_Test = Réponse.ToString()
    Jrn_Add("SDK_00122")
    Jrn_Add("SDK_00100", {File_Save_Name})
  End Sub
End Module