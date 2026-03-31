Option Strict On
Option Explicit On
Imports System.IO

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

    ' Le nom fait 11 caractères
    File_Save_Name &= "SDK_" & Code_FMDE & "_" & CStr(Prd_Number).PadLeft(5, CChar("0")) & ".txt"
    ' 31 Création du Fichier
    Try
      Using file As New IO.StreamWriter(File_Save_Name, False, Text.Encoding.UTF8)
        Dim Grille_Ini As String = ""
        Dim Grille_Val As String = ""
        Dim Grille_Sol As String = ""
        ' 31 Enregistrement des lignes du Fichier
        file.WriteLine(Msg_Read("PRD_20030", {File_Save_Name}))
        Dim s As Integer = InStrRev(File_Save_Name, "\")
        Dim Name As String = Mid$(File_Save_Name, s + 1, File_Save_Name.Length - s - 4)
        file.WriteLine(Msg_Read("PRD_20010", {Name}))
        For i As Integer = 0 To 80
          Grille_Ini &= Prd.Prd_Ini(i)
          Grille_Val &= Prd.Prd_Ini(i)
          Grille_Sol &= Prd.Prd_Val(i)
        Next i
        file.WriteLine(Msg_Read("PRD_20020", {Grille_Ini.Replace(" ", ".")}))
        file.WriteLine(Msg_Read("PRD_20021", {Grille_Val.Replace(" ", ".")}))
        file.WriteLine(Msg_Read("PRD_20023", {Grille_Sol.Replace(" ", ".")}))

        ' Affichage de DLCode et DLSolution 
        file.WriteLine(Msg_Read("PRD_20025", {Prd.Prd_DlCode}))
        file.WriteLine(Msg_Read("PRD_20026", {Prd.Prd_DlSolution}))
        file.WriteLine(Msg_Read("PRD_20030", {"Date_Heure de création       : " & Prd_DateTime}))
        file.WriteLine(Msg_Read("PRD_20030", {"Version                      : " & SDK_Version}))

        ' 32 Enregistrement des Informations de Prd
        file.WriteLine(Msg_Read("PRD_20030", {"Plcy_Strg_Profondeur         : " & Prd.Prd_Plcy_Strg_Profondeur}))
        file.WriteLine(Msg_Read("PRD_20030", {"Contrainte                   : " & Prd.Prd_Cnt_Type & CStr(Prd.Prd_Cnt_Valeur)}))
        file.WriteLine(Msg_Read("PRD_20030", {"Nb_Cellules_Demandées        : " & CStr(Prd.Prd_Create_Nb_Cel_Demandées)}))

        Dim S1 As String = "Str: "
        Dim S2 As String = "Crt: "
        Dim S3 As String = "Slv: "
        For i As Integer = 0 To 10
          S1 &= CStr(i).PadLeft(1) & "-" & Stg_List_Code(i) & " "
          S2 &= CStr(Prd.Crt_Strg_Nb(i)).PadLeft(5) & " "
          S3 &= CStr(Prd.Slv_Strg_Nb(i)).PadLeft(5) & " "
        Next i
        file.WriteLine(Msg_Read("PRD_20030", {S1}))
        file.WriteLine(Msg_Read("PRD_20030", {S2}))
        file.WriteLine(Msg_Read("PRD_20030", {S3}))

        If Prd.Prd_Ext_Triplet <> "#" Then
          file.WriteLine(Msg_Read("PRD_20030", {"Triplet                      : " & Prd.Prd_Ext_Triplet & " " & U_cr(Prd.Prd_Ext_Triplet_Cellule)}))
        End If
        If Prd.Prd_Ext_XWing <> "#" Then
          file.WriteLine(Msg_Read("PRD_20030", {"Xwing                        : " & Prd.Prd_Ext_XWing & " " & U_cr(Prd.Prd_Ext_XWing_Cellule)}))
        End If

        file.WriteLine(Msg_Read("PRD_20030", {"..."}))
      End Using
    Catch ex As Exception
      ' Une erreur se produit lors de la création du fichier
      '28/05/2024 le message permet de comprendre l'arrêt anormal du traitement
      Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      MsgBox(ex.ToString(),, MsgTit)
    End Try
    Return File_Save_Name
  End Function

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

    Jrn_Add(, {Proc_Name_Get() & " Modfification de PR_Stg"})
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
          Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
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
    Game_New_Game(Gnrl:="Nrm", "   ", Nom:=Nom, Prb:=Ini, Jeu:=Val, Sol:=Sol, Cdd729:=Cdd729, Frc:="5", Proc_Name_Get())
  End Sub

  Sub Pzzl_Load_Partie_Test()
    Jrn_Add(, {Proc_Name_Get()})

    Dim lastFile As String = My.Settings.SDK_Partie_Test
    Dim lastDir As String = Path.GetDirectoryName(lastFile)
    Dim lastName As String = Path.GetFileName(lastFile)

    Dim OFD As New OpenFileDialog With {
        .Title = "Choisir une Partie Test",
        .InitialDirectory = lastDir,
        .Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*",
        .FileName = lastName
    }

    If OFD.ShowDialog() <> DialogResult.OK Then Exit Sub

    Dim fileName As String = OFD.FileName
    If String.IsNullOrWhiteSpace(fileName) Then Exit Sub

    My.Settings.SDK_Partie_Test = fileName
    Jrn_Add("SDK_00130", {fileName})

    Pzzl_Open(fileName)
  End Sub

  Sub Pzzl_Write_Partie_Test()
    Jrn_Add(, {Proc_Name_Get()})

    Dim lastFile As String = My.Settings.SDK_Partie_Test
    Dim lastDir As String = Path.GetDirectoryName(lastFile)
    Dim lastName As String = Path.GetFileName(lastFile)

    Dim SFD As New SaveFileDialog With {
        .Title = "Enregistrer une Partie Test",
        .InitialDirectory = lastDir,
        .Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*",
        .FileName = lastName
    }

    If SFD.ShowDialog() <> DialogResult.OK Then Exit Sub

    Dim fileName As String = SFD.FileName
    If String.IsNullOrWhiteSpace(fileName) Then Exit Sub

    My.Settings.SDK_Partie_Test = fileName

    Dim dt As String = Now.ToString("yyyy_MM_dd_HH_mm_ss_fff")

    Using fs As New IO.StreamWriter(fileName, False, Text.Encoding.UTF8)

      Dim name As String = Path.GetFileNameWithoutExtension(fileName)
      fs.WriteLine(Msg_Read("PRD_20030", {fileName}))
      fs.WriteLine(Msg_Read("PRD_20010", {name}))
      Dim ini As New Text.StringBuilder()
      Dim val As New Text.StringBuilder()
      Dim sol As New Text.StringBuilder()
      Dim cdd As New Text.StringBuilder()

      For i As Integer = 0 To 80
        ini.Append(U(i, 1))
        val.Append(If(U(i, 2) <> " ", U(i, 2), U(i, 1)))
        sol.Append(U_Sol(i))
        cdd.Append(U(i, 3).Replace(" ", "")).Append(";"c)
      Next
      fs.WriteLine(Msg_Read("PRD_20020", {ini.ToString().Replace(" ", ".")}))
      fs.WriteLine(Msg_Read("PRD_20021", {val.ToString().Replace(" ", ".")}))
      fs.WriteLine(Msg_Read("PRD_20023", {sol.ToString().Replace(" ", ".")}))
      fs.WriteLine(Msg_Read("PRD_20027", {cdd.ToString()}))
      fs.WriteLine(Msg_Read("PRD_20024", {Stg_Profondeur}))
      fs.WriteLine(Msg_Read("PRD_20030", {"Date_Heure d'enregistrement : " & dt}))
      fs.WriteLine(Msg_Read("PRD_20030", {"..."}))

    End Using

    Jrn_Add("SDK_00122")
    Jrn_Add("SDK_00100", {fileName})
  End Sub

End Module