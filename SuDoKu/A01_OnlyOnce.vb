'Date de création: Dimanche 28/07/2024
'Ce Module regroupe l'ensemble des traitements qui ne sont faits qu'une seule fois

Friend Module A01_OnlyOnce
  Public Sub OO_000_SDK_Load()
    OO_100_Paths()
    OO_110_Variables_My_Settings()
    OO_120_Variables_Autres()
    Stg_List_Init()
    OC_Plcy_Stg_UOBTXYSJZKQ()
    OO_200_Msg_Load()
    OO_300_U_Init()
    OO_310_Hw_Init()
    OO_320_Hw_20Cell_Coll_Init()
    OO_330_Cell_FY_List_Build()
  End Sub

  Public Sub OO_999_SDK_Load_End()
    'Procédures avec affichage de messages
    OO_400_Pzzl_Load_Txt()
    OO_410_Pzzl_Hodoku_Load_Txt()
    OO_460_File_Delete()
  End Sub

  Public Sub OO_100_Paths()
    'Il est possible d'installer le répertoire SuDoKu_2026 dans n'importe quel autre répertoire
    Dim Path1 As String = Application.ExecutablePath
    Dim i As Integer = InStr(1, Path1, Base_Folder, vbTextCompare)
    Dim Path2 As String = Path1.Substring(0, i - 1)
    Path_SDK = Path2 & Base_Folder & "\"

    'Les Paths
    File_SDK = Path_SDK & "S50_SDK\"
    File_Correction = Path_SDK & "S90_Corrections\M95_Correction.txt"
    File_SDKMsg = Path_SDK & "S20_Initial\Msg.ini"
    File_Puzzle = Path_SDK & "S20_Initial\GSP_Puzzle.txt"
    File_Hodoku = Path_SDK & "S20_Initial\GSP_PuzzleHodoku.txt"
    File_SDKEtd = Path_SDK & "S50_SDK\Etude\"
    File_SDKJrn = Path_SDK & "S50_SDK\Jrn_"

    Path_Batch = Path_SDK & "S50_SDK\Batch\"
    Path_Batch_Poubelle = Path_SDK & "S50_SDK\Batch_Poubelle\"
    Path_Save = Path_SDK & "S50_SDK\"
    Path_SDK_Autres_Jeux = Path2 & Base_Folder & "_S95\"
    File_SDKDoc = Path_SDK & "S01_Documentation\Rapport.docx"
    File_ValUsi = Path_SDK & "S20_Initial\Valeurs_Initiales.ini"

  End Sub

  Public Sub OO_110_Variables_My_Settings()
    'Initialisation des variables à/p de My.Settings
    With My.Settings
      WH = .Prf_01C_Taille_Cellule
      Obj_Symbol = .Obj_Symbol
      Obj_Forme = .Obj_Forme

      If WH = 0 Then WH = 80
      Plcy_Dancing_Link = .Prf_05D_Plcy_Dancing_Link
      Plcy_Open_Display = .Prf_05D_Plcy_Open_Display
      Plcy_Fond_Grille = .Prf_01C_Fond_Grille
      Plcy_MouseClick_Middle = .Prf_01C_MouseClick_Middle
      Plcy_Generate_Batch = .Prf_02C_Plcy_Batch_Generate
      Create_Contrainte_Originale = .Prf_02C_Create_Contrainte
      Create_Chat = .Prf_02C_Create_Chat
      Create_Grille_CdU = .Prf_02C_Create_Grille_CdU
      Create_Nb_Cel_Demandées = CStr(.Prf_02C_Create_Nb_Cel_Demandées)
      Create_Nb_Tentatives = CStr(.Prf_02C_Create_Nb_Tentatives)
      Lettre_Flèche = .Prf_08H_Flèche
      Select Case Lettre_Flèche
        Case 0 : Lettre_Flèche_ChrW = 64
        Case 1 : Lettre_Flèche_ChrW = 96
        Case 2 : Lettre_Flèche_ChrW = 944
        Case 3 : Lettre_Flèche_ChrW = 1039
        Case Else : Lettre_Flèche_ChrW = 64
      End Select
      Plcy_Stg_Clb = .Prf_03R_Plcy_Stg_Clb
      'Placé à cet endroit pour conserver les couleurs des Valeurs et des Candidats
      'OC_Thèmes_Couleurs(.Thème_Clr)
      Select Case .Format_DAB
        Case 2 : OC_Thèmes_Couleurs(Rdc.Next(0, 3))
        Case Else : OC_Thèmes_Couleurs(.Thème_Clr)
      End Select
      Color_VI = Color.FromName(.Prf_01C_Clr_ComboBoxVI)
      Color_VCdd = Color.FromName(.Prf_01C_Clr_ComboBoxVCdd)
      Select Case .Format_DAB
        Case 0 To 1 : Plcy_Format_DAB = .Format_DAB
        Case 2 : Plcy_Format_DAB = Rdc.Next(0, 2)
      End Select
    End With
  End Sub
  Public Sub OO_120_Variables_Autres()
    WHhalf = (WH \ 2)
    WHthird = (WH \ 3)
    WHquart = (WH \ 4)
    Bld_WH_Grid = (WH * 9) + 3 + 1 + 1 + 3 + 1 + 1 + 3 + 1 + 1 + 3
  End Sub
  ''' <summary>Charge le fichier des messages dans un dictionnaire.</summary>
  Public Sub OO_200_Msg_Load()
    'Cette procédure améliore la vitesse d'affichage d'un message qui sont chargés dans un dictionnaire
    Dim Msg As String
    Dim MsgID As String
    Dim Rcd As New IO.StreamReader(File_SDKMsg, Text.Encoding.Default)
    Dim Ligne As String
    Dim l As Integer
    Ligne = Rcd.ReadLine
    Do Until Ligne Is Nothing
      Try 'Le premier enregistrement comporte la balise [MSG]
        l = Ligne.Length
        If Ligne.Substring(0, 1) <> ";" Then
          MsgID = Ligne.Substring(0, 9)
          Msg = Ligne.Substring(10, l - 10)
          Msg_Dcty.Add(MsgID, Msg)
        End If
      Catch ex As Exception
        'la procédure Jrn_Add n'est pas encore opérationnelle
      End Try
      Ligne = Rcd.ReadLine()
    Loop
    Rcd.Close()
  End Sub
  Sub OO_300_U_Init()
    For i As Integer = 0 To 80
      U(i, 0) = CStr(i)                                                     ' Numéro de la Cellule
      U(i, 1) = " "                                                         ' Valeur Initiale                
      U(i, 2) = " "                                                         ' Valeur                         
      U(i, 3) = Cnddts                                                      ' Les candidats 123456789        
      U_CddExc(i) = Cnddts_Blancs                                           ' Les candidats exclus           
      U_Sol(i) = " "                                                        ' Solution                      
      U_Col(i) = i Mod 9
      U_Row(i) = i \ 9
      U_Reg(i) = Wh_Région_Cel(i)
      'U_Suggest(i) = "0"
    Next i
    For i As Integer = 0 To 80
      U_cr(i) = "L" & U_Row(i) + 1 & "_" & "C" & U_Col(i) + 1               ' Format Lx_Cy
    Next i
    'Une fois U_Reg construit
    For i As Integer = 0 To 80
      Select Case U_Reg(i)
        Case 0 : U_Bh(i) = 0 : U_Bv(i) = 0
        Case 1 : U_Bh(i) = 0 : U_Bv(i) = 1
        Case 2 : U_Bh(i) = 0 : U_Bv(i) = 2
        Case 3 : U_Bh(i) = 1 : U_Bv(i) = 0
        Case 4 : U_Bh(i) = 1 : U_Bv(i) = 1
        Case 5 : U_Bh(i) = 1 : U_Bv(i) = 2
        Case 6 : U_Bh(i) = 2 : U_Bv(i) = 0
        Case 7 : U_Bh(i) = 2 : U_Bv(i) = 1
        Case 8 : U_Bh(i) = 2 : U_Bv(i) = 2
      End Select
    Next i
  End Sub
  Sub OO_Wh_27CellulesBande()
    ' La fonction calcule les 27 cellules de chaque bande {"Bh1", "Bh2", "Bh3", "Bv1", "Bv2", "Bv3"}
    Dim Index As Integer = 0
    For Each kvp As KeyValuePair(Of String, Integer) In Bande_Dcty
      Dim n As Integer = 0
      For cellule As Integer = 0 To 80
        If (kvp.Key.StartsWith("Bh") AndAlso U_Bh(cellule) = kvp.Value) _
    OrElse (kvp.Key.StartsWith("Bv") AndAlso U_Bv(cellule) = kvp.Value) Then
          U_Bandes(Index, n) = cellule : n += 1
        End If
      Next cellule
      Index += 1
    Next kvp
  End Sub
  Sub OO_310_Hw_Init()
    'Ce traitement initialise une fois pour toutes les 9 cellules composant une Ligne, une Colonne et une Région
    For i As Integer = 0 To 8
      U_9CelRow(i) = Wh_9CellulesRow_Row(i)
      U_9CelCol(i) = Wh_9CellulesColumn_Col(i)
      U_9CelReg(i) = Wh_9CellulesRégion_Rég(i)
    Next i
    'Composition des 6 bandes (3 horizontales et 3 verticales)
    OO_Wh_27CellulesBande()
  End Sub
  Sub OO_320_Hw_20Cell_Coll_Init()
    'Initialisation des 20 Cellules Collatérales des 81 Cellules de la grille
    'L'utilisation des cellules collatérales ne prend pas en compte la cellule d'origine
    'On remplit 8+8+8=24 cellules, on omet les cellules en double de la région
    '                              on trie
    '                              on documente le tableau U_20Cell_Coll
    'Traitement des 81 Cellules
    For Cellule As Integer = 0 To 80
      Dim Grp20(19) As Integer
      Dim g As Integer = 0
      'Chargement des 8 cellules de la ligne
      Dim GrpRow As Integer() = U_9CelRow(U_Row(Cellule))
      For rw As Integer = 0 To 8
        If GrpRow(rw) = Cellule Then Continue For 'La cellule origine n'est pas ajoutée
        Grp20(g) = GrpRow(rw)
        g += 1
      Next rw
      'Chargement des 8 cellules de la colonne
      Dim GrpCol As Integer() = U_9CelCol(U_Col(Cellule))
      For cl As Integer = 0 To 8
        If GrpCol(cl) = Cellule Then Continue For 'La cellule origine n'est pas ajoutée
        Grp20(g) = GrpCol(cl)
        g += 1
      Next cl
      'Chargement des 8 cellules de la région
      Dim GrpReg As Integer() = U_9CelReg(U_Reg(Cellule))
      For rg As Integer = 0 To 8
        If GrpReg(rg) = Cellule Then Continue For 'La cellule origine n'est pas ajoutée
        Dim CelExistante As Boolean = False
        For k As Integer = 0 To 16
          If GrpReg(rg) = Grp20(k) Then CelExistante = True
        Next k
        If Not CelExistante Then           'Une cellule déjà existante n'est pas ajoutée
          Grp20(g) = GrpReg(rg)
          g += 1
        End If
      Next rg
      'Tri et chargement
      Array.Sort(Grp20)
      U_20Cell_Coll(Cellule) = Grp20
    Next Cellule
  End Sub
  Sub OO_330_Cell_FY_List_Build()
    ' Liste aléatoire des cellules suivant algorithme de Fisher-Yates, Knuth
    ' Étape 1 : Créer la liste Cell_FY_Lis aléatoire des entiers de 0 à 80
    For i As Integer = 0 To 80
      Cell_FY_List.Add(i)
    Next
    ' Étape 2 : Mélanger la liste
    ' - Le mélange est fait avec l’algorithme de Fisher-Yates, qui est efficace et équitable.
    For i As Integer = Cell_FY_List.Count - 1 To 1 Step -1
      Dim j As Integer = Rdf.Next(i + 1)
      ' Échanger les éléments i et j
      Dim temp As Integer = Cell_FY_List(i)
      Cell_FY_List(i) = Cell_FY_List(j)
      Cell_FY_List(j) = temp
    Next
  End Sub
  Sub OO_400_Pzzl_Load_Txt()
    ' Chargement des parties du fichier GSP_Puzzle.txt dans Pzzl_List
    Dim Ligne As String()
    Dim Nom, Force, Problème, Solution As String
    Dim Nb_CI As Integer
    Using FileTxt As New Microsoft.VisualBasic.FileIO.TextFieldParser(File_Puzzle)
      FileTxt.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.FixedWidth
      FileTxt.SetFieldWidths(15, 1, 1, 1, 81, 1, 81, 1, 6, 1, 6, 1, 1, 1)
      '15|1|81|81|6|6|X|
      While Not FileTxt.EndOfData
        Try
          Ligne = FileTxt.ReadFields()
          Nom = Ligne(0)                                 'Le Nom
          Force = Ligne(2).Substring(0, 1)               'La Force
          Problème = Ligne(4).Substring(0, 81)           'Le Problème 
          Solution = Ligne(6).Substring(0, 81)           'La Solution 
          Nb_CI = 81                                     'Nombre de Cellules Initiales
          For i As Integer = 0 To 80
            If Problème.Substring(i, 1) = "0" Then Nb_CI += -1
          Next i
          Pzzl_List.Add(New Pzzl_Cls(Nom, Force, Problème, Solution, Nb_CI))
        Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
          Jrn_Add("SDK_00051", {Proc_Name_Get(), ex.Message})
        End Try
      End While
      'Jrn_Add("SDK_00050", {File_Puzzle, CStr(Pzzl_List.Count + 1)})
    End Using
  End Sub
  Sub OO_410_Pzzl_Hodoku_Load_Txt()
    Dim Rcd As New System.IO.StreamReader(File_Hodoku, System.Text.Encoding.Default)
    Dim Ligne, Nom, Jeu, Prb, Point81 As String
    Dim l As Integer
    Ligne = Rcd.ReadLine
    Do Until Ligne Is Nothing
      Try
        Nom = ""
        Jeu = "_"
        Prb = ""
        Point81 = StrDup(81, ".")
        l = Ligne.Length
        If l > 1 Then
          Nom = Ligne.Substring(83, l - 83)
          Prb = Ligne.Substring(0, 81)
          If Prb <> Point81 Then Jeu = "Oui"
          Pzzl_Hodoku_List.Add(New Pzzl_Hodoku_Cls(Nom, Jeu, Prb))
        End If
      Catch ex As Exception
        Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
        Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
      End Try
      Ligne = Rcd.ReadLine()
    Loop
    Rcd.Close()
    'Jrn_Add("SDK_00052", {File_Hodoku, CStr(Pzzl_Hodoku_List.Count + 1)})
  End Sub

  Sub OO_460_File_Delete()
    Jrn_Add("SDK_00060")
    Dim Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(File_SDK)
    OO_461_File_Delete(File_SDK, "txt")
    OO_461_File_Delete(File_SDK, "log")
    OO_461_File_Delete(File_SDK, "docx")
    Files = From File In IO.Directory.GetFiles(Path_Batch)
    'SDK_00061 = le répertoire des parties "en stock" %0 contient %1 fichier(s). 
    Jrn_Add("SDK_00061", {Path_Batch, CStr(Files.Count)})

    OO_461_File_Delete(Path_Batch_Poubelle, "txt")
    'Files = From File In IO.Directory.GetFiles(Path_Batch_Poubelle)
    'Jrn_Add("SDK_00061", {Path_Batch_Poubelle, CStr(Files.Count)})
  End Sub

  Sub OO_461_File_Delete(Répertoire As String, Extension As String)
    Dim Limite As Integer = 5000
    Dim Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(Répertoire)
                                          Where File.Contains(Extension)
                                          Order By File Descending
    Dim Nb_Files As Integer = Files.Count
    If Nb_Files > Limite Then
      Jrn_Add("SDK_00062", {Répertoire, CStr(Files.Count).PadLeft(6), Extension})
      Jrn_Add("SDK_00063", {CStr(Limite)})
      Dim Nb As Integer = 0
      For Each File As String In Files
        Nb += 1
        If Nb > Limite Then
          Try
            Jrn_Add("SDK_00064", {CStr(Nb).PadLeft(3), File})
            My.Computer.FileSystem.DeleteFile(File,
               Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
               Microsoft.VisualBasic.FileIO.RecycleOption.DeletePermanently)
          Catch ex As Exception
            Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
            Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
          End Try
        End If
      Next file
      Select Case Nb - Limite
        Case 0 : Jrn_Add("SDK_00065", {File_SDK})
        Case 1 : Jrn_Add("SDK_00066")
        Case Else : Jrn_Add("SDK_00067", {CStr(Nb - Limite)})
      End Select
    End If
  End Sub
End Module