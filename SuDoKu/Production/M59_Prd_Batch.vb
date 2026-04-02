Imports System.Threading

'08/12/2023  
Friend Module M59_Prd_Batch
  Public Sub Batch_Initial()
    '///////////////////////////////////////////////////////////////////////////////////////
    '#616 Samedi 14/02/2026
    'Je bloque temporairement la génération de grilles en arrière-plan 
    ' Pour des problèmes de stabilisation de l'application
    ' je soupçonne ce traitement de grilles en arrière-plan d'être à l'origine de certains arrêts anormaux de l'application
    '///////////////////////////////////////////////////////////////////////////////////////
    Jrn_Add_Red(Proc_Name_Get() & " #616 La génération de grilles en arrière-plan est temporairement désactivée.")
    Exit Sub

    'Une fois l'application lancée, id dès que la grille est affichée,
    'la procédure est lancée If Plcy_Generate_Batch
    'Un thread en arrière-plan est automatiquement tué quand l'application se ferme.
    'Pour éviter cela, on utilise un thread en premier plan  

    'Nom: devenv
    'Programme: C : \Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\devenv.exe
    'Je lance SDK avec SuDoKu.exe         NON
    'Je lance SDK avec SuDoKu 16          OUI
    'Je lance SDK avec Visual Studio 2022 OUI

    If ProcessStarted("devenv") Then
      Event_OnPaint_MAP = Proc_Name_Get()
      Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      Dim reponse As MsgBoxResult
      reponse = MsgBox("Vous êtes en maintenance d'application " & vbCrLf &
                   "devenv est démarré " & vbCrLf &
                   "Souhaitez-vous lancer la génération de grilles en arrière-plan ?",
                   MsgBoxStyle.YesNoCancel, MsgTit)
      If reponse = vbNo Or reponse = vbCancel Then Exit Sub
    End If

    Batch_Thread = New Thread(AddressOf Batch_Sudoku) _
      With {.IsBackground = False,
            .Priority = ThreadPriority.Lowest,
            .Name = "Batch_Sudoku"}
    Batch_Thread.Start()
  End Sub
  Public Sub Batch_Sudoku()
    'Cette procédure s'exécute dans un Thread Batch_Thread
    'et à la fin le Thread est terminé
    'Ce thread dépend du thread principal, lorsque celui-ci est arrêté,
    '          alors Batch_Thread est également arrêté
    Dim Nb_F As Integer = File_Nb("SDK_F")
    Dim Nb_M As Integer = File_Nb("SDK_M")
    Dim Nb_D As Integer = File_Nb("SDK_D")
    Dim Nb_E As Integer = File_Nb("SDK_E")
    Try
      Batch_en_Cours = True
      'Stock de grilles
      Dim Nb_Max As Integer = My.Settings.Prf_02C_Nb_Batch_Generate
      'Les stocks de grilles sont complétés pour les grilles faciles, ensuite moyennes et enfin difficiles
      'Un nombre limité de création est défini 
      Dim Nb_Limite As Integer = 5
      Dim Nb As Integer = 0

      ' 01 Définition des paramètres de la structure Prd
      '   Prd est d'abord initialisée à partir de Préférences / Création
      Dim Prd As Prd_Struct = Nothing
      ' La contrainte est définie à ce niveau, ie les puzzles créés pendant un batch ont tous la même contrainte
      Dim U_temp(0 To 80, 0 To 3) As String
      Prd_Init(Prd, U_temp, "B")
      Nb = 0
      Jrn_Add_Red(Proc_Name_Get() & " Phase F.")
      For i As Integer = 1 To Nb_Max - Nb_F
        Nb += 1
        If Nb > Nb_Limite Then Exit For
        'Prd est ensuite aménagée des valeurs particulières de la difficulté demandée pour la grille
        Prd.Prd_Chat = False
        Prd.Prd_Create_Nb_Cel_Demandées = 30 ' Grilles Faciles
        Pzzl_Prd_Batch("P", Prd)
      Next i

      Nb = 0
      Jrn_Add_Red(Proc_Name_Get() & " Phase M.")
      For i As Integer = 1 To Nb_Max - Nb_M
        Nb += 1
        If Nb > Nb_Limite Then Exit For
        'Prd est ensuite aménagée des valeurs particulières de la difficulté demandée pour la grille
        Prd.Prd_Chat = False
        Prd.Prd_Create_Nb_Cel_Demandées = 26 ' Grilles Moyennes
        Pzzl_Prd_Batch("P", Prd)
      Next i

      Nb = 0
      Jrn_Add_Red(Proc_Name_Get() & " Phase D.")
      For i As Integer = 1 To Nb_Max - Nb_D
        Nb += 1
        If Nb > Nb_Limite Then Exit For
        'Prd est ensuite aménagée des valeurs particulières de la difficulté demandée pour la grille
        Prd.Prd_Chat = False
        Prd.Prd_Create_Nb_Cel_Demandées = 22 ' Grilles Difficiles
        Pzzl_Prd_Batch("P", Prd)
      Next i

      'Nb = 0
      'Jrn_Add_Red(Proc_Name_Get() & " Phase E.")
      'For i As Integer = 1 To Nb_Max - Nb_E
      '  Nb += 1
      '  If Nb > Nb_Limite Then Exit For
      '  'Prd est ensuite aménagée des valeurs particulières de la difficulté demandée pour la grille
      '  Prd.Prd_Chat = False
      '  Prd.Prd_Create_Nb_Cel_Demandées = 22 ' Grilles Difficiles
      '  Pzzl_Prd_Batch("P", Prd)
      'Next i

    Catch ex As Exception
      '28/05/2024 le message permet de comprendre l'arrêt anormal du traitement
      Dim MsgTit As String = Proc_Name_Get() & " " & Application.ProductName & " " & SDK_Version
      MsgBox(ex.ToString(),, MsgTit)
    End Try

    '#616 Samedi 14/02/2026
    ' il me semble que le stock ne peut pas dépasser Nb_LImite
    'Stock_Excédent_Delete("SDK_F")
    'Stock_Excédent_Delete("SDK_M")
    'Stock_Excédent_Delete("SDK_D")
    'Stock_Excédent_Delete("SDK_E")

    Batch_en_Cours = False
    Event_OnPaint_MAP = Proc_Name_Get()
    Event_OnPaint = "Global"
    Frm_SDK.Invalidate()

    'Fin du Thread Batch_Thread
  End Sub

  Public Sub Stock_Excédent_Delete(Code_FMDE As String)
    Dim Extension As String = ".Txt"
    Dim Répertoire As String = Path_Save & "Batch\"

    ' Récupère tous les fichiers commençant par Code_FMDE et ayant la bonne extension
    Dim Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(Répertoire)
                                          Where IO.Path.GetFileName(File).StartsWith(Code_FMDE) AndAlso
                                                IO.Path.GetExtension(File).Equals(Extension, StringComparison.OrdinalIgnoreCase)
                                          Order By IO.File.GetCreationTime(File) Descending
    Dim Nb_Max As Integer = My.Settings.Prf_02C_Nb_Batch_Generate
    If Files.Count() > Nb_Max Then
      Dim FilesToDelete As IEnumerable(Of String) = Files.Skip(Nb_Max)
      For Each FilePath As String In FilesToDelete
        Try
          IO.File.Delete(FilePath)
        Catch ex As Exception
          'Jrn_Add(, {"Erreur lors de la suppression de " & FilePath & " : " & ex.Message})
        End Try
      Next
    End If
  End Sub


  Public Function ProcessStarted(Name As String) As Boolean
    ' La fonction permet de vérifier si un process est lancé ou non

    Dim processes() As Process = Process.GetProcesses()
    For Each proc As Process In processes
      Try
        If proc.ProcessName = Name Then
          Return True
        End If
      Catch ex As Exception
        ' Certains processus système ne permettent pas l'accès à MainModule
        'Jrn_Add(, {"Nom       : " & proc.ProcessName & " → Accès refusé"})
      End Try
    Next proc
    Return False
  End Function

End Module
