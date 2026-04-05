Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading.Tasks

'-------------------------------------------------------------------------------
' HPOmen 16 dispose de 8 coeurs et 16 processeurs logiques
'-------------------------------------------------------------------------------
Friend Module V01_CodeAnalyse
  Public Sub Sql_SuDoKu()
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})

    Sql_Truncate()
    Sql_Insert_CodT()
    Sql_Insert_MotT()
    Sql_Insert_PrcT()
    Sql_CountParallel_PrcT()
    Sql_CountParallel_MsgT()
    Sql_CountSequentiel_PrmT()
    Dim DE As String = "Date d'extraction: " & Format(Now, "dddd d MMM yyyy") & " " & Format(Now, "H:mm:ss")
    My.Settings.Date_Extraction = DE

    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add("SDK_Space")
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Sub Sql_Truncate()
    Jrn_Add("SDK_Space")
    Dim Sql_Str As String = ""
    Dim Tables As String() = {"CodT", "MotT", "PrcT", "MsgT", "PrmT"}
    Using Sql_Con As New SqlConnection(My.Settings.Connect)
      Sql_Con.Open()
      For Each Table As String In Tables
        Sql_Str = "Truncate Table " & Table
        Using Sql_Cmd As New SqlCommand(Sql_Str, Sql_Con)
          Sql_Cmd.ExecuteNonQuery()
        End Using
        Jrn_Add(, {Sql_Str})
      Next Table
    End Using
  End Sub
  Public Sub Sql_Insert_CodT()
    ' Insère le Code et les Commentaires et détermine éventuellement le Nom de la procédure
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Application_SDK As String = Path_SDK & "SuDoKu\SuDoKu"

    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})

    Dim Sql_Str As String = ""
    Dim Sql_Con As New SqlConnection(My.Settings.Connect)
    Dim Sql_Cmd As New SqlCommand
    Sql_Cmd.CommandType = Data.CommandType.Text
    Sql_Cmd.Connection = Sql_Con

    Dim File_Nb As Integer = 0
    Dim Ligne_Nb As Integer
    Dim Ligne_Mod As String
    Dim Ligne As String
    Dim Ligne_M As String
    Dim Procédure As String = ""
    Dim Procédure_RAZ As Boolean = False
    Dim Nb_Erreurs As Integer
    Dim Phase As String = ""
    Sql_Con.Open()

    Dim files As IEnumerable(Of String) = From file In Directory.EnumerateFiles(Application_SDK, "*.vb", SearchOption.AllDirectories)
    For Each File As String In files
      If File.Contains("AssemblyAttributes.vb") = True Then Continue For
      Dim d As Integer = Application_SDK.Length
      Dim l As Integer = File.Length
      Dim File_Sht As String = Mid$(File, d + 2, l - d + 2)
      If File_Sht = "ApplicationEvents.vb" Then Continue For
      If File_Sht = "Settings.vb" Then Continue For
      If Mid$(File_Sht, 1, 11) = "My Project\" Then Continue For
      If Mid$(File_Sht, 1, 8) = "obj\x86\" Then Continue For
      File_Nb += 1
      Ligne_Nb = 0

      Procédure_RAZ = False
      Dim Module_File As StreamReader = My.Computer.FileSystem.OpenTextFileReader(File)
      Do
        'Insertion des enregistrements
        Ligne_Nb += 1
        Ligne = Module_File.ReadLine()
        If Ligne IsNot Nothing Then
          Ligne_M = Ligne
          Ligne_Mod = "" 'Ligne modifiée
          'La chaîne SQl ne peut comporter le caractère quote
          If Ligne.Contains("'") = True Then
            Ligne_Mod = "MQt" : Ligne_M = Ligne_M.Replace("'", " ")
          End If

          'Nom de la procédure
          Dim fsd As Integer
          Dim fsf As Integer
          If Procédure_RAZ = False Then Procédure = ""

          If Ligne.Trim().StartsWith("'") = False Then
            'Les lignes commentaires ne sont pas concernées 

            If Ligne.Contains(" Function ") = True And Ligne.Contains(" Exit Function ") = False And Ligne.Contains(" Declare Function ") = False And Ligne.Contains(".IndexOf") = False Then
              Procédure_RAZ = True
              Try
                Phase = "A"
                fsd = Ligne.IndexOf(" Function ", 0, Ligne.Length)
                fsf = Ligne.IndexOf("(", 0, Ligne.Length)
                If fsd <> -1 Or fsf <> -1 Then Procédure = Ligne.Substring(fsd + 10, fsf - fsd - 10)
              Catch ex As Exception
                Jrn_Add(, {File_Sht & " " & Phase & " " & Ligne})
                Jrn_Add(, {File_Sht & " " & Phase & " " & ex.Message})
                Nb_Erreurs += 1
                If Nb_Erreurs > 5 Then Exit Sub
              End Try
            End If
            If Ligne.Contains("End Function") = True And Ligne.Contains("Procédure_RAZ") = False Then Procédure_RAZ = False

            If Ligne.Contains(" Sub ") = True And Ligne.Contains(" Exit Sub ") = False And Ligne.Contains(".IndexOf") = False Then
              Procédure_RAZ = True
              Try
                Phase = "B"
                fsd = Ligne.IndexOf(" Sub ", 0, Ligne.Length)
                fsf = Ligne.IndexOf("(", 0, Ligne.Length)
                If fsd <> -1 Or fsf <> -1 Then Procédure = Ligne.Substring(fsd + 5, fsf - fsd - 5)
              Catch ex As Exception
                Jrn_Add(, {File_Sht & " " & Phase & " " & Ligne})
                Jrn_Add(, {File_Sht & " " & Phase & " " & ex.Message})
                Nb_Erreurs += 1
                If Nb_Erreurs > 5 Then Exit Sub
              End Try
            End If
            If Ligne.Contains("End Sub") = True And Ligne.Contains("Procédure_RAZ") = False Then Procédure_RAZ = False

          End If

          Sql_Str = "Insert Into CodT (Module, Ligne_No, Ligne_Mod, Procédure, Ligne) "
          Sql_Str &= "Values('" & File_Sht & "' ," & Ligne_Nb & " ,'" & Ligne_Mod & "' ,'" & Procédure & "' ,'" & Ligne_M & "')"

          Sql_Cmd.CommandText = Sql_Str
          Try
            Phase = "C"
            Nsd_i = Sql_Cmd.ExecuteNonQuery()
          Catch ex As Exception
            Jrn_Add(, {File_Sht & " " & Phase & " " & Ligne})
            Jrn_Add(, {File_Sht & " " & Phase & " " & ex.Message})
            Jrn_Add(, {File_Sht & " " & Phase & " " & Sql_Str})
            Nb_Erreurs += 1
            If Nb_Erreurs > 5 Then Exit Sub
          End Try
        End If
      Loop While Ligne IsNot Nothing
      Module_File.Close()
    Next File
    Sql_Cmd.Dispose()
    Sql_Con.Close()
    Jrn_Add(, {" Insertion de " & CStr(File_Nb) & " fichiers *vb."})

    Sql_Str = "Select Count(*) From CodT"
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})

    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Sub Sql_Insert_MotT()
    '     la PS InsertMot ajoute le mot ou incrémente le compteur 
    'Sql: Select Case Count(*) From MotT Count : 5951
    'Sql:  Select Case Count(*) From MotT Where Nombre <> -1  Count : 5951
    'Sql:    Select Case Sum(Nombre) From MotT Where Nombre <> -1  Count : 87487
    'Durée: 19984 ms.
    'Soit: 00:00:19
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})

    Dim Con_Strg As String = My.Settings.Connect
    Dim Sql_Str As String = ""
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Using Sql_Con As New SqlConnection(Con_Strg)
      Sql_Con.Open()
      Dim Sql_Sel As String = "Select Ligne From CodT"
      Dim Sql_Cmd_Sel As New SqlCommand(Sql_Sel, Sql_Con)
      Using Sql_Read As SqlDataReader = Sql_Cmd_Sel.ExecuteReader()  'Lecture de la table CodT
        While Sql_Read.Read()
          Dim Ligne As String = Sql_Read("Ligne").ToString().Trim() 'Ligne est un champ nchar(2048)
          ' Tableau des caractères et des chaînes de caractères à remplacer
          Dim unwantedChars As Char() = {""""c, "("c, "~"c, ")"c, "{"c, "}"c, "="c, "."c, ">"c, "<"c, "+"c, ","c, ";"c, ":"c, "!"c, "&"c, "#"c, "?"c, "*"c}
          Dim unwantedStrings As String() = {"__", "--", "//"}
          Ligne = Ligne.Trim()
          For Each ch As Char In unwantedChars
            Ligne = Ligne.Replace(ch, " "c)
          Next ch
          For Each str As String In unwantedStrings
            Ligne = Ligne.Replace(str, " ")
          Next str
          If Ligne = "" Then Continue While
          Dim T As String() = Split(Ligne, " ")
          For i As Integer = 0 To T.Count - 1
            T(i) = Trim(T(i))
            Dim Mot As String = T(i)
            If Mot = "" Or Mot.Length < 3 Or Mot.Length > 256 Then Continue For
            Dim Première_lettre As Char = Mot(0)
            If Not Char.IsLetter(Première_lettre) Then Continue For

            ' Il est nécessaire d'ouvrir une nouvelle connexion à cause du SqlDataReader
            Using Sql_Con_Mot As New SqlConnection(Con_Strg)
              Sql_Con_Mot.Open()
              Dim Sql_Prc As String = "InsertMot"
              'La procédure insère un mot si le mot n'existe pas déjà
              'ou incrémente le compteur Nombre si le mot existe
              Using Cmd_Mot As New SqlCommand(Sql_Prc, Sql_Con_Mot)
                Cmd_Mot.CommandType = CommandType.StoredProcedure
                Cmd_Mot.Parameters.AddWithValue("@Mot", Mot)
                Cmd_Mot.ExecuteNonQuery()
              End Using
            End Using
          Next i
        End While
      End Using
    End Using

    'Nombre de Mots
    Sql_Str = "Select Count(*) From MotT"
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})
    Sql_Str = "Select Count(*) From MotT Where Nombre <> -1 "
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})
    Sql_Str = "Select Sum(Nombre) From MotT Where Nombre <> -1 "
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})

    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Sub Sql_Insert_PrcT()
    '     Insère en une seule fois l'ensemble des procédures de CodT
    'Sql: Insert Into PrcT (Module, Procédure)
    '     Select Distinct Module, Procédure From CodT Where Procédure <> '' and Procédure <> 'New'
    'Sql: Select Case Count(*) From PrcT Count : 641
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})

    Dim Sql_Str As String = ""
    Dim Con_Strg As String = My.Settings.Connect
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Using Sql_Con As New SqlConnection(Con_Strg)
      Sql_Con.Open()
      Dim Sql_Ins As String = "Insert Into PrcT (Module, Procédure) "
      Sql_Ins &= "Select Distinct Module, Procédure From CodT Where Procédure <> '' and Procédure <> 'New'"
      Jrn_Add(, {"Sql : " & Sql_Ins})
      Using Sql_Cmd As New SqlCommand(Sql_Ins, Sql_Con)
        Sql_Cmd.ExecuteNonQuery()
      End Using
    End Using

    Sql_Str = "Select Count(*) From PrcT"
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})
    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Sub Sql_CountParallel_PrcT()
    ' Connexion avec un Max Pool Size=200
    ' Limite à 14 tâches parallèles et teste cette limite
    ' Utilise une procédure stockée
    'Sql: Select Case Count(*) From PrcT Count : 641
    'Sql:  Select Case Count(*) From PrcT Where Nombre <> -1  Count : 639

    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Sql_Str As String
    Dim Con_Strg As String = My.Settings.Connect & ";Max Pool Size=200"
    'Dim MaxParallelTasks As Integer = 14
    Dim MaxParallelTasks As Integer = Environment.ProcessorCount - 2
    Using Sql_Con As New SqlConnection(Con_Strg)
      Sql_Con.Open()
      Sql_Str = "Select Count(*) From PrcT"
      Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})

      Dim Sql_Sel As String = "Select IdPrc, Procédure FROM PrcT Order by Procédure"
      Dim Sql_Cmd As New SqlCommand(Sql_Sel, Sql_Con)
      Using Sql_Read As SqlDataReader = Sql_Cmd.ExecuteReader()  'Lecture de la table PrcTp
        Dim Tasks As New List(Of Task)()
        Try
          While Sql_Read.Read()
            Dim Procédure As String = Sql_Read("Procédure").ToString().Trim() 'Procédure est un champ nchar(64)
            Dim IdPrc As String = Sql_Read("IdPrc").ToString().Trim()

            If Tasks.Count >= MaxParallelTasks Then                       ' Vérifier si la limite de tâches est atteinte
              Task.WaitAny(Tasks.ToArray())                               ' Attendre que les tâches se terminent
              Tasks = Tasks.Where(Function(t) Not t.IsCompleted).ToList() ' Filtrer les tâches non terminées
            End If

            Tasks.Add(Task.Run(Sub()
                                 Using Sql_Con_Prc As New SqlConnection(Con_Strg)
                                   Sql_Con_Prc.Open()
                                   Dim Sql_Prc As String = "UpdatePrc"
                                   ' La procédure stockée exécute un Update avec le Select Count(*)
                                   Using Cmd_Mot As New SqlCommand(Sql_Prc, Sql_Con_Prc)
                                     Cmd_Mot.CommandType = CommandType.StoredProcedure
                                     Cmd_Mot.Parameters.AddWithValue("@Procédure", Procédure)
                                     Cmd_Mot.Parameters.AddWithValue("@IdPrc", IdPrc)
                                     Cmd_Mot.ExecuteNonQuery()
                                   End Using
                                 End Using
                               End Sub))
          End While
          ' Attendre que toutes les tâches soient terminées, indéfiniment
          Nsd_b = Task.WaitAll(Tasks.ToArray(), -1)
        Catch ex As Exception
          Jrn_Add(, {ex.Message})
        End Try
      End Using
    End Using

Sql_CountParallel_PrcT_End:
    Sql_Str = "Select Count(*) From PrcT Where Nombre <> -1 "
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})
    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Sub Sql_CountParallel_MsgT()
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Sql_Str As String
    Dim Con_Strg As String = My.Settings.Connect & ";Max Pool Size=200"
    Dim MaxParallelTasks As Integer = Environment.ProcessorCount - 2
    Dim Tasks As New List(Of Task)()

    Dim File_SDKMsg As String = Path_SDK & "S20_Initial\Msg.ini"

    Using Rcd As New System.IO.StreamReader(File_SDKMsg, System.Text.Encoding.Default)
      Dim Ligne As String = Rcd.ReadLine()
      While Ligne IsNot Nothing
        Try
          If Tasks.Count >= MaxParallelTasks Then                       ' Vérifier si la limite de tâches est atteinte
            Task.WaitAny(Tasks.ToArray())                               ' Attendre que les tâches se terminent
            Tasks = Tasks.Where(Function(t) Not t.IsCompleted).ToList() ' Filtrer les tâches non terminées
          End If

          ' Ignorer les lignes vides ou les commentaires
          If String.IsNullOrWhiteSpace(Ligne) OrElse Ligne.StartsWith(";") Then
            Ligne = Rcd.ReadLine()
            Continue While
          End If

          ' Vérifier la longueur de la ligne
          If Ligne.Length < 10 Then
            Ligne = Rcd.ReadLine()
            Continue While
          End If

          ' Extraire MsgID et Msg
          Dim MsgID As String = Ligne.Substring(0, 9)
          Dim Msg As String = Ligne.Substring(10).Trim()

          ' Remplacer les apostrophes et limiter la longueur du message
          Msg = Msg.Replace("'", " ")
          If Msg.Length > 128 Then
            Msg = Msg.Substring(0, 127)
          End If

          ' Capturer les valeurs dans les tâches
          Tasks.Add(Task.Run(Sub()
                               Dim Sql_Prc As String = "InsertMsg"
                               Using Sql_Con As New System.Data.SqlClient.SqlConnection(Con_Strg)
                                 Sql_Con.Open()
                                 'La procédure insère un message et son count(*)
                                 Using Sql_Cmd As New SqlCommand(Sql_Prc, Sql_Con)
                                   Sql_Cmd.CommandType = CommandType.StoredProcedure
                                   Sql_Cmd.Parameters.AddWithValue("@MsgId", MsgID)
                                   Sql_Cmd.Parameters.AddWithValue("@Msg", Msg)
                                   Sql_Cmd.ExecuteNonQuery()
                                 End Using
                               End Using
                             End Sub))
        Catch ex As Exception
          Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
          Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
        End Try

        Ligne = Rcd.ReadLine()
      End While
      ' Attendre que toutes les tâches soient terminées, indéfiniment
      Task.WaitAll(Tasks.ToArray())
    End Using

    Sql_Str = "Select Count(*) From MsgT"
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})
    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Sub Sql_CountSequentiel_PrmT()
    'Sql: Select Case Count(*) From PrmT Count : 53
    'Durée: 158312 ms.
    'Soit: 00:02:38

    ' Nom de la fonction pour le journal
    Jrn_Add("SDK_Space")
    Jrn_Add(, {Proc_Name_Get()})

    ' Temps de début pour mesurer la durée
    Dim Durée_Déb As Integer = CInt(NativeMethods.GetTickCount64)
    ' Chemin du fichier de configuration
    Dim File As String = Path_SDK & "SuDoKu\SuDoKu\app.config"

    ' Utilisation de "Using" pour la connexion SQL
    Using Sql_Con As New System.Data.SqlClient.SqlConnection(My.Settings.Connect)
      Sql_Con.Open()

      ' Lecture du fichier de configuration
      Using Cfg_File As StreamReader = My.Computer.FileSystem.OpenTextFileReader(File)
        Dim Ligne As String = Cfg_File.ReadLine()
        While Ligne IsNot Nothing
          Dim P1 As Integer = InStr(1, Ligne, "<setting name=")
          If P1 = 0 Then
            Ligne = Cfg_File.ReadLine()
            Continue While
          End If

          Dim P2 As Integer = InStr(P1 + 15, Ligne, Chr(34))
          If P2 > P1 Then
            Dim Paramètre As String = Mid$(Ligne, P1 + 15, P2 - P1 - 15)
            ' Exécution de la procédure stockée
            Dim Sql_Prc As String = "InsertPrm"
            Using Cmd_Mot As New SqlCommand(Sql_Prc, Sql_Con)
              Cmd_Mot.CommandType = CommandType.StoredProcedure
              Cmd_Mot.Parameters.AddWithValue("@Paramètre", Paramètre)
              Cmd_Mot.ExecuteNonQuery()
            End Using
          End If
          Ligne = Cfg_File.ReadLine()
        End While
      End Using
    End Using

    ' Nombre de lignes de Paramètres
    Dim Sql_Str As String = "Select Count(*) From PrmT"
    Jrn_Add(, {"Sql : " & Sql_Str & " Count : " & (Sql_Value_Int(Sql_Str))})

    ' Temps de fin et calcul de la durée
    Dim Durée_Fin As Integer = CInt(NativeMethods.GetTickCount64)
    Dim Durée As Integer = Durée_Fin - Durée_Déb
    Dim Ts As TimeSpan = TimeSpan.FromMilliseconds(Durée)
    Jrn_Add(, {"Durée : " & CStr(Durée).PadLeft(5) & " ms."})
    Dim S As String = String.Format("{0:00}:{1:00}:{2:00}", Ts.Hours, Ts.Minutes, Ts.Seconds)
    Jrn_Add(, {"Soit  : " & S})
    Jrn_Add(, {"/" & Proc_Name_Get()})
  End Sub
  Public Function Sql_Value_Int(Sql_Str As String) As Integer
    ' Retourne la valeur integer d'une requête
    Dim Value As Integer
    Using Sql_Con As New SqlConnection(My.Settings.Connect)
      Sql_Con.Open()
      Using Sql_Cmd As New SqlCommand(Sql_Str, Sql_Con)
        Value = Convert.ToInt32(Sql_Cmd.ExecuteScalar())
      End Using
    End Using
    Return Value
  End Function
  Public Sub Wh_Nb_Lignes_App(Application_SDK As String)
    'Exemples:
    ' Wh_Nb_Lignes_App(Path_SDK & "SuDoKu\SuDoKu")
    ' Wh_Nb_Lignes_App(Path_SDK & "S95_AutresJeux\CP_Sudoku")
    ' Wh_Nb_Lignes_App(Path_SDK & "S95_AutresJeux\Planete")
    ' Wh_Nb_Lignes_App(Path_SDK & "S95_AutresJeux\Sudoku_VisualBasic")
    Jrn_Add(, {"Liste des fichiers de code: " & Application_SDK})
    Jrn_Add(, {" Code   Cmt  Prc  Vide    Module"})
    Dim Nbl_T As Integer = 0
    Dim Nbl_T_Code As Integer = 0
    Dim Nbl_T_Desi As Integer = 0

    Dim Nbv_T As Integer = 0
    Dim Nbc_T As Integer = 0
    'EnumerateFiles travaille en parallèle dans les sous-répertoires et donne IMMEDIATEMENT accès au début des données
    'GetFiles       donne les résultats à la fin du traitement
    'Dim Files As String() = Directory.GetFiles(Application_SDK, "*.vb")
    Dim files As IEnumerable(Of String) = From file In Directory.EnumerateFiles(Application_SDK, "*.vb", SearchOption.AllDirectories)
    For Each File As String In files
      Dim NbL As Integer = 0
      Dim NbL_Code As Integer = 0
      Dim NbL_Desi As Integer = 0
      Dim Nbv As Integer = 0
      Dim Nbc As Integer = 0
      Dim Ligne As String
      Dim Module_File As StreamReader = My.Computer.FileSystem.OpenTextFileReader(File)
      Do
        Ligne = Module_File.ReadLine()
        NbL += 1
        If File.Contains("esigner.vb") Then
          NbL_Desi += 1
        Else
          NbL_Code += 1
        End If
        If Trim(Ligne).Length = 0 Then Nbv += 1
        If Mid$(Trim(Ligne), 1, 1) = "'" Then Nbc += 1
      Loop While Ligne IsNot Nothing
      Module_File.Close()
      Dim Prc As Integer = CInt(((Nbc * 100) / NbL))
      Dim d As Integer = Application_SDK.Length
      Dim l As Integer = File.Length
      Jrn_Add(, {CStr(NbL).PadLeft(5) & " " &
                           CStr(Nbc).PadLeft(5) & " " &
                           CStr(Prc).PadLeft(3) & "%  " &
                           CStr(Nbv).PadLeft(4) & "    " &
      Mid$(File, d + 2, l - d + 2)})
      Nbl_T += NbL
      Nbl_T_Code += NbL_Code
      Nbl_T_Desi += NbL_Desi
      Nbv_T += Nbv
      Nbc_T += Nbc
    Next File
    Dim Prc_T As Double = CInt(((Nbc_T * 100) / Nbl_T))
    Jrn_Add(, {CStr(Nbl_T).PadLeft(5) & " Lignes de code pour " & CStr(files.Count) & " fichiers." &
                    " dont" & CStr(Nbc_T).PadLeft(5) & " Lignes de commentaires (soit): " & CStr(Prc_T).PadLeft(3) & "%"})
    Jrn_Add(, {CStr(Nbv_T).PadLeft(5) & " Lignes vides. "})
    Jrn_Add(, {CStr(Nbl_T_Code).PadLeft(5) & " Lignes de Code. "})
    Jrn_Add(, {CStr(Nbl_T_Desi).PadLeft(5) & " Lignes Designer. "})
    Jrn_Add(, {CStr(CInt(Nbl_T / (192 * 22))).PadLeft(5) & " cahiers de 192 pages de 22 lignes."})

    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From CodT")).PadLeft(5) & " Lignes dans CodT."})
    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From DicTionnaire")).PadLeft(5) & " Lignes dans DicTionnaire."})
    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From MotT")).PadLeft(5) & " Lignes dans MotT."})
    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From MotTRéservés")).PadLeft(5) & " Lignes dans MotTRéservés."})
    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From MsgT")).PadLeft(5) & " Lignes dans MsgT."})
    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From PrcT")).PadLeft(5) & " Lignes dans PrcT."})
    Jrn_Add(, {CStr(Sql_Value_Int("Select Count(*) From PrmT")).PadLeft(5) & " Lignes dans PrmT."})
    Jrn_Add(, {"/End"})
  End Sub
  Public Sub Files_List()
    'La procédure liste les fichiers de l'application
    Dim Application_SDK As String = Path_SDK & "SuDoKu\SuDoKu"
    Dim File_nb As Integer

    Dim files As IEnumerable(Of String) = From file In Directory.EnumerateFiles(Application_SDK, "*.*", SearchOption.AllDirectories)
    Jrn_Add(, {Path_SDK & "SuDoKu\SuDoKu" & " : " & files.Count})

    For Each File As String In files
      Dim d As Integer = Application_SDK.Length
      Dim l As Integer = File.Length
      Dim File_Sht As String = Mid$(File, d + 2, l - d + 2)
      If Mid$(File_Sht, 1, 3) = ".vs" Then Continue For
      If Mid$(File_Sht, 1, 3) = "bin" Then Continue For
      If Mid$(File_Sht, 1, 11) = "My Project\" Then Continue For
      If Mid$(File_Sht, 1, 3) = "obj" Then Continue For
      File_nb += 1
      Jrn_Add(, {CStr(File_nb).PadLeft(3) & " " & File_Sht})
    Next File
  End Sub
End Module