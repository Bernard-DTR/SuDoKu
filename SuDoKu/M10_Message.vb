Imports System.Threading
' Gestion des Messages
' V()  as string comporte les variables %0,%1,...,%9

'Msg_01000 = Définition du message avec des variables %0, %1 à %9; éventuellement %vbcrlf permet d'aller à la ligne.%Sp-xy pour placer xy caractères blancs
'Le message commence immédiatement à droite du signe égal
Module M10_Message
  '===========================================================
  ' ThemeManager : gestion centralisée des polices et couleurs
  '===========================================================
  Public NotInheritable Class ThemeManager

    '--- Polices du journal ---
    Public Shared ReadOnly FontJournalRegular As Font
    Public Shared ReadOnly FontJournalItalic As Font
    Public Shared ReadOnly FontJournalBold As Font

    '--- Styles du journal ---
    Public Shared ReadOnly JournalStyles As Dictionary(Of String, (Font As Font, Fore As Color, Back As Color))

    Shared Sub New()

      ' Polices (une seule allocation → pas de fuite GDI)
      FontJournalRegular = New Font(Font_Journal, FontStyle.Regular)
      FontJournalItalic = New Font(Font_Journal, FontStyle.Italic)
      FontJournalBold = New Font(Font_Journal, FontStyle.Bold)

      JournalStyles = New Dictionary(Of String, (Font As Font, Fore As Color, Back As Color)) From
        {
        {"", (FontJournalRegular, Color.Black, Nothing)},               'Conserve le fond du thème
        {"Insertion", (FontJournalItalic, Color.Red, Nothing)},         'Conserve le fond du thème
        {"Refaire", (FontJournalItalic, Color.Blue, Nothing)},
        {"Annuler", (FontJournalItalic, Color.Green, Nothing)},
        {"Info", (FontJournalRegular, Color.Black, Color.White)},       'Fond explicite
        {"Yellow", (FontJournalRegular, Color.Black, Color.Yellow)},    'Fond explicite
        {"Orange", (FontJournalRegular, Color.Black, Color.Orange)},
        {"White", (FontJournalRegular, Color.Black, Color.White)},
        {"Blue", (FontJournalItalic, Color.Blue, Color.White)},
        {"Red", (FontJournalItalic, Color.Red, Color.White)},
        {"Italique", (FontJournalItalic, Color.Black, Nothing)},        'Conserve le fond du thème
        {"Erreur", (FontJournalBold, Color.Black, Color.Red)}           'Fond explicite
        }
    End Sub
  End Class

  ''' <summary>Retourne un message documenté des variables.</summary>
  Public Function Msg_Read(MsgId As String, Optional V() As String = Nothing) As String
    ' Cette fonction retourne un message avec des variables documentées

    Dim Ve(10) As String
    ' Initialisation des variables de remplacement avec une valeur par défaut
    For i As Integer = 0 To 9
      Ve(i) = " # "
    Next i

    ' Si un tableau de valeurs est fourni, remplacez les valeurs par défaut
    If V IsNot Nothing Then
      For i As Integer = 0 To Math.Min(UBound(V), 9)
        Ve(i) = V(i)
      Next i
    End If

    ' Valeurs spéciales
    If MsgId = "SDK_Space" Then Return " "

    Try
      Dim M As String = Msg_Dcty.Item(MsgId)

      ' Insertion des variables
      For i As Integer = 0 To 9
        Dim Repl As String = "%" & i.ToString()
        If Ve(i) IsNot Nothing Then M = M.Replace(Repl, Ve(i))
      Next i

      ' Insertion des valeurs spéciales
      For i As Integer = 1 To 9
        M = M.Replace("%sp-" & i.ToString() & " ", Space(i))
      Next i
      M = M.Replace("%vbcrlf ", vbCrLf)

      Return M
    Catch ex As KeyNotFoundException
      Return "#" & MsgId & "#_ La clé donnée est absente du dictionnaire."
    End Try
  End Function

  ''' <summary>Affiche le fichier des Messages.</summary>
  Public Sub Msg_Display()
    Dim V(10) As String
    Dim i As Integer
    For j As Integer = 0 To 9 : V(j) = Nothing : Next j
    Jrn_Add("SDK_Space")
    Jrn_Add("SDK_00090", {File_SDKMsg})
    Try
      Dim Rcd As New IO.StreamReader(File_SDKMsg, Text.Encoding.UTF7)
      Dim C As String
      Do While Rcd.Peek() >= 0   'Peek: Retourne le prochain caractère disponible, mais ne le consomme pas.
        C = Rcd.ReadLine()
        If C.Contains("=") = True Then i += 1
        Jrn_Add(, {C, V(0), V(1), V(2), V(3), V(4), V(5), V(6), V(7), V(8), V(9)})
      Loop
      Rcd.Close()
    Catch ex As Exception
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
    Jrn_Add("SDK_00080", {File_SDKMsg, CStr(i)})
    Jrn_Add("SDK_Space")
  End Sub

  '===========================================================
  ' Jrn_Add : version optimisée avec ThemeManager
  '===========================================================
  ''' <summary>Ajoute une ligne de message dans le journal.</summary>
  ''' <param name="MsgId">Identification du message.</param>
  ''' <param name="V">Tableau optionel des variables du message.</param>
  ''' <param name="Style">Affichage du message ("Insertion","Annuler","Info","Orange","Erreur",Vide,"Italique").</param>
  Public Sub Jrn_Add(Optional MsgId As String = "SDK_00000",
                   Optional V() As String = Nothing,
                   Optional Style As String = "")

    Dim Inf As String
    'Insère une information dans le journal RTF
    '--------------------------------------------'
    ' La procédure NE DOIT PAS comporter Jrn_Add '
    '--------------------------------------------'
    If Thread.CurrentThread.IsBackground Then
      ' La fonction est exécutée dans un traitement d'arrière-plan
      Exit Sub
    End If

    If Msg_Dsp_MsgId Then
      Inf = MsgId & " " & Msg_Read(MsgId, V)
    Else
      Inf = Msg_Read(MsgId, V)
    End If

    With Frm_SDK.Journal
      '--- Purge si trop volumineux ---
      If .Rtf.Length >= 3000000 Then  ' soit une dizaine de copies de la grille 
        Dim File_SDK As String = Path_SDK & "S50_SDK\Journal_" & Format(Now, "yyyyMMdd_HHmmss") & ".rtf"
        IO.File.WriteAllText(File_SDK, Frm_SDK.Journal.Rtf)
        'Pour recharger le fichier RTF dans le journal
        'Frm_SDK.Journal.LoadFile(File_SDK, RichTextBoxStreamType.RichText)
        .Clear()
        .AppendText("Le journal a été enregistré sous : " & vbCrLf & File_SDK & vbCrLf)
        .AppendText(Format(Now, "dddd d MMM yyyy") & "; à " & DateAndTime.TimeOfDay & "." & vbCrLf)
      End If

      '--- Toujours écrire à la fin ---
      .SelectionStart = .TextLength
      .SelectionLength = 0

      '--- Style via ThemeManager ---
      Dim st As (Font As Font, Fore As Color, Back As Color)
      If ThemeManager.JournalStyles.ContainsKey(Style) Then
        st = ThemeManager.JournalStyles(Style)
      Else
        st = (ThemeManager.FontJournalRegular, Color.White, Color.Yellow)
        Inf &= $" Style Inconnu : /{Style}/"
      End If

      .SelectionFont = st.Font
      .SelectionColor = st.Fore
      If st.Back = Nothing Then
        .SelectionBackColor = Nothing   'Conserve le fond du thème
      Else
        .SelectionBackColor = st.Back   'Applique un fond explicite
      End If

      '--- Ajout de la ligne ---
      .AppendText(Inf & Environment.NewLine)

      '--- Défilement ---
      'Fait défiler le contenu du contrôle vers la position indiquée par le signe insertion.
      'Cette méthode n’a aucun effet si le contrôle n’a pas le focus ou si le signe insertion est déjà positionné dans la zone visible du contrôle.
      'Swt_DéroulerJournal = 1 le texte défile
      '                     -1 le contrôle est bloqué
      Select Case Swt_DéroulerJournal
        Case 1
          .SelectionStart = .TextLength
          .ScrollToCaret()
        Case -1
          .SelectionStart = Journal_Emp_Blocage
      End Select
    End With
  End Sub
  Public Sub Jrn_Add_Yellow(V As String)
    'Affichage rapide d'une information en jaune
    Jrn_Add(, {V}, "Yellow")
  End Sub
  Public Sub Jrn_Add_Orange(V As String)
    'Affichage rapide d'une information en orange
    Jrn_Add(, {V}, "Orange")
  End Sub
  Public Sub Jrn_Add_White(V As String)
    'Affichage rapide d'une information en blanc
    Jrn_Add(, {V}, "White")
  End Sub
  Public Sub Jrn_Add_Blue(V As String)
    'Affichage rapide d'une information en blanc
    Jrn_Add(, {V}, "Blue")
  End Sub
  Public Sub Jrn_Add_Red(V As String)
    'Affichage rapide d'une information en blanc
    Jrn_Add(, {V}, "Red")
  End Sub
  Public Sub Jrn_Exemple()
    Jrn_Add_Yellow("Jrn_Add_Yellow affiche un message en jaune")
    Jrn_Add_Orange("Jrn_Add_Orange affiche un message en orange")
    Jrn_Add_White("Jrn_Add_White  affiche un message en blanc")
    Jrn_Add_Blue("Jrn_Add_Blue   affiche un message en bleue")
    Jrn_Add_Red("Jrn_Add_Red    affiche un message en rouge")
    Jrn_Add(, {"Aucun style        affichage standard"})
    Jrn_Add(, {"Insertion          Italic  et rouge                "}, "Insertion")
    Jrn_Add(, {"Refaire            Italic  et bleu                 "}, "Refaire")
    Jrn_Add(, {"Annuler            Italic  et vert                 "}, "Annuler")
    Jrn_Add(, {"Info               Regular et noir  / blanc        "}, "Info")
    Jrn_Add(, {"Yellow             Regular et noir  / Jaune        "}, "Yellow")
    Jrn_Add(, {"Orange             Regular et noir  / Orange       "}, "Orange")
    Jrn_Add(, {"White              Regular et noir  / blanc        "}, "White")
    Jrn_Add(, {"Blue               Italic  et bleu  / blanc        "}, "Blue")
    Jrn_Add(, {"Red                Italic  et rouge / blanc        "}, "Red")
    Jrn_Add(, {"Italique           Italic  et blanc                "}, "Italique")
    Jrn_Add(, {"Erreur             Bold    et blanc / rouge        "}, "Erreur")
    Jrn_Add(, {"Else               Regular et blanc / jaune        "}, "Else")
  End Sub
  Public Function Jrn_RcdRTF() As String
    Dim File_SDK As String = Path_SDK & "S50_SDK\Journal_" & Format(Now, "yyyyMMdd_HHmmss") & ".rtf"
    IO.File.WriteAllText(File_SDK, Frm_SDK.Journal.Rtf)
    With Frm_SDK.Journal
      .Clear()
      .AppendText("Le journal a été enregistré sous : " & vbCrLf & File_SDK & vbCrLf)
      .AppendText(Format(Now, "dddd d MMM yyyy") & "; à " & DateAndTime.TimeOfDay & "." & vbCrLf)
    End With
    Return File_SDK
  End Function
End Module