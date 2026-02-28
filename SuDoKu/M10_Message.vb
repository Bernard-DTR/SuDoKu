Option Strict On
Option Explicit On

Imports System.IO
Imports System.Threading
Imports Word = Microsoft.Office.Interop.Word 'Pour enregistrer en RTF le Journal


'-------------------------------------------------------------------------------
' Gestion des Messages
' V()  as string comporte les variables %0,%1,...,%9
' Ve() as string

'Msg_01000 = Définition du message avec des variables %0, %1 à %9; éventuellement %vbcrlf permet d'aller à la ligne.%Sp-xy pour placer xy caractères blancs
'Le message commence immédiatement à droite du signe égal
'Utilisation d'un dictionnaire des messages pour ne plus utiliser MSInternalLibrary.NativeMethods.GetPrivateProfileString
'-------------------------------------------------------------------------------

Module M10_Message

  ''' <summary>Retourne un message documenté des variables.</summary>
  Public Function Msg_Read_IA(MsgId As String, Optional V() As String = Nothing) As String
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

  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
  '
  ' Journal
  '
  '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  '
  ''' <summary>Ajoute une ligne de message dans le journal.</summary>
  ''' <param name="MsgId">Identification du message.</param>
  ''' <param name="V">Tableau optionel des variables du message.</param>
  ''' <param name="Style">Affichage du message ("Insertion","Annuler","Info","Orange","Erreur",Vide,"Italique").</param>
  Public Sub Jrn_Add(Optional MsgId As String = "SDK_00000", Optional V() As String = Nothing, Optional Style As String = "")
    '09/07/2023 MsgId est optionnel
    'Insère une information dans le journal RTF
    '-------------------------------------------'
    ' La procédure NE DOIT PAS comporter Jrn_Add '
    '-------------------------------------------'
    If Thread.CurrentThread.IsBackground Then
      ' La fonction est exécutée dans un traitement d'arrière-plan
      Exit Sub
    End If

    Dim Inf As String = "#Nothing"
    If Msg_Dsp_MsgId Then
      Inf = MsgId & " " & Msg_Read_IA(MsgId, V)
    Else
      Inf = Msg_Read_IA(MsgId, V)
    End If

    With Frm_SDK.Journal
      'La taille du journal est définie dans Frm_SDK_Load à .MaxLength = 262144 
      If .Text.Length >= 261000 Then
        Dim Jrn() As String = .Lines
        Dim File_SDK As String = Jrn_Rcd(Jrn)
        .Clear()
        .AppendText("Le journal a été enregistré sous : " & vbCrLf & File_SDK)
        .AppendText(Environment.NewLine)
        .AppendText(Format(Now, "dddd d MMM yyyy") & "; à " & DateAndTime.TimeOfDay & ".")
        .AppendText(Environment.NewLine)
      End If
      'S'il est fait une sélection dans le journal, alors la ligne suivante perd son style
      ' et la sélection change de style !
      Select Case Style
        Case ""             'Affichage standard d'une ligne dans le journal
          .SelectionFont = New Font(Font_Journal, FontStyle.Regular)
          .SelectionColor = Color.Black
          .SelectionBackColor = Nothing
        Case "Insertion"    'Insertion d'une valeur
          .SelectionFont = New Font(.Font, FontStyle.Italic)
          .SelectionColor = Color.Red
        Case "Refaire"      'Insertion d'une valeur
          .SelectionFont = New Font(.Font, FontStyle.Italic)
          .SelectionColor = Color.Blue
        Case "Annuler"      'Insertion d'une valeur
          .SelectionFont = New Font(.Font, FontStyle.Italic)
          .SelectionColor = Color.Green
        Case "Info"         'Pour info 
          .SelectionFont = New Font(.Font, FontStyle.Regular)
          .SelectionColor = Color.Black
          .SelectionBackColor = Color.Yellow
        Case "Yellow"       'Pour info 
          .SelectionFont = New Font(.Font, FontStyle.Regular)
          .SelectionColor = Color.Black
          .SelectionBackColor = Color.Yellow
        Case "Orange"       'Pour info 
          .SelectionFont = New Font(.Font, FontStyle.Regular)
          .SelectionColor = Color.Black
          .SelectionBackColor = Color.Orange
        Case "White"       'Pour info 
          .SelectionFont = New Font(.Font, FontStyle.Regular)
          .SelectionColor = Color.Black
          .SelectionBackColor = Color.White
        Case "Blue"       'Pour info 
          .SelectionFont = New Font(.Font, FontStyle.Italic)
          .SelectionColor = Color.Blue
          .SelectionBackColor = Color.White
        Case "Red"       'Pour info 
          .SelectionFont = New Font(.Font, FontStyle.Italic)
          .SelectionColor = Color.Red
          .SelectionBackColor = Color.White
        Case "Italique"    'Pour info, en italique
          .SelectionFont = New Font(.Font, FontStyle.Italic)
          .SelectionColor = Color.Black
        Case "Erreur"      'Affichage d'une erreur 
          .SelectionFont = New Font(.Font, FontStyle.Bold)
          .SelectionColor = Color.Black
          .SelectionBackColor = Color.Red
        Case Else          'Affichage standard d'une ligne dans le journal
          .SelectionFont = New Font(.Font, FontStyle.Regular)
          .SelectionColor = Color.White
          .SelectionBackColor = Color.Yellow
          Inf = Inf & " Style Inconnu : /" & Style & "/"
      End Select
      .AppendText(Inf)
      .AppendText(Environment.NewLine)
      '.ScrollToCaret()
      'Fait défiler le contenu du contrôle vers la position indiquée par le signe insertion.
      'Cette méthode n’a aucun effet si le contrôle n’a pas le focus ou si le signe insertion est déjà positionné dans la zone visible du contrôle.
      'Swt_DéroulerJournal = 1 le texte défile
      '                     -1 le contrôle est bloqué
      Select Case Swt_DéroulerJournal
        Case 1
          .SelectionStart = .Text.Length
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
  ''' <summary>Efface le journal.</summary>
  Public Sub Jrn_Clear()
    Frm_SDK.Journal.Text = ""
    Jrn_Add("SDK_00011", JourDateHeure())
  End Sub
  ''' <summary>Enregistre le journal.</summary>
  Public Function Jrn_Rcd(Jrn() As String) As String
    Dim File_SDKJrn_Name As String = File_SDKJrn & Format(Now, "yyyy_MM_d_hh_mm_ss") & ".txt"
    Try
      Dim File As System.IO.FileStream = System.IO.File.Create(File_SDKJrn_Name)
      Dim WT_lig0 As String = File_SDKJrn_Name
      Dim WLigne0 As Byte() = New System.Text.UTF8Encoding(True).GetBytes(WT_lig0 & vbCrLf)
      File.Write(WLigne0, 0, WLigne0.Length)

      For i As Integer = 0 To Jrn.GetUpperBound(0)
        Dim WLigne2 As Byte() = New System.Text.UTF8Encoding(True).GetBytes(Jrn(i) & vbCrLf)
        File.Write(WLigne2, 0, WLigne2.Length)
      Next i
      File.Close()
    Catch ex As Exception
      Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
      Jrn_Add("ERR_00000", {ex.Message})
      Jrn_Add("ERR_00000", {ex.ToString()})
    End Try
    Return File_SDKJrn_Name
  End Function

  Public Function Jrn_RcdRTF() As String
    Cursor.Current = Cursors.WaitCursor
    Dim aWord As Word._Application
    aWord = CType(CreateObject("Word.Application"), Word.Application)
    Dim dWord As Word._Document = aWord.Documents.Add()
    Dim ObjName As Object = File_SDKJrn & Format(Now, "yyyy_MM_d_hh_mm_ss") & ".docx"
    Try
      Clipboard.Clear() 'Copie le journal dans le clipboard en format RTF 
      Clipboard.SetData(DataFormats.Rtf, CType(Frm_SDK.Journal.Rtf, Object))
      aWord.Selection.Paste()
      'Enregistre le journal RTF dans un fichier, le nom de l'objet comporte la date et l'heure
      dWord.SaveAs(FileName:=ObjName)
      dWord.Close(True)
      aWord.Quit(True)
      Clipboard.Clear()
    Catch ex As Exception
      Jrn_Add("ERR_00000", {Proc_Name_Get()}, "Erreur")
      Jrn_Add("ERR_00000", {ex.Message})
      Jrn_Add("ERR_00000", {ex.ToString()})
    End Try
    Jrn_Add(, {"Le journal est enregistré : " & ObjName.ToString})
    Return ObjName.ToString()
    Cursor.Current = Cursors.Default
  End Function

End Module