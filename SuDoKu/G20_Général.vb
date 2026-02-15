Option Strict On
Option Explicit On

'Dans le menu projet, cliquez sur Ajouter une référence.
'     Sous l'onglet COM ,recherchez la Bibliothèque d'objets Microsoft Word, 
'     cliquez sur Sélectionner.
Imports System.Reflection                                     ' Nécessaire pour Assembly et StackFrame
Imports System.Runtime.InteropServices                        ' Nécessaire pour Marshal.SizeOf(DisplayDevice)
Imports SuDoKu.NativeMethods


Module G20_Général
  '-------------------------------------------------------------------------------
  ' Traitement 
  '-------------------------------------------------------------------------------
  Public Function JourDateHeure() As String()
    Return {StrConv(Format(Now, "dddd"), VbStrConv.ProperCase) & " " & Format(Now, "d MMM yyyy"), Format(Now, "H:mm:ss")}
  End Function

  Public Function Wh_Nb_Cell(U_temp(,) As String) As Wh_Nb_Cell_Struct
    With Wh_Nb_Cell
      ' Initialisation des ventilations
      ReDim .Val_Nb(9), .Cdd_Nb(9)
      For i As Integer = 0 To 9
        .Val_Nb(i) = 0 : .Cdd_Nb(i) = 0
      Next i
      For i As Integer = 0 To 80
        If U_temp(i, 2) = " " Then .Vides += 1
        If U_temp(i, 2) <> " " Then
          .Remplies += 1
          .Val_Nb(0) += 1            ' Somme des ventilations
          Dim Valeur As Integer = CInt(U_temp(i, 2))
          .Val_Nb(Valeur) += 1       ' Ventilation des Valeurs
        End If
        If U_temp(i, 3) <> Cnddts_Blancs Then
          For j As Integer = 0 To 8
            Dim Candidat As String = U_temp(i, 3).Substring(j, 1)
            If Candidat <> " " Then
              Dim Cdd As Integer = CInt(Candidat)
              .Cdd_Nb(0) += 1    ' Somme des ventilations
              .Cdd_Nb(Cdd) += 1  ' Ventilation des Candidats
            End If
          Next j
        End If
        If U_temp(i, 1) <> " " Then .Initiales += 1
        If U_temp(i, 2) = " " And U_temp(i, 3) = Cnddts_Blancs Then .Vides_sans_Candidats += 1
      Next i
    End With
  End Function
  Public Sub Wh_Nb_Cell_Display(Wh_Nb_Cell As Wh_Nb_Cell_Struct)
    With Wh_Nb_Cell
      Jrn_Add(, {"Cell Initiales   : " & CStr(.Initiales).PadLeft(2)})
      Jrn_Add(, {"Cell Vides       : " & CStr(.Vides).PadLeft(2)})
      Jrn_Add(, {" Sans Candidats  : " & CStr(.Vides_sans_Candidats).PadLeft(2)})
      Jrn_Add(, {"Cell Remplies    : " & CStr(.Remplies).PadLeft(2)})
      Jrn_Add(, {"Cell Trouvées    : " & CStr(.Remplies - .Initiales).PadLeft(2)})

      Dim S1 As String = "Ventilation      :  S|  1|  2|  3|  4|  5|  6|  7|  8|  9|"
      Dim S2 As String = "Valeurs          :"
      Dim S3 As String = "Candidats        :"
      For i As Integer = 0 To 9
        S2 &= CStr(.Val_Nb(i)).PadLeft(3) & "|"
        S3 &= CStr(.Cdd_Nb(i)).PadLeft(3) & "|"
      Next i
      Jrn_Add(, {S1})
      Jrn_Add(, {S2})
      Jrn_Add(, {S3})
    End With
  End Sub

  Function File_Nb(Racine As String) As Integer
    'Retourne le nombre de parties existantes
    Return (From File In IO.Directory.GetFiles(Path_Batch) Where File.Contains(Racine)).Count
  End Function

  Sub File_Nb_Size(Application As String)
    Dim T_Nb As Integer
    Dim T_Size As Double

    Try
      Dim Files As IEnumerable(Of String) = From File In IO.Directory.GetFiles(Application, "*.*", IO.SearchOption.AllDirectories)
      'Il y a récursivité, ne prend pas les fichiers non autorisés
      For Each File As String In Files
        Dim FInfo As New IO.FileInfo(File)
        T_Nb += 1                 ' Nombre
        T_Size += FInfo.Length    ' Size et non la longueur de la chaîne "Nom du fichier"
      Next File
      Dim MT_Size As Double = ((T_Size / 1024) / 1024)
      MT_Size = CDbl(FormatNumber(MT_Size, 2))
      Jrn_Add(, {Application.PadRight(60) & " " & CStr(T_Nb).PadLeft(8) & " " & CStr(MT_Size).PadLeft(8)})
    Catch ex As Exception
      Jrn_Add(, {Application & " ... Non autorisé"})
      Jrn_Add("ERR_00000", {ex.Message}, "Erreur")
      Jrn_Add("ERR_00000", {ex.ToString()}, "Erreur")
    End Try
  End Sub
  Sub Pzzl_Load(ByRef Nom_Pzzl As String)
    'Chargement d'une grille 
    'si le jeu est en Sas, la partie est chargée en Sas
    For i As Integer = 0 To Pzzl_List.Count - 1
      If Pzzl_List.Item(i).Nom = Nom_Pzzl Then
        Game_New_Game(Gnrl:=Plcy_Gnrl, Nom:=Pzzl_List.Item(i).Nom, Prb:=Pzzl_List.Item(i).Problème, Jeu:=Pzzl_List.Item(i).Problème, Sol:=Pzzl_List.Item(i).Solution, Cdd729:=StrDup(729, " "), Frc:=Pzzl_List.Item(i).Force)
        Exit For
      End If
    Next i
    'La partie indiquée n'a pas été trouvée, peu probable sauf en cas de MàJ du fichier
    If LP_Nom = "" Then
      Jrn_Add("SDK_00111", {Proc_Name_Get(), Nom_Pzzl})
      Exit Sub
    End If
  End Sub

  '-------------------------------------------------------------------------------
  '
  ' Procédures générales
  '
  '-------------------------------------------------------------------------------
  Public Sub Get_AllMenuItems(Menu As MenuStrip)
    'La procédure liste les options et sous-options d'un menu
    'Exemple Get_AllMenuItems(Frm_SDK.Mnu)
    'https://stackoverflow.com/questions/64369013/how-to-extract-all-toolstripmenuitems-of-a-contextmenustrip-using-recursion
    For Each Item As ToolStripMenuItem In Menu.Items.OfType(Of ToolStripMenuItem)
      Jrn_Add(, {"       " & (Mid$((Item.Name).PadRight(30), 1, 30) & " " & Item.Text)})
      For Each SubItem As ToolStripMenuItem In Get_SubMenuItems(Item)
        Jrn_Add(, {"         " & (Mid$((SubItem.Name).PadRight(30), 1, 30) & " " & SubItem.Text)})
      Next SubItem
    Next Item
  End Sub
  Public Iterator Function Get_SubMenuItems(Parent As ToolStripMenuItem) As IEnumerable(Of ToolStripMenuItem)
    For Each Item As ToolStripMenuItem In Parent.DropDownItems.OfType(Of ToolStripMenuItem)
      If Item.HasDropDownItems Then
        Yield Item
        For Each SubItem As ToolStripMenuItem In Get_SubMenuItems(Item)
          Yield SubItem
        Next SubItem
      Else
        Yield Item
      End If
    Next Item
  End Function
  Public Sub Get_AllCtxtMenuItems(Menu As ContextMenuStrip)
    'La procédure liste les options et sous-options d'un menu contextuel
    'Exemple Get_AllCtxtMenuItems(Frm_SDK.Mnu_Cel)
    Dim Item As ToolStripItem
    For Each Item In Menu.Items
      Jrn_Add(, {(Mid$((Item.Name).PadRight(30), 1, 30) & " " & Item.Text)})
    Next Item
  End Sub
  Public Sub Get_AllBarreOutilsItems(Menu As ToolStrip)
    'La procédure liste les options et sous-options d'une Barre d'Outils  
    'Exemple Get_AllBarreOutilsItems(Frm_SDK.BarreOutils)
    Dim Item As ToolStripItem
    For Each Item In Menu.Items
      Jrn_Add(, {(Mid$((Item.Name).PadRight(30), 1, 30) & " " & Item.Text)})
    Next Item
  End Sub
  Sub Get_AllForms()
    Dim FormTypes As List(Of Type) = Get_FormTypes()
    For Each FormType As Type In FormTypes
      Jrn_Add(, {FormType.Name})
    Next FormType
  End Sub
  Function Get_FormTypes() As List(Of Type)
    Dim FormTypes As New List(Of Type)
    Dim Assembly As Assembly = Assembly.GetExecutingAssembly()
    For Each Type As Type In Assembly.GetTypes()
      If Type.IsSubclassOf(GetType(Form)) Then
        FormTypes.Add(Type)
      End If
    Next Type
    Return FormTypes
  End Function
  Sub Get_AllFormMenus(Frm As Form)
    For Each Ctl As Control In Frm.Controls
      If Ctl.GetType.ToString() = "System.Windows.Forms.MenuStrip" Then
        Jrn_Add(, {Ctl.ToString()})
      End If
    Next Ctl
  End Sub
  Sub Get_All_Forms_MenuStrip_MenuItems()
    Dim Assembly As Assembly = Assembly.GetExecutingAssembly()
    ' 1 Récupération à/p de l'Assembly des Formulaire; ils sont stockés dans une List
    Dim FormsList As New List(Of Type)
    ' Lister tous les types de formulaires
    For Each Type As Type In Assembly.GetTypes()
      If Type.IsSubclassOf(GetType(Form)) Then
        FormsList.Add(Type)
      End If
    Next Type

    ' 2  Traitement de chaque Formulaire 
    For Each FormType As Type In FormsList
      ' 3 Création d'une instance du formulaire
      Dim FormInstance As Form = CType(Activator.CreateInstance(FormType), Form)
      Jrn_Add(, {"F  " & FormInstance.Name})
      ' Parcourir les contrôles du formulaire pour trouver les MenuStrip
      For Each Control As Control In FormInstance.Controls
        If TypeOf Control Is MenuStrip Then
          Dim MenuStrip As MenuStrip = CType(Control, MenuStrip)
          Jrn_Add(, {"M    " & MenuStrip.Name})
          Get_AllMenuItems(MenuStrip)
        End If
      Next Control
    Next FormType
  End Sub

  '-------------------------------------------------------------------------------
  ' Gestion des fichiers INI
  '-------------------------------------------------------------------------------
  Function Ini_Read(Section As String, Poste As String) As String
    'Cette fonction retourne la valeur du poste ou ? issue du fichier des Valeurs d'Usine
    Dim Valeur_def As String = ""
    Try
      Dim Ch As New System.Text.StringBuilder(NativeMethods.MAX_ENTRY)
      Dim l As Long = NativeMethods.GetPrivateProfileString(Section, Poste, Valeur_def, Ch, NativeMethods.MAX_ENTRY, File_ValUsi)
      Return Ch.ToString()
    Catch
      Return Valeur_def
    End Try
  End Function
  Public Sub Ini_Write(Section As String, Poste As String, Valeur As String)
    'Cette procédure met à jour la valeur du poste
    NativeMethods.WritePrivateProfileString(Section, Poste, Valeur, File_ValUsi)
  End Sub

  '-------------------------------------------------------------------------------
  ' Gestion de A ( 11, 1000) as string
  ' N°                      4
  ' Action                 10-15
  ' Cellule                 2 N°
  ' Cellule                 5 Lx-Cy
  ' Valeur                  1
  ' Candidats              10
  ' Commentaire             V
  ' Jeu av                 81
  ' Tous_les_Candidats av 729
  ' Jeu ap                 81
  ' Tous_les_Candidats ap 729
  '-------------------------------------------------------------------------------
  Function Act_Build_Ligne(i As Integer) As String
    Dim S As String
    S = "| " & Act(4, i) &
        " | " & Act(5, i) &
        " | " & Act(6, i) &
        " | " & Act(1, i).PadLeft(9) &
        " | " & Act(2, i).PadRight(15).Substring(0, 14) &
        " | " & Act(7, i).PadRight(20)
    Act_Build_Ligne = S
  End Function
  Sub Act_Add(Cellule As Integer,
                 Action As String,
                 Valeur As String,
                 Candidats As String,
                 Commentaire As String,
                 Av_Jeu As String,
                 Av_AllCdd As String,
       Optional Jrn As Boolean = True)
    'La procédure ajoute une ligne au tableau des actions Act
    '             affiche la ligne dans le journal
    ' 1 Gestion de Act (n 100, 11) as string
    Act_Index += 1          'Incrémentation de l'index
    If Act_Index > UBound(Act, 2) Then ReDim Preserve Act(11, UBound(Act, 2) + 100) 'Redimensionnement de A par tranche de 100

    ' 2 Remplissage de  
    Act(1, Act_Index) = CStr(Act_Index)
    Act(2, Act_Index) = Action
    Select Case Cellule
      Case -1
        Act(3, Act_Index) = "_"
        Act(4, Act_Index) = "_____"
      Case -2 'Espace
        Act(3, Act_Index) = " "
        Act(4, Act_Index) = "     "
      Case Else
        Act(3, Act_Index) = CStr(Cellule)
        Act(4, Act_Index) = U_cr(Cellule)
    End Select
    Act(5, Act_Index) = Valeur
    Act(6, Act_Index) = Candidats
    Act(7, Act_Index) = Trim(Commentaire)
    'Sont ajoutés le problème en cours, ie LP_Jeu et les 729 Candidats AVANT 
    Act(8, Act_Index) = Av_Jeu
    Act(9, Act_Index) = Av_AllCdd
    'Sont ajoutés le problème en cours, ie LP_Jeu et les 729 Candidats APRES 
    Act(10, Act_Index) = Act_Jeu()
    Act(11, Act_Index) = Act_Candidats()

    ' 3 Gestion du menu
    If Game_Undo_Redo = "Normal" Then
      UR_Nb += 1
      UR_Nb_Annuler = 0
      UR_Nb_Refaire = 0
    Else
      UR_Nb = 0
    End If
    'Dès qu'il y a un mouvement, le menu Annuler doit être accessible
    If UR_Nb > 1 Then Frm_SDK.Mnu02_Annuler.Enabled = True
    If (UR_Nb_Annuler = 0 And UR_Nb_Refaire = 0) Then Frm_SDK.Mnu02_Refaire.Enabled = False

    ' 4 Affichage dans le journal 
    If Jrn = True Then Jrn_Add(, {Act_Build_Ligne(Act_Index)}, "Insertion")
  End Sub
  Sub Act_Display()
    Jrn_Add("SDK_30040")
    Jrn_Add("SDK_30041", {CStr(LBound(Act, 2)), CStr(UBound(Act, 2)), CStr(Act_Index)})
    Jrn_Add("SDK_30042")
    Jrn_Add("SDK_30043")
    Jrn_Add("SDK_30044")
    For i As Integer = LBound(Act, 2) To UBound(Act, 2)
      If i > Act_Index Then Exit For
      Jrn_Add(, {Act_Build_Ligne(i)})
    Next i
    Jrn_Add("SDK_30044")
  End Sub
  Function Act_Jeu() As String
    Act_Jeu = ""
    For i As Integer = 0 To 80 : Act_Jeu &= U(i, 2) : Next i
  End Function
  Function Act_Candidats() As String
    Act_Candidats = ""
    For i As Integer = 0 To 80 : Act_Candidats &= U(i, 3) : Next i
  End Function

  Function TTT_MEF_Cdd(Candidats As String) As String
    Dim MEF As String = ""
    For i As Integer = 0 To 8
      ' Permet d'avoir des valeurs alignées, même avec une police proportionnelle
      'Dim Cdd As String = If(Candidats.Substring(i, 1) <> " ", " " & Candidats.Substring(i, 1), "   ")
      Dim Cdd As String = If(Candidats.Substring(i, 1) <> " ", " " & Candidats.Substring(i, 1), "  ")
      'la valeur de esp ne change en rien la taille de TTT
      'Dim esp As String = "        "
      Dim esp As String = ""
      Select Case i
        Case 0 : MEF &= esp & Cdd
        Case 1 : MEF &= Cdd
        Case 2 : MEF &= Cdd & vbCrLf
        Case 3 : MEF &= esp & Cdd
        Case 4 : MEF &= Cdd
        Case 5 : MEF &= Cdd & vbCrLf
        Case 6 : MEF &= esp & Cdd
        Case 7 : MEF &= Cdd
        Case 8 : MEF &= Cdd
        Case Else : Exit Select
      End Select
    Next i
    Return MEF
  End Function


  '-------------------------------------------------------------------------------
  ' Résolutions
  '-------------------------------------------------------------------------------

  Public Function Get_Scale_IA(Device_Number As Integer) As PointF
    ' Paire ordonnée x et y en virgule flottante pour définir l'échelle personnalisée
    Dim Scale_Personnalisée As New PointF(1.0, 1.0)
    Dim Screens As Screen() = Screen.AllScreens
    'Screens.count = 1 s'il n'y a pas d'écran connecté. 
    'Screens.count = 2 si le BENQ est connecté (en fonction ou éteint)
    'SystemInformation VirtualScreen {X=0,Y=0,Width=4480,Height=1080}
    '                                donne la taille de la largeur totale  
    '                                2560 + 1440 = 4480 
    With Screens(Device_Number)
      '.Bounds.propose la dimension "agrandie"
      ' Get_Résolution_Physique_IA(.DeviceName). propose le dimension physique réelle
      Scale_Personnalisée.X = CSng(Get_Résolution_Physique_IA(.DeviceName).X / .Bounds.Width)
      Scale_Personnalisée.Y = CSng(Get_Résolution_Physique_IA(.DeviceName).Y / .Bounds.Height)
    End With
    Return Scale_Personnalisée
  End Function

  Public Function Get_Résolution_Physique_IA(Device_Name As String) As PointF
    ' Donne la Résolution Physique Réelle de l'écran
    Dim Résolution_Physique As New PointF(-1, -1)
    Dim DisplayDevice As New NativeMethods.DISPLAY_DEVICE()
    DisplayDevice.cb = Marshal.SizeOf(DisplayDevice)
    Dim DevMode As New NativeMethods.DEVMODE()
    Dim i As Integer = 0

    ' Vérification du paramètre Device_Name
    ' Throw lève une exception.
    ' Try   intercepte et gère les exceptions qui peuvent se produire pendant l’exécution d’un bloc de code.
    If String.IsNullOrWhiteSpace(Device_Name) Then
      Throw New ArgumentException("Le nom du périphérique ne peut pas être nul ou vide.", NameOf(Device_Name))
    End If

    ' Tentative d'énumération des périphériques d'affichage
    Try
      While NativeMethods.EnumDisplayDevices(Nothing, i, DisplayDevice, 0)
        If NativeMethods.EnumDisplaySettings(DisplayDevice.DeviceName, NativeMethods.ENUM_CURRENT_SETTINGS, DevMode) Then
          ' Vérifie si le périphérique actuel correspond à Device_Name
          If DisplayDevice.DeviceName = Device_Name Then
            Résolution_Physique.X = DevMode.dmPelsWidth
            Résolution_Physique.Y = DevMode.dmPelsHeight
            Return Résolution_Physique
          End If
        End If
        i += 1
      End While
    Catch ex As Exception
      ' Gestion des erreurs
      Jrn_Add(, {"Erreur lors de la récupération de la résolution physique : " & ex.Message})
    End Try

    ' Retourne la valeur par défaut si l'écran spécifié n'est pas trouvé
    Return Résolution_Physique
  End Function

  Public Function Proc_Name_Get() As String
    Dim stackFrame As New StackFrame(1)
    Dim methodBase As MethodBase = stackFrame.GetMethod()
    Dim declaringType As Type = methodBase.DeclaringType
    Return methodBase.Name
  End Function
  Public Sub Processing_Start(File As String)
    Try
      If Not System.IO.File.Exists(File) Then      ' Vérifie si le fichier existe
        MsgBox("Le fichier suivant n'existe pas : " & vbCrLf & File)
        Exit Sub
      End If

      Using Processing As New Diagnostics.Process()
        Processing.StartInfo.WindowStyle = ProcessWindowStyle.Maximized
        Processing.StartInfo.FileName = File
        Processing.Start()
      End Using
    Catch ex As Exception
      MsgBox("L'application ci-dessous doit être installée." & vbCrLf & File)
    End Try
  End Sub

  Public Function Subst_Police(Source As String) As String
    'Source est compris entre 1 et 9
    'Donc Subst_Police(Cstr(0)) retourne vide et le bouton n'affiche rien
    Dim Cible As String = String.Empty
    If Source >= "1" And Source <= "9" Then
      Dim V As Integer = CInt(Source)
      Select Case Plcy_Fantasy_Name
        Case "Arial" : Cible = Subst_Arial_____(V)
        Case "Wingdings" : Cible = Subst_Wingding__(V)
        Case "MS Outlook" : Cible = Subst_MS_Outlook(V)
        Case "Webdings" : Cible = Subst_Webdings__(V)
        Case Else : Cible = Subst_Arial_____(V)
      End Select
    End If
    Return Cible
  End Function



End Module