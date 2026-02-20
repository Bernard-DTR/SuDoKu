Option Strict On
Option Explicit On

Imports System.Drawing.Text ' Nécessaire pour InstalledFontCollection() et PrivateFontCollection()
Imports System.Threading    ' Nécessaire pour Thread

'------------------------------------------------------------------------------------------
'Date de création: Samedi 16/07/2022
'Ce Module regroupe pratiquement toutes les variables globales de SDK
'------------------------------------------------------------------------------------------
Module A00_Public

#Region "00 Généralités"
  'Le nom de l'application est Application.ProductName
  Public SDK_Version As String = "V2026_02_10 #612"
  Public Phase_Démarrage_Terminée As Boolean = False
#End Region

#Region "01 Les Tailles"
  Public WH As Integer ' Width et Height Standard
  Public WHhalf As Integer = (WH \ 2)
  Public WHthird As Integer = (WH \ 3)
  Public WHquart As Integer = (WH \ 4)
  Public SI_CaptionHeight As Integer = SystemInformation.CaptionHeight   ' Hauteur de la Barre de Titre d'un formulaire
  Public Barre_Menu_Hauteur As Integer = 32
  Public Barre_Outils_Standard As Integer = SystemInformation.ToolWindowCaptionHeight ' Hauteur de la Barre d'outils d'un formulaire
  Public Barre_Outils_Hauteur As Integer = 23 ' 25

  Public Bld_Marge_LT As Integer = 5       ' Séparation entre le trait du formulaire et le trait de l'entourage de la grille
  Public Bld_Trait_1 As Integer = 1
  Public Bld_Trait_2 As Integer = 2
  Public Bld_Trait_3 As Integer = 3
  Public Bld_WH_Grid As Integer = (WH * 9) + Bld_Trait_3 + Bld_Trait_1 + Bld_Trait_1 +
                                             Bld_Trait_3 + Bld_Trait_1 + Bld_Trait_1 +
                                             Bld_Trait_3 + Bld_Trait_1 + Bld_Trait_1 +
                                             Bld_Trait_3
  Public Bld_Journal_Width As Integer = 750 ' quelque soit la taille de WH
  Public Bld_Journal_Affiché_Width As Integer
#End Region

#Region "02 Les Paths"
  'Voir G10_Présentation_SDK.vb Main_Variables_Init() pour les valeurs
  'Il est possible d'installer le répertoire SuDoKu_2026 dans n'importe quel autre répertoire

  Public Base_Folder As String = "SuDoKu_2026"

  Public Path_SDK As String
  Public File_SDK As String
  Public File_Correction As String
  Public File_SDKMsg As String
  Public File_SDKTxt As String
  Public File_Puzzle As String
  Public File_Hodoku As String
  Public File_SDKEtd As String
  Public File_SDKJrn As String
  Public Path_Batch As String
  Public Path_Batch_Poubelle As String
  Public Path_Save As String
  Public Path_SDKAJ As String
  Public File_SDKDoc As String
  Public File_ValUsi As String
  Public File_Journal As String
#End Region

#Region "03 LP Last Puzzle"
  Public LP_Nom As String = "#"
  Public LP_Prb As String = "#"
  Public LP_Jeu As String = "#"
  Public LP_Sol As String = "#"
  Public LP_Frc As String = "#"
  Public LP_Cdd As String = "#"
  Public LP_CddExc As String = "#"
#End Region

#Region "04 Constantes"
  Public Const Cnddts As String = "123456789"
  Public Const Cnddts_Blancs As String = "         "
  Public Const Les9Candidats As String = "Les9Candidats"
  Public Const LesCandidatsEligibles As String = "LesCandidatsEligibles"
#End Region

#Region "10 U et ses dérivés"
  Public U(0 To 80, 0 To 3) As String
  ' La colonne 0 ne comporte pas d'information
  '            1 comporte la Valeur Initiale
  '            2 comporte la Valeur Saisie
  '            3 comporte les Candidats
  Public Const UNbCopy As Integer = 324         ' soit (U.GetUpperBound(0) + 1) * (U.GetUpperBound(1) + 1)
  Public U_CddExc(0 To 80) As String            ' Comporte les candidats exclus 
  Public U_Sol(0 To 80) As String
  Public U_Clr_Cell_Fond(0 To 80) As Color      ' Couleur de fond de chaque cellule
  Public U_Clr_Cell_Val(0 To 80) As Color       ' Couleur de la valeur de chaque cellule
  'Ce tableau concerne les stratégies Cdd, CdU, CdO, Flt et Cbl à Unq
  'Il stocke les cellules concernées par la stratégie (Niveau de base Simple et Aide Graphique) pour n'effacer que celles-là
  Public U_Strg(0 To 80) As Boolean

  ' U_Strg_Val_Ins comporte pour chaque cellule LE SEUL CANDIDAT à INSéRER, c'est donc une zone de 1 caractère
  ' U_Strg_Cdd_Exc comporte pour chaque cellule le ou les candidats à exclure, c'est donc une zone de 9 caractères
  '                Soit la stratégie définie un candidat 
  '                Soit il y a plusieurs stratégies, donc plusieurs candidats
  '                Soit la stratégie (Unq) définie deux candidats
  Public U_Strg_Val_Ins(0 To 80) As String           'Initialisé à "" dans G4_Grid_Stratégie_All()
  '                                                   Documenté dans G4_Stratégie_CdO/CdU/Flt
  '                                                   Utilisé dans Mnu_Mngt
  Public U_Strg_Cdd_Exc(0 To 80) As String           'Initialisé à  Cnddts_Blancs dans G4_Grid_Stratégie_All()
  '                                                   Documenté dans Strategy_BTXYSJZKQ_Aide_Simple
  '                                                   Utilisé dans Mnu_Mngt

  Public U_Col(0 To 80) As Integer                   'Indique dans quelle Colonne se trouve une cellule (de 0 à 8)
  Public U_Row(0 To 80) As Integer                   '        dans quelle Ligne   se trouve une cellule (de 0 à 8)
  Public U_cr(0 To 80) As String                     '        les coordonnées Lx_Cy de chaque cellule
  Public U_Reg(0 To 80) As Integer                   '        dans quelle Région  se trouve une cellule (de 0 à 8)

  'Le terme Rectangle est remplacé par le terme Bande 
  Public U_Bh(0 To 80) As Integer                    '        N° de la bande horizontale
  Public U_Bv(0 To 80) As Integer                    '        N° de la bande verticale
  Public U_Bandes(0 To 5, 0 To 26) As Integer        ' les bandes 0,1,2 sont les bandes horizontales 1, 2 et 3 (0, 1, 2)
  '                                                  ' les bandes 3,4,5 sont les bandes verticales   1, 2 et 3 (0, 1, 2)
  Public Bande_Dcty As New Dictionary(Of String, Integer) _
    From {{"Bh1", 0}, {"Bh2", 1}, {"Bh3", 2}, {"Bv1", 0}, {"Bv2", 1}, {"Bv3", 2}}

  Public U_9CelCol(8)() As Integer
  Public U_9CelRow(8)() As Integer
  Public U_9CelReg(8)() As Integer
  Public U_20Cell_Coll(0 To 80)() As Integer

  Public Sqr_Img(0 To 80) As Image                  '  Le tableau comporte les 81 images du fond du jeu 
  Public U_Pt20(80, 19) As PointF                   '  Comporte 20 points pour chaque cellule   
  Public Sqr_Cel(0 To 80) As Rectangle              '  Le tableau comporte les informations des 81 rectangles du jeu
  Public Sqr_Cdd(809) As Rectangle                  '  Le tableau comporte les informations des 9 rectangles de chaque candidat de chaque cellule
  Public Grid_Pth As Drawing2D.GraphicsPath         '  Tracé de la Grille avec angle arrondi
  Public Reg_Pth(8) As Drawing2D.GraphicsPath       '  Tracés des 9 régions avec angle arrondi
  Public Sqr_Pth(0 To 80) As Drawing2D.GraphicsPath '  Tracés des squares du jeu avec ou sans angle arrondi ou biseauté
  ' Les traitements, compute ou paint, sont identiques que les paths soient des carrés, avec angles arrondis ou biseautés

  Public Structure MdC_Struct                       '  Structure des Modèles Composites d'une stratégie
    Public Cellule As Integer                       '  Cellule concernée
    Public MdE_Exist As Boolean                     '  Il y a au moins un modèle élémentaire
    Public MdE() As String                          '  Tableau des Modèles Elémentaires
  End Structure

  Public U_MdC(0 To 80) As MdC_Struct

#End Region

#Region "15 Undo-Redo"
  Public UR_A_Index As Integer = -1
  Public UR_Nb As Integer = 0
  Public UR_Nb_Annuler As Integer = 0
  Public UR_Nb_Refaire As Integer = 0
#End Region

#Region "20 Les Eléments graphiques: Color, ..."
  Public Event_OnPaint As String = "#"
  Public Event_OnPaint_MAP As String = ""
  Public Gz_Pt_TopLeft As Point
  'Structure des Traits
  'Il y a 10 Traits Horizontaux et 10 Traits Verticaux
  'Ces traits ont la même position ajoutée à Gz_Pt_TopLeft et la même longueur
  'Un trait dépasse d'une demi-épaisseur égale de chaque côté de la ligne et pas à son extrémité.
  Public Gz_Trait_Pos_xy(9) As Integer
  'Transparence
  'La transparence est associée au pixel A. 
  '   les couleurs Color_Couche_Stratégique et Color_Cell_Select sont transparente à 128 
  '   afin de laisser en-dessous les fonds et les chiffres (valeurs et candidats)
  '   les figures lors des stratégies doivent être également en transparence
  '   l'image de fond est également mise en tranparence pour ne pas trop masquer les chiffres et les candidats.
  ' Fond d'effacement
  Public Color_Frm_BackColor As Color = Color.FromArgb(255, 216, 245, 216) ' Couleur Fond du formulaire et du Grid
  Public Color_Trait As Color = Color.Green                                ' Couleur des traits du Grid
  Public Color_Fond_Typ_I As Color = Color.FromArgb(255, 192, 255, 192)    ' Couleur Fond Valeurs Initiales
  Public Color_Fond_Typ_RV As Color = Color.FromArgb(255, 129, 224, 129)   ' Couleur Fond Cellule Remplie/Vide
  Public Color_Stratégique As Color = Color.FromArgb(255, 15, 196, 101)

  Public Color_VI As Color = Color.FromArgb(255, 211, 40, 10)              ' Couleur d'affichage graphique des Valeurs Initiales       (Rouge)
  Public Color_VCdd As Color = Color.FromArgb(255, 23, 94, 23)             ' Couleur d'affichage graphique des Valeurs et des Candidats (Vert)
  'Public Color_Cell_Select As Color = Color.FromArgb(128, Color.White)
  Public Color_Cell_Select As Color = Color.FromArgb(128, 255, 255, 0)     ' A 128 R 255 G 255 B   0 
  Public Color_Cdd_Insérer As Color = Color.Yellow
  Public Color_Cdd_Exclure As Color = Color.Red
  Public Format_Center As New StringFormat With
          {
          .Alignment = StringAlignment.Center,
          .LineAlignment = StringAlignment.Center
          }

  'Polices non proportionnelles (pour le journal qui est RichTextBox)
  Public Font_Journal As New Font("Courier New", 10, FontStyle.Regular) 'Police non proportionnelle
  'Police utilisée pour le Frm_SDK.Mnu, et les menus contextuels du Journal et de l'Edition
  Public Font_Mnu As New Font("Segoe UI", 9, FontStyle.Regular)         'Police proportionnelle
  'Au démarrage, SDK utilise Arial et Segoe UI
  'Il faut ensuite gérer 2 polices:
  ' La police des valeurs et des candidats, et des Sqr_Fantasy
  Public Plcy_Fantasy As Boolean = False
  Public Plcy_Fantasy_Name As String = "Arial"

  ' Font_Mnu_Cel DOIT afficher le menu contextuel texte 
  Public Font_Mnu_Cel As New Font("Segoe UI", 10, FontStyle.Regular)
  Public Font_Name_ValCdd As String = "Arial"

  Public Font_Val_Size As Single       ' calcul effectué dans G0_Grid_Compute_Font_Size()
  Public Font_Cdd_Size As Single       ' calcul effectué dans G0_Grid_Compute_Font_Size() 
  Public Font_Cdd_Size_Zoom As Single  ' calcul effectué dans G0_Grid_Compute_Font_Size()
  Public Font_TTT_Txt As New Font("Arial", 15, FontStyle.Italic)

  'Certains caractères des polices ci-dessous ne sont pas explicites,
  ' ils sont donc remplacés
  ' Pour la police Wingding_2, le chiffres 7 est remplacés par F 
  ' La procédure Subst_Police effectue la substitution
  Public Subst_Arial_____ As String() = {"", "1", "2", "3", "4", "5", "6", "7", "8", "9"} ' Le standard b123456789
  Public Subst_MS_Outlook As String() = {"", "A", "B", "C", "M", "E", "F", "G", "I", "J"} ' glyphes homogènes
  Public Subst_Webdings__ As String() = {"", "A", "B", "F", "G", "I", "J", "P", "R", "S"} ' idem
  Public Subst_Wingding__ As String() = {"", "I", "2", "O", "4", "5", "3", "7", "8", "9"} ' idem
  ' Le tableau comporte ces 9 images
  ' La procédure Gz_Grid_Compute crée ces images
  Public Sqr_Fantasy(10) As Image
  ' La procédure Mnu_Mngt_Barre_Outils gère les images pour la Barre d'Outils
  '              Mnu_Mngt              gère les images pour le menu contextuel 
  Public InstalledFontCollection As New InstalledFontCollection()
  Public InstalledFontFamily() As FontFamily
  Public PrivateFontCollection As New PrivateFontCollection()
  Public PrivateFontFamily() As FontFamily

#End Region

#Region "30 Policies générales"
  '      Mode Nrm Normal 
  Public Plcy_Gnrl As String = "Nrm"
  'Policies de stratégie lorsque Plcy_Gnrl = "Nrm" 
  '      Peut prendre TOUTES les valeurs de Strg_Insertion_List, Strg_Exclusion_List, Plcy_Strg_FiltreVal_List
  '                   ou "   ""
  Public Plcy_Strg As String = "   "
  Public Prv_Plcy_Strg As String = "###"
  Public Plcy_Strg_Swt As Integer = -1
  Public Plcy_Stg_Clb As String            ' CkeckedListBox

  Public Plcy_Generate_Batch As Boolean = False
  Public Batch_en_Cours As Boolean = False
  Public Batch_Thread As Thread
  Public Plcy_Typ_I, Plcy_Typ_R, Plcy_Typ_V_sans_Cdd, Plcy_Typ_V_avec_Cdd As Boolean
  Public Plcy_Mnu_Item, Plcy_Mnu_Sep As Boolean

  'Policies particulières 
  Public Plcy_Saisir_Commencer As Boolean = False              ' Affichage de la grille de saisie avec les candidats éligibles
  Public Plcy_Solution_Existante As Boolean = False
  Public Plcy_AfficherDCdd_Bande As Boolean = True
  Public Plcy_Fond_Grille As Integer = 0                       ' (Numéro de la grille de fond)
  Public Plcy_MouseClick_Middle As Boolean = False
  Public Plcy_FIC_Frm_Insérer_Candidats As Boolean
  Public Plcy_FIC_TTT As String
  Public Plcy_FIC_Zone_Aimantée As String
  Public Plcy_Dancing_Link As Boolean = False
  Public Plcy_Open_Display As Boolean = False
  Public Const Dl_Nb_VI_minimal As Integer = 17                'https://fr.wikipedia.org/wiki/Math%C3%A9matiques_du_sudoku
  Public Plcy_Format_DAB As Integer
#End Region

#Region "45 Tableau des actions"
  'Ce tableau est documenté lors de Cell_Val_Insert     Ajouter / ? Ajouter      
  '                                 Cell_Val_Delete     Effacer
  '                                 Cell_Cdd_Insert     Replacer
  '                                 Cell_Cdd_Exclude    Exclure_Cdd
  Public Act(11, 100) As String
  Public Act_Index As Integer = -1
#End Region

#Region "70 Préférences"
  Public Create_Nb_Cel_Demandées As String
  Public Create_Nb_Tentatives As String
  Public Create_Contrainte_Originale As Integer
  Public Create_Chat As Boolean = True
  Public Lettre_Flèche As Integer
  Public Lettre_Flèche_ChrW As Integer
#End Region

#Region "80 Collection"
  'Défini à ce niveau:
  '(Recommandation : https://docs.microsoft.com/fr-fr/dotnet/api/system.random?view=net-6.0 )
  Public Rd0 As New Random
  Public Rd1 As New Random
  Public Rd2 As New Random
  Public Rd3 As New Random
  Public Rd4 As New Random
  Public Rd5 As New Random
  Public Rd6 As New Random
  Public Rd7 As New Random
  Public Rd8 As New Random
  Public Rd9 As New Random
  Public Rda As New Random
  Public Rdb As New Random
  Public Rdc As New Random
  Public Rde As New Random
  Public Rdf As New Random
  Public RdX As New Random
#End Region

#Region "90 Divers"
  Public T_Excel(80, 40) As String                           'Comporte les Cellules Excel
  Public Insertion_Exclusion_Nb_Erreurs As Integer = 0       'Comptage des insertions et exclusions non conformes
  Public Insert_Nb_Cell As Integer = 0
  Public Me_AutoScaleMode_Standard As AutoScaleMode = AutoScaleMode.Font ' Placés dans les Load de chaque formulaire.

#End Region

#Region "99 Dans le désordre"
  ''' <summary>Indication de la Cellule Clickée.</summary>
  Public Pbl_Cell_Select As Integer
  Public Pbl_Valeur_CdS As String
  Public Pbl_Cell_Candidat_Select As Integer
  ''' <summary>Indication de la Cellule Clickée Précédente.</summary>
  Public Prv_Pbl_Cell_Select As Integer
  Public Prv_Pbl_Cell_Candidat_Select As Integer
  Public Msg_Dcty As New Dictionary(Of String, String)()

  'uNuSeD
  Public Nsd_i As Integer
  Public Nsd_s As String
  Public Nsd_b As Boolean
  Public Nsd_CH As ColumnHeader
  Public Nsd_LV As ListViewItem.ListViewSubItem
  Public Nsd_P As Process
  'Préférences_Grille
  'Public Swt_Mode_Suggestion As Integer = -1                 '(Non)
  Public Swt_Mode_Dessin As Integer = -1                     '(Non)
  Public Swt_DéroulerJournal As Integer = 1                  '(Oui)
  Public Journal_Emp_Blocage As Integer = 0

  Public Swt_ModeEdition As Integer = -1
  Public Paint_Partie_Terminée_Nb As Integer = 0
  Public Msg_Dsp_MsgId As Boolean = False         ' Affiche ou non le MsgId devant le texte du message

  Public Game_Undo_Redo As String = String.Empty
  Public Game_Nb_Cellules_Initiales As Integer    ' Permet d'éviter le calcul systématique et répété de la zone B_Info
#End Region

#Region "Coloration"
  Public Class Color_Cls
    Public Symbol As String ' A, B, C, D
    Public Color As Color   ' Couleur correspondant au symbole et à la couleur choisie
  End Class
  '      Color_List est la List comportant les Couleurs en cours
  Public Color_List As New List(Of Color_Cls)
  '      Color_Originale_List est la list des "Couleurs d'usine"
  Public Color_Originale_List As New List(Of Color_Cls) From {
    New Color_Cls With {.Symbol = "A", .Color = Color.FromArgb(255, 255, 128, 0)},
    New Color_Cls With {.Symbol = "B", .Color = Color.FromArgb(255, 255, 192, 255)},
    New Color_Cls With {.Symbol = "C", .Color = Color.FromArgb(255, 0, 255, 255)},
    New Color_Cls With {.Symbol = "D", .Color = Color.FromArgb(255, 145, 32, 19)}}

  Public Class Objet_Cls
    Public Symbol As String ' A, B, C, D
    Public Forme As String  ' Cadre, Carré, Cercle, Disque, Croix, Flèche
    Public Cel_From As Integer
    Public Cdd_From As Integer
    Public Cel_To As Integer
    Public Cdd_To As Integer
    Public Point_From As PointF
    Public Point_To As PointF
  End Class
  Public Objet_List As New List(Of Objet_Cls)
  Public Pbl_PtF As PointF

  Public Structure Points_Struct
    Public Pt_From As PointF
    Public Pt_To As PointF
  End Structure

  Public Obj_Color As Color                                   ' Couleur choisie  
  Public Obj_Symbol As String                                 ' Symbole de la couleur choisie  
  Public Obj_Forme As String                                  ' Forme choisie  
  'Les paramètres My.settings sont:
  'My.Settings.Obj_Colors                                     ' Les 4 couleurs utilisées
  'My.Settings.Obj_Symbol                                     ' Dernier symbole de la couleur choisie
  'My.Settings.Obj_Forme                                      ' Dernière forme choisie  

  Public C1Items() As ToolStripMenuItem
  Public O1Items() As ToolStripMenuItem

  Public Sub InitMenuItems()
    With Frm_SDK
      .ContextMenuStrip = .Mnu_Obj
      C1Items = { .A, .B, .C, .D}
      O1Items = { .Cadre, .Carré, .Cercle, .Disque, .Croix, .Flèche}
    End With
  End Sub

  Public Flè_Cel_From As Integer
  Public Flè_Cdd_From As Integer
  Public Flè_From As Integer
  Public Flè_Pt_From As PointF
  Public Flè_Cel_To As Integer
  Public Flè_Cdd_To As Integer
  Public Flè_To As Integer
  Public Flè_Pt_To As PointF
#End Region

#Region "Structures"

  Public Structure Wh_Nb_Cell_Struct
    Public Initiales As Integer              ' Nombre de Cellules Initiales
    Public Vides As Integer                  ' Nombre de Cellules Vides
    Public Vides_sans_Candidats As Integer   ' Nombre de Cellules Vides sans candidats
    Public Remplies As Integer               ' Nombre de Cellules remplies
    Public Val_Nb() As Integer               ' Ventilation des Valeurs
    Public Cdd_Nb() As Integer               ' Ventilation des Candidats
  End Structure
#End Region

#Region "Classes_Lists"
  Public Class Pzzl_Cls 'Classe structurant les problèmes
    Public Property Nom As String           ' 0 Nom
    Public Property Force As String         ' 1 Force
    Public Property Problème As String      ' 2 Problème                             
    Public Property Solution As String      ' 3 Solution
    Public Property Nb_CI As Integer        ' 4 Nombre de Cellules Initiales
    Sub New(New_Nom As String,
            New_Force As String,
            New_Problème As String,
            New_Solution As String,
            New_Nb_CI As Integer)
      Nom = New_Nom
      Force = New_Force
      Problème = New_Problème
      Solution = New_Solution
      Nb_CI = New_Nb_CI
    End Sub
  End Class
  Public Pzzl_List As New List(Of Pzzl_Cls)  ' List comportant les problèmes

  Public Class Stg_Cls 'Classe structurant les stratégies
    Public Property Code As String         ' 0 Code
    Public Property Lettre As String       ' 1 Lettre ou N (pas de lettre) ou "L" (lien)
    Public Property Dsp_BO As String       ' 2 Afficher dans la barre d'outils ("O"/"N" ou "L" pour les stratégies de Liens)
    Public Property Type As String         ' 3 Type: Insertion, Exclusion, Non 
    Public Property Prd As String          ' 4 Prd: O/N, utilisé dans les calculs de Production/Résolution  
    Public Property Family As String       ' 5 0, de 1 à x pour gérer les affichages  
    Public Property Texte As String        ' 6 Texte                            
    'Constructeur paramétré
    Sub New(New_Code As String,
            New_Lettre As String,
            New_Dsp_BO As String,
            New_Type As String,
            New_Prd As String,
            New_Family As String,
            New_Texte As String)
      Code = New_Code
      Lettre = New_Lettre
      Dsp_BO = New_Dsp_BO
      Type = New_Type
      Prd = New_Prd
      Family = New_Family
      Texte = New_Texte
    End Sub
    ' Constructeur sans paramètres (IA)
    'Public Sub New()
    'End Sub
  End Class
  Public Stg_List As New List(Of Stg_Cls)        ' List comportant les stratégies
  Public Stg_List_Code As New List(Of String)    ' If Stg_List.Item(i).Prd = "O"
  Public Stg_List_Lettre As New List(Of String)  ' If Stg_List.Item(i).Prd = "O"
  Public Stg_List_Link As New List(Of String)    ' If Stg_List.Item(i).Lettre = "L"
  Public Stg_Bll() As Boolean                    ' UOBTXYSJZKQ_0123456789 (Prendra la taille lors de son utilisation)
  Public Stg_Profondeur As String

  Public Class Pzzl_Hodoku_Cls 'Classe structurant les problèmes tests Hodoku
    Public Property Nom As String           ' 0 Nom
    Public Property Jeu As String           ' 1 Est-ce un Commentaire ou un Jeu
    Public Property Problème As String      ' 2 Problème                             
    Sub New(New_Nom As String,
            New_Jeu As String,
            New_Problème As String)
      Nom = New_Nom
      Jeu = New_Jeu
      Problème = New_Problème
    End Sub
  End Class
  Public Pzzl_Hodoku_List As New List(Of Pzzl_Hodoku_Cls)  ' List comportant les problèmes Hodoku
#End Region

#Region "Stratégie Chain"
  ' Les noms:
  '     XLink, XLinks    un lien,   des liens
  '     XRoad, XRoads    un chemin, des chemins
  '     XRslt            le résultat de la structure
  '
  ' Les cellules sont initialisées à -1, les candidats à "0"
  ' Dim Variable() as Qqch est équivalent à Dim Variable as Qqch()
  '     Pour initier la variable: Il faut préférer Dim Variable as String() = {"0", "0", "0", "0", "0"} 

  Public Class XCdd_Cls ' Classe déterminant le nombre de candidats existants
    Public Property Cdd As String          ' Le candidat
    Public Property Nb As Integer          ' Le nombre de candidats
  End Class

  Public Class GCel_Cls ' Classe des cellules avec ses 9 candidats
    Public Property Cel As Integer = -1
    Public Property Cdd As String() = Enumerable.Repeat("0", 9).ToArray()
  End Class

  Public Class XCel_Excl_Cls 'Classe déterminant la structure des Cellules concernées Cdd_Excluses de la stratégie
    Public Property Cel As Integer = -1          ' La Cellule
    Public Property Cdd As String = "0"          ' Le Candidat
    Public Property Exc As Integer() = {-1, -1}  ' Les Cellules Origines de l'exclusion
  End Class

  Public Color_Link_S As Color = Color.Blue
  Public Color_Link_W As Color = Color.Green

  Public XCdds_List As New List(Of XCdd_Cls)
  Public XLinks_List As New List(Of XLink_Cls)         ' List des liens 

  Public Cell_Coll_Modifiées_List As New List(Of Integer)

  Public XAllRoads_List As New List(Of List(Of XLink_Cls))
  Public Const XRoads_Max As Integer = 645120          ' Nombre de chemins pour 7 liens
  '                                                      (2 ** n ) x n! 2 puissance n fois factoriel n
  Public GLinks As New List(Of GLink_Cls)              ' Liste des liens 
  Public GAllRoads As New List(Of List(Of GLink_Cls))  ' Liste de tous les chemins possibles

  Public GCels As New List(Of GCel_Cls)                ' Liste des cellules
  Public Const GRoads_Max As Integer = 645120
  Public Structure XRslt_Struct
    Public Code As String                         ' Code Stratégie: XCx / XCy /XRp /XNl
    Public Candidat() As String                   ' Le ou les candidats de la stratégie
    Public Cellule() As Integer                   ' La ou les cellules de la stratégie
    Public CelExcl As List(Of XCel_Excl_Cls)      ' Liste de Cellules concernées avec Cdd
    '                                             ' La liste comporte la cellule concernée, le candidat et les 2 cellules d'extrémités 
    Public Productivité As Boolean
    Public XRoads_Nombre As Integer
    Public XRoads_Numéro As Integer
    Public XLinks_Nombre As Integer
    Public RoadRight As List(Of XLink_Cls)        ' Liste des liens du Chemin correct
    Public Durée_ms As Integer                    ' Durée en ms du calcul de la stratégie
  End Structure
  Public XRslt As New XRslt_Struct
  Public Xap As Boolean = False
  Public XSolution As String
  Public U_Road(0 To 80) As Boolean               ' Tableau des Cellules sur le chemin

  Public Structure GRslt_Struct
    Public Code As String                         ' Code Stratégie 
    Public Candidat() As String                   ' Le ou les candidats de la stratégie
    Public Cellule() As Integer                   ' La ou les cellules de la stratégie

    Public Nb_Liens As Integer
    Public Nb_Noeuds As Integer
    Public Nb_Paths As Integer
    Public Path_Number As Integer

    Public RoadRight As List(Of GLink_Cls)        ' Liste des liens du Chemin correct
    Public Productivité As Boolean                ' If GRslt.CelExcl.Count > 0 Then GRslt.Productivité = True
    Public CelExcl As List(Of GCel_Excl_Cls)      ' Liste de Cellules concernées avec Cdd
    '                                             ' la liste comporte la cellule concernée, le candidat et les 2 cellules d'extrémités 
  End Structure
  Public GRslt As New GRslt_Struct
  Public Class GCel_Excl_Cls 'Classe déterminant la structure des Cellules concernées Cdd_Excluses de la stratégie
    Public Property Cel As Integer = -1          ' La Cellule
    Public Property Cdd As String = "0"          ' Le Candidat
    Public Property Exc As Integer() = {-1, -1}  ' Les Cellules Origines de l'exclusion
  End Class
  Public Structure RRslt_Struct
    Public Code_Strg As String                  ' Code Stratégie: CdU
    Public Code_Sous_Strg As String             ' Sous-Stratégie
    Public Code_LCR As String
    Public LCR As Integer
    Public Candidat As String                   ' Le candidat de la stratégie
    Public Candidats As String                   ' Les candidats de la stratégie (Unq)
    Public Cellule() As Integer                 ' Les cellules de la stratégie
    Public CelExcl() As Integer                 ' Les cellules concernées par l'exclusion
    Public Productivité As Boolean              ' False/ True
  End Structure
  Public RRslt As New RRslt_Struct

  Public Class DCdd_Cls 'Classe structurant la stratégie du Dernier Candidat
    Public Property Cellule As Integer
    Public Property Candidat As String
    Public Property Stratégie As String
    Public Property Sous_Stratégie As String
    'Constructeur
    Public Sub New(cellule As Integer, candidat As String, stratégie As String, sous_stratégie As String)
      Me.Cellule = cellule
      Me.Candidat = candidat
      Me.Stratégie = stratégie
      Me.Sous_Stratégie = sous_stratégie
    End Sub
  End Class
  Public DCdd_List As New List(Of DCdd_Cls)
  Public DCdd_Max As Integer = 5

  ' Liste aléatoire des cellules suivant algorithme de Fisher-Yates, Knuth
  Public Cell_FY_List As New List(Of Integer)
#End Region

End Module
