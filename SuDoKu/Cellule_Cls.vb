Option Strict On
Option Explicit On
Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices   ' Nécessaire à <DllImport("user32.dll")>
Imports System.Threading

Public Class Cellule_Cls
#Region "Propriétés"

  ' --- Listes statiques partagées par toutes les cellules ---
  Private Shared ReadOnly Format12 As New HashSet(Of Integer) From {0, 8, 72, 80}
  Private Shared ReadOnly Format34 As New HashSet(Of Integer) From {
        0, 2, 3, 5, 6, 8, 18, 26, 27, 35, 45, 53, 54, 62, 72, 74, 75, 77, 78, 80}
  Private Shared ReadOnly Format56_Extra As New HashSet(Of Integer) From {
        20, 21, 23, 24, 29, 30, 32, 33, 47, 48, 50, 51, 56, 57, 59, 60}

  ' --- Propriété ---
  Private _numéro As Integer
  Private _coté As Integer
  Private _candidat_unique As Boolean
  Private _typologie As String
  Private _couleur_fond As Color
  Private _couleur_valeur As Color
  Private _position As Point
  Private _position_Center As Point
  Private _text_tooltip As String
  Private _valeur As Integer                     ' Nouveau : INTEGER 0 si rien ou 1 à 9
  Private _valeur_initiale As Boolean            ' Nouveau : True ou False
  Private _candidats As String                   ' Nouveau : 123456789 ou 9blancs ou 1b3bb6b89
  Private ReadOnly _nombre_candidats As Integer  ' Nouveau : -1     la cellule est vide et il n'y a plus de candidat, ie #Erreur
  '                                                           0     la cellule est remplie
  '                                                           1 à 9 nombre de candidat
  '                                                           <>  Wh_Cell_Nb_Candidats() = 0  :  cellule vide sans candidat 
  Private _cellule_arrondie As Boolean

  ''' <summary>Numéro de la Cellule, compris entre 0 et 80, sinon 0.</summary>
  Public Property Numéro As Integer
    Get
      Dim num As Integer
      If _numéro < 0 Or _numéro > 80 Then
        num = 0
      Else
        num = _numéro
      End If
      Return num
    End Get
    Set(value As Integer)
      _numéro = value
    End Set
  End Property
  ''' <summary>Text_ToolTip de la cellule.</summary>
  Public Property Text_ToolTip As String
    Get
      Return _text_tooltip
    End Get
    Set(value As String)
      _text_tooltip = value
    End Set
  End Property
  ''' <summary>le Candidat est-il Unique ? </summary>
  Public ReadOnly Property Candidat_Unique As Boolean
    'Propriété dépendante de U
    Get
      _candidat_unique = False
      If (U(Numéro, 2) = " " And Trim(U(Numéro, 3)).Length = 1) Then _candidat_unique = True
      Return _candidat_unique
    End Get
  End Property
  ''' <summary> La cellule est-elle une valeur initiale.</summary>
  Public ReadOnly Property Valeur_Initiale As Boolean
    'Propriété dépendante de U
    Get
      _valeur_initiale = False
      If U(Numéro, 1) <> " " Then _valeur_initiale = True
      Return _valeur_initiale
    End Get
  End Property
  ''' <summary> Candidats de la cellule.</summary>
  Public ReadOnly Property Candidats As String
    'Propriété dépendante de U
    Get
      Dim blancs As String = StrDup(9, " ")
      _candidats = blancs
      If U(Numéro, 3) <> blancs Then _candidats = U(Numéro, 3)
      Return _candidats
    End Get
  End Property
  ''' <summary> Nombre de Candidats de la cellule.</summary>
  Public ReadOnly Property Nombre_Candidats As Integer
    'Propriété dépendante de U
    'Retourne le nombre de valeurs candidates pour une cellule (1 à 9) la cellule a x candidats
    '            0   si la cellule est remplie
    '            -1 (#Erreur) si la cellule est vide et n'a plus de candidats
    Get
      Dim n As Integer
      If U(Numéro, 2) <> " " Then n = 0
      If U(Numéro, 2) = " " And U(Numéro, 3) = Cnddts_Blancs Then n = -1
      If U(Numéro, 2) = " " And U(Numéro, 3) <> Cnddts_Blancs Then
        n = 0
        For i As Integer = 0 To 8
          If U(Numéro, 3).Substring(i, 1) <> " " Then n += 1
        Next i
      End If
      Return n
    End Get
  End Property
  ''' <summary> Valeur de la cellule </summary>
  Public ReadOnly Property Valeur As Integer
    'Propriété dépendante de U
    Get
      _valeur = 0
      If U(Numéro, 2) <> " " Then _valeur = CInt(U(Numéro, 2))
      Return _valeur
    End Get
  End Property
  ''' <summary>Taille de la Cellule, définie dans Préférence/Grille/Taille.</summary>
  Public ReadOnly Property Coté As Integer
    Get
      _coté = WH
      Return _coté
    End Get
  End Property
  ''' <summary>Position de la Cellule.</summary>
  Public ReadOnly Property Position As Point
    Get
      _position.X = Sqr_Cel(Numéro).X
      _position.Y = Sqr_Cel(Numéro).Y
      Return _position
    End Get
  End Property
  ''' <summary>Position Centrale de la Cellule.</summary>
  Public ReadOnly Property Position_Center As Point
    Get
      _position_Center.X = Sqr_Cel(Numéro).X + WHhalf
      _position_Center.Y = Sqr_Cel(Numéro).Y + WHhalf
      Return _position_Center
    End Get
  End Property
  ''' <summary>Typologie de la Cellule.</summary>
  ''' <returns><c>I (Valeur Initiale),R (Cellule Remplie), ou V (Cellule Vide) </c> .</returns>
  Public ReadOnly Property Typologie As String
    Get 'Propriété dépendante de U
      _typologie = "X"
      If U(Numéro, 1) <> " " Then _typologie = "I"
      If U(Numéro, 1) = " " And U(Numéro, 2) <> " " Then _typologie = "R"
      'La Typologie V n'implique pas que la cellules ait des candidats.
      'En mode Sas, une cellule vide n'a de candidats que s'ils ont été insérés.
      If U(Numéro, 1) = " " And U(Numéro, 2) = " " Then _typologie = "V"
      Return _typologie
    End Get
  End Property
  ''' <summary>Couleur de Fond de la Cellule.</summary>
  Public ReadOnly Property Couleur_Fond As Color
    Get 'Propriété dépendante de U
      _couleur_fond = Color_Frm_BackColor
      If U(Numéro, 1) <> " " Then _couleur_fond = Color_Fond_Typ_I
      If U(Numéro, 1) = " " Then _couleur_fond = Color_Fond_Typ_RV
      Return _couleur_fond
    End Get
  End Property
  ''' <summary>Couleur de la valeur de la Cellule.</summary>
  Public ReadOnly Property Couleur_Valeur As Color
    Get 'Propriété dépendante de U
      Select Case Typologie
        Case "I" : Return Color_VI
        Case "R" : Return Color_VCdd
        Case Else : Return Color.Transparent
      End Select
    End Get
  End Property
  ''' <summary>Précise si la cellule est arrondie ou non.</summary>
  Public ReadOnly Property Cellule_Arrondie As Boolean
    ' Définition des cellules arrondies selon les formats DAB
    Get 'Propriété dépendante de U et de Plcy_Format_DAB
      Select Case Plcy_Format_DAB
        Case 1, 2 : Return Format12.Contains(Numéro)
        Case 3, 4 : Return Format34.Contains(Numéro)
        Case 5, 6 : Return Format34.Contains(Numéro) OrElse Format56_Extra.Contains(Numéro)
        Case Else
          Return False
      End Select
    End Get
  End Property
#End Region

#Region "Méthodes"
  ' Définition
  ' Cell_Cls   désigne une cellule, il y en a 81
  ' Grid       désigne les 81 cellules
  ' Pzzl       désigne une grille correcte, c'est à dire un Puzzle SuDoKu
  '

  ''' <summary>Peint le Fond de la Cellule .</summary>
  Public Sub G2_Cellule_Paint_Fond()
    'Concerne le fond d'une cellule quelque soit sa Typologie : Initiale, Remplie ou Vide ou une image
    'Plcy_Fond_Grille représente le n° de fond choisi dans la liste des fonds d'image
    '                 0 est le "Fond Standard", ie une couleur et non une photo

    Dim g As Graphics = Frm_SDK.CreateGraphics
    Using brsh_0 As New SolidBrush(U_Clr_Cell_Fond(Numéro)),
          brsh As New SolidBrush(Color_Frm_BackColor)


      If Plcy_Fond_Grille = 0 Then
        ' Un fond standard est affiché
        If Cellule_Arrondie Then
          g.FillPath(brsh_0, Sqr_Pth(Numéro))
        Else
          g.FillRectangle(brsh_0, Sqr_Cel(Numéro))
        End If
      Else
        ' L'image de fond est affichée
        If Cellule_Arrondie Then
          g.ResetClip()
          g.SetClip(Sqr_Pth(Numéro), CombineMode.Replace)
          g.DrawImage(Sqr_Img(Numéro), Sqr_Cel(Numéro).X, Sqr_Cel(Numéro).Y)
        Else
          g.FillRectangle(brsh, Sqr_Cel(Numéro))
          g.DrawImage(Sqr_Img(Numéro), Sqr_Cel(Numéro).X, Sqr_Cel(Numéro).Y)
        End If
      End If
    End Using
    g.Dispose()

    'Traite les Cas particuliers et la propriété Text_ToolTip
    If Plcy_Gnrl = "Nrm" And Plcy_Strg = "   " Then
      Select Case Typologie
        Case "I"
        Case "R"
          ' 1 Indication d'une valeur non conforme à la solution (la solution existe et il y a une erreur)
          If Plcy_Solution_Existante And U(Numéro, 2) <> U_Sol(Numéro) Then
            G0_Cell_Figure(Numéro, "Disque", Color.FromArgb(128, Color.Green))
            Text_ToolTip = "Valeur non conforme à la solution."
          End If
        Case "V"
          ' 2 Indication d'une cellule sans candidat
          If Nombre_Candidats = -1 Then
            G0_Cell_Figure(Numéro, "Disque", Color.FromArgb(128, Color.Red))
            Text_ToolTip = "La cellule n'a plus de candidat."
          End If
          ' 3 Mode Suggestion
          If Swt_Mode_Suggestion = 1 And U_Suggest(Numéro) <> "0" Then
            G0_Cell_Figure(Numéro, "Disque", Color.FromArgb(64, Color.Yellow))
            Text_ToolTip = "... Cellule à jouer."
          End If
      End Select
    End If
  End Sub

  ''' <summary>Peint la valeur d'une cellule IR.</summary>
  Public Sub G5_Cellule_Paint_Valeur()
    'Concerne l'ensemble des Cellules Initiales et Remplies
    'Les valeurs sont peintes dans une couleur différentes suivant leur typologie I/R
    Dim g As Graphics = Frm_SDK.CreateGraphics
    g.DrawString(Subst_Police(U(Numéro, 2)),
                 New Font(Font_Name_ValCdd, Font_Val_Size, FontStyle.Regular),
                 New SolidBrush(Couleur_Valeur),
                 Position_Center.X, Position_Center.Y, Format_Center)
    g.Dispose()
  End Sub
  ''' <summary>Dessine UN Candidat de la Cellule.</summary>
  Public Sub G6_Cellule_Paint_Candidat(Candidat As String, Couleur As Color)
    'Dessine UN Candidat d'une cellule dans un cercle de couleur
    Dim Coté_6 As Integer = (WH \ 6)
    If Not Candidats.Contains(Candidat) Then Exit Sub
    Dim Cdd_n As Integer = (Numéro * 10) + CInt(Candidat)
    Dim Sqr_Cdd_n As Rectangle = Sqr_Cdd(Cdd_n)
    Sqr_Cdd_n.Inflate(-1, -1)    'Diminution du cercle du candidat  
    Using g As Graphics = Frm_SDK.CreateGraphics
      g.FillPie(New SolidBrush(Color.FromArgb(128, Couleur)), Sqr_Cdd_n, 0.0F, 360.0F)
      g.DrawString(Subst_Police(Candidat),
                       New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Regular),
                       New SolidBrush(Color_VCdd),
                       Sqr_Cdd(Cdd_n).X + Coté_6, Sqr_Cdd(Cdd_n).Y + Coté_6, Format_Center)
    End Using
  End Sub
  ''' <summary>Dessine les Candidats de la Cellule.</summary>
  Public Sub G6_Cellule_Paint_Candidats(ByVal typeCdd As String)
    'Procédure utilisée pour dessiner le fond de sélection d'une cellule
    If Typologie = "I" Or Typologie = "R" Then Exit Sub
    Dim Coté_6 As Integer = Coté \ 6
    Dim cdd_n As Integer
    Using g As Graphics = Frm_SDK.CreateGraphics
      For cdd As Integer = 1 To 9
        If (typeCdd = "Les9Candidats") _
        Or (typeCdd = "LesCandidatsEligibles" And Candidats.Contains(cdd.ToString())) Then
          cdd_n = (Numéro * 10) + cdd
          g.DrawString(Subst_Police(CStr(cdd)),
                       New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Regular),
                       New SolidBrush(Color_VCdd),
                       Sqr_Cdd(cdd_n).X + Coté_6, Sqr_Cdd(cdd_n).Y + Coté_6, Format_Center)
        End If
      Next cdd
    End Using
  End Sub
  ''' <summary>Dessine les Candidats de la Cellule sous conditions habituelles.</summary>
  Public Sub G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
    '10 occurences
    'Concerne UNIQUEMENT les cellules Vides avec des Candidats et dans les conditions précisées
    If (Plcy_Gnrl = "Nrm" And Plcy_Strg = "Cdd") _
    Or (Plcy_Gnrl = "Edi") _
    Or (Plcy_Gnrl = "Sas") Then
      G6_Cellule_Paint_Candidats("LesCandidatsEligibles")
    End If
  End Sub
  ''' <summary>Peint la sélection de la Cellule.</summary>
  Public Sub G7_Cellule_Paint_Select()
    ' C'est l'affichage des candidats qui constitue l'affichage de la sélection
    Dim g As Graphics = Frm_SDK.CreateGraphics
    If Cellule_Arrondie Then
      g.FillPath(New SolidBrush(Color_Cell_Select), Sqr_Pth(Numéro))
    Else
      g.FillRectangle(New SolidBrush(Color_Cell_Select), Sqr_Cel(Numéro))
    End If
    g.Dispose()

    If Typologie = "V" Then
      If (Plcy_Gnrl = "Sas" And U(Numéro, 3) = Cnddts_Blancs) _
      Or (Plcy_Gnrl = "Nrm" And Plcy_Strg = "   ") _
      Or (Plcy_Gnrl = "Nrm" And Stg_Get(Plcy_Strg).Type = "I" And Not Plcy_AideGraphique) _
      Or (Plcy_Gnrl = "Nrm" And Stg_Get(Plcy_Strg).Type = "E" And Not Plcy_AideGraphique) _
      Or (Plcy_Gnrl = "Nrm" And Mid$(Plcy_Strg, 1, 2) = "FV" And Not Plcy_AideGraphique) Then
        Select Case Plcy_Saisir_Commencer
          Case True : G6_Cellule_Paint_Candidats("LesCandidatsEligibles")
          Case False : G6_Cellule_Paint_Candidats("Les9Candidats")
        End Select
      End If

      If (Plcy_Gnrl = "Nrm" And Stg_Get(Plcy_Strg).Type = "I" And Plcy_AideGraphique) _
      Or (Plcy_Gnrl = "Nrm" And Stg_Get(Plcy_Strg).Type = "E" And Plcy_AideGraphique) _
      Or (Plcy_Gnrl = "Nrm" And Mid$(Plcy_Strg, 1, 2) = "FV" And Plcy_AideGraphique) Then
        G6_Cellule_Paint_Candidats("LesCandidatsEligibles")
      End If
    End If

    Select Case Plcy_Gnrl
      Case "Nrm", "Sas" : Mnu_Mngt(Numéro)
      Case Else
    End Select

    Select Case Plcy_Gbl_Etendue
      Case True : Frm_SDK.B_Position.Text = U_Coord(Numéro) & " (" & Numéro & ")"
      Case False : Frm_SDK.B_Position.Text = U_Coord(Numéro)
    End Select
  End Sub
  Public Sub Cellule_Refresh()
    'La Cellule est rafraîchie, la Cellule n'est pas sélectionnée 
    G2_Cellule_Paint_Fond()
    G4_Grid_Stratégie_All()
    G5_Cellule_Paint_Valeur()
    G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
  End Sub
  ''' <summary>Rafraîchit la Cellule et les Cellules Collatérales</summary>
  Public Sub Cellule_Refresh_Cell_Coll()
    'La cellule et les cellules collatérales sont rafraîchies, la Cellule n'est pas sélectionnée
    'Traitement identique à Cellule_Refresh pour la Cellule Originelle 
    'Il y a 20 cellules collatérales pour une Cellule
    'A1 Début de Traitement de la Cellule Originale
    G2_Cellule_Paint_Fond()
    'B  Traitement des Cellules Collatérales
    Dim Grp() As Integer = U_20Cell_Coll(Numéro)
    Dim sc_Grp As New Cellule_Cls
    For g As Integer = 0 To UBound(Grp)
      sc_Grp.Numéro = Grp(g)
      sc_Grp.G2_Cellule_Paint_Fond()
      sc_Grp.G5_Cellule_Paint_Valeur()
      sc_Grp.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
    Next g
    'A2 Fin de Traitement de la Cellule Originale
    G4_Grid_Stratégie_All()
    G5_Cellule_Paint_Valeur()
    G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
  End Sub
#End Region
End Class

Public Class Grille_Cls
#Region "Propriétés"
  Private _nb_cellules_initiales As Integer            ' Nouveau  
  Private _nb_cellules_remplies As Integer             ' Nouveau  
  Private _nb_cellules_vides As Integer                ' Nouveau  
  Public ReadOnly Property Nb_Cellules_Initiales As Integer
    Get
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        If sc.Valeur_Initiale Then _nb_cellules_initiales += 1
      Next i
      Return _nb_cellules_initiales
    End Get
  End Property
  Public ReadOnly Property Nb_Cellules_Remplies As Integer
    Get
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        If sc.Valeur <> 0 Then _nb_cellules_remplies += 1
      Next i
      Return _nb_cellules_remplies
    End Get
  End Property
  Public ReadOnly Property Nb_Cellules_Vides As Integer
    Get
      Dim sc As New Cellule_Cls
      For i As Integer = 0 To 80
        sc.Numéro = i
        If sc.Valeur = 0 Then _nb_cellules_vides += 1
      Next i
      Return _nb_cellules_vides
    End Get
  End Property
#End Region

#Region "Méthodes"
  Public Sub G2_Grille_Paint_Fond()
    Dim sc As New Cellule_Cls
    For i As Integer = 0 To 80
      sc.Numéro = i
      sc.G2_Cellule_Paint_Fond()
    Next i
  End Sub

  ''' <summary>Compose la couche indirecte.</summary>
  Public Sub G3_Grille_Paint_Indirecte()
    'Les couches G2 et G3 sont à traiter ensemble pour la couche indirecte
    ' G2_Grille_Paint_Fond = 81 G2_Cellule_Paint_Fond
    ' G3_Grille_Paint_Indirecte
    ' Les 2 Fonctions reprennent les mêmes conditions pour les effets indirects.

    Dim U_G3(80) As Boolean ' Un tableau booléen est créé FALSE
    'A Analyse des causes et des effets indirects
    For i As Integer = 0 To 80
      ' 1 Indication d'une valeur non conforme à la solution (la solution existe et il y a une erreur)
      ' 2 Indication d'une cellule sans candidat
      If Plcy_Gnrl = "Nrm" And Plcy_Strg = "   " And Wh_Cell_Nb_Candidats(U, i) = 0 Then U_G3(i) = True
      ' 3 Indication Dernier Candidat d'une Unitée en mode Nrm ou Sas
      ' 4 Mode Suggestion
      If Plcy_Gnrl = "Nrm" And Plcy_Strg = "   " And Swt_Mode_Suggestion = 1 And U_Suggest(i) <> "0" Then U_G3(i) = True
    Next i

    'B Répercussion effets indirects dans la grille
    '  La fonction Cellule_Refresh effectue G2_Cellule_Paint_Fond()
    Dim sc As New Cellule_Cls
    For i As Integer = 0 To 80
      If U_G3(i) = True Then
        sc.Numéro = i
        sc.Cellule_Refresh()
      End If
    Next i
  End Sub

  Public Sub Grille_Refresh()
    'La grille est rafraîchie entièrement, aucune cellule n'est sélectée
    G1_Grid_Paint()
    Dim Gril As New Grille_Cls
    Gril.G2_Grille_Paint_Fond()
    G4_Grid_Stratégie_All()
    Dim sc As New Cellule_Cls
    For i As Integer = 0 To 80
      sc.Numéro = i
      sc.G5_Cellule_Paint_Valeur()
      sc.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
    Next i
  End Sub

  Public Sub Grille_Refresh_g(g As Graphics)
    'La grille est rafraîchie entièrement, aucune cellule n'est sélectée
    Jrn_Add_Yellow(Procédure_Name_Get())
    G1_Grid_Paint_g(g)
    Dim Gril As New Grille_Cls
    Gril.G2_Grille_Paint_Fond()
    G4_Grid_Stratégie_All()
    Dim sc As New Cellule_Cls
    For i As Integer = 0 To 80
      sc.Numéro = i
      sc.G5_Cellule_Paint_Valeur()
      sc.G6_Cellule_Paint_Candidats_Conditions_Sas_Nrm_Cdd()
    Next i
  End Sub

  Public Sub G8_Grille_Partie_Terminée()
    Dim Cellule_Clct As New Collection
    ' Ne se fait que si Plcy_Gnrl = "Nrm" ou "Sas"
    If Plcy_Gnrl <> "Nrm" And Plcy_Gnrl <> "Sas" Then Exit Sub
    ' Il faut que les 81 cellules soient remplies et que la grille soit correcte
    ' Il faut que la partie ait été jouée (Act_Index > 1)

    Dim U_Chk(80, 3) As String
    Array.Copy(U, U_Chk, UNbCopy)
    Dim U_Check As U_Check_Struct = U_Checking(U_Chk)
    'La grille est  remplie et elle est correcte
    If (Nb_Cellules_Remplies = 81 And U_Check.Check = True And Nb_Cellules_Initiales < 81) Then
      'Permet de ne pas afficher le message en mode 2-3 de génération
      If Act_Index = 0 Then Exit Sub
      If Paint_Partie_Terminée_Nb > 2 Then Exit Sub
      Cursor.Current = Cursors.WaitCursor
      Paint_Partie_Terminée_Nb += 1
      Select Case Plcy_Gnrl
        Case "Nrm"
          Select Case Insertion_Exclusion_Nb_Erreurs
            Case 0 : Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_50030")
            Case 1 : Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_50031")
            Case Else : Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_50032", {CStr(Insertion_Exclusion_Nb_Erreurs)})
          End Select
        Case "Sas"
          Frm_SDK.B_Info.Text = Msg_Read_IA("SDK_50029")
      End Select

      Dim sc As New Cellule_Cls
      'Collection des valeurs initiales
      For i As Integer = 0 To 80
        sc.Numéro = i
        If sc.Valeur_Initiale Then Clct_Add(Cellule_Clct, sc.Numéro)
      Next i
      ' Animation de la Grille
      Dim Rect As Rectangle
      Dim Inflate As Integer = 0
      Using g As Graphics = Frm_SDK.CreateGraphics
        For i As Integer = 1 To Cellule_Clct.Count
          sc.Numéro = Clct_Random(Cellule_Clct)
          Rect = Sqr_Cel(sc.Numéro)
          Inflate += 2        '2
          If Inflate > WH \ 2 Then Exit For
          Rect.Inflate(New Size(Inflate - 2, Inflate - 2))
          g.DrawIcon(My.Resources.SuDoKu, Rect)
          Thread.Sleep(100)
          Rect.Inflate(New Size(Inflate, Inflate))
          g.DrawIcon(My.Resources.SuDoKu, Rect)
          Thread.Sleep(100)
          Rect.Inflate(New Size(Inflate + 1, Inflate + 1))
          g.DrawIcon(My.Resources.SuDoKu, Rect)
          Thread.Sleep(100)
        Next i
      End Using
      Cursor.Current = Cursors.Default
    End If
  End Sub
#End Region
End Class

Public Class ProgressBarGraphic_Cls
  '------------------------------------------------------------------------------------------
  'Date de création: Samedi 10/09/2022
  'Mise en place d'un barre de progression graphique qui adopte la couleur du trait du thème
  'Elle est positionnée en lieu et place de Frm_SDK.B_Info
  '------------------------------------------------------------------------------------------
#Region "Propriétés"
  Private _position As Point
  Private _position_left As Integer
  Private _position_top As Integer
  Private _hauteur As Integer
  Private _largeur As Integer
  Private _couleur As Color
  ''' <summary>Position de la Progress Bar Graphic.</summary>
  Public ReadOnly Property Position As Point
    Get
      _position.X = Frm_SDK.B_Info.Location.X
      _position.Y = Frm_SDK.B_Info.Location.Y
      Return _position
    End Get
  End Property
  ''' <summary>Position Left de la Progress Bar Graphic.</summary>
  Public ReadOnly Property Position_Left As Integer
    Get
      _position_left = Frm_SDK.B_Info.Location.X
      Return _position_left
    End Get
  End Property
  ''' <summary>Position_Top de la Progress Bar Graphic.</summary>
  Public ReadOnly Property Position_Top As Integer
    Get
      _position_top = Frm_SDK.B_Info.Location.Y
      Return _position_top
    End Get
  End Property
  ''' <summary>Hauteur de la Progress Bar Graphic.</summary>
  Public ReadOnly Property Hauteur As Integer
    Get
      _hauteur = Frm_SDK.B_Info.Height
      Return _hauteur
    End Get
  End Property
  ''' <summary>Largeur de la Progress Bar Graphic.</summary>
  Public ReadOnly Property Largeur As Integer
    Get
      _largeur = Frm_SDK.B_Info.Width
      Return _largeur
    End Get
  End Property
  ''' <summary>Couleur de la Progress Bar Graphic.</summary>
  Public ReadOnly Property Couleur As Color
    Get
      _couleur = Color_Trait
      Return _couleur
    End Get
  End Property
#End Region

#Region "Méthodes"
  ''' <summary>Peint la Progress Bar Graphic.</summary>
  Public Sub Draw_ProgressBarGraphic(Prc As Integer)
    ' Sécurisation de la valeur
    Prc = Math.Max(0, Math.Min(100, Prc))
    Dim PBG_Rct As New Rectangle(x:=Position_Left, y:=Position_Top,
                                 width:=(Largeur * Prc) \ 100,
                                 height:=Hauteur)
    Dim PBG_Rct_Txt As New Rectangle(x:=Position_Left, y:=Position_Top,
                                     width:=Largeur,
                                     height:=Hauteur)
    Using g As Graphics = Frm_SDK.CreateGraphics
      g.FillRectangle(New SolidBrush(Color_Frm_BackColor), PBG_Rct_Txt) ' Rafraîchissement
      g.FillRectangle(New SolidBrush(Couleur), PBG_Rct)
      g.DrawString(CStr(Prc) & " %",
                   New Font(Font_Name_ValCdd, 10, FontStyle.Regular),
                   New SolidBrush(Color.Red),
                   Position_Left + Largeur \ 2, Position_Top + Hauteur \ 2, Format_Center)
    End Using
  End Sub
#End Region

End Class

'https://www.bing.com/search?q=ColorDialog%20title%20modification%20Visual%20Studio%20Visual%20Basic&qs=ds&form=ATCVAJ
Public Class SDK_ColorDialog
  'Vendredi 05/04/2024 Personnalisation du Titre de la Boîte des couleurs
  Inherits ColorDialog

  <DllImport("user32.dll")>
  Private Shared Function SetWindowText(hWnd As IntPtr, lpString As String) As Boolean
  End Function

  Private _title As String = String.Empty
  Private _titleSet As Boolean = False

  Public Property Title As String
    Get
      Return _title
    End Get
    Set(value As String)
      If value IsNot Nothing AndAlso value <> _title Then
        _title = value
        _titleSet = False
      End If
    End Set
  End Property

  Protected Overrides Function HookProc(hWnd As IntPtr, msg As Integer, wparam As IntPtr, lparam As IntPtr) As IntPtr
    If Not _titleSet Then
      SetWindowText(hWnd, _title)
      _titleSet = True
    End If
    Return MyBase.HookProc(hWnd, msg, wparam, lparam)
  End Function
End Class

Friend Module Cell_Grid

  ''' <summary>Retourne False/True en fonction du numéro de la cellule</summary>
  Public Function Cell_Numéro_Valide(ByRef Cellule As Integer) As Boolean
    Cell_Numéro_Valide = False
    If Cellule >= 0 And Cellule <= 80 Then
      Cell_Numéro_Valide = True
    End If
    Return Cell_Numéro_Valide
  End Function

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