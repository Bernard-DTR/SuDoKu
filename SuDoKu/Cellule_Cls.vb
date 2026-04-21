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
  Private _isvalid As Boolean                    ' Nouveau : True compris entre 0 et 80, sinon False
  Private _coté As Integer
  Private _candidat_unique As Boolean
  Private _typologie As String
  Private _position As Point
  Private _position_Center As Point
  Private _valeur As Integer                     ' Nouveau : INTEGER 0 si rien ou 1 à 9
  Private _valeur_initiale As Boolean            ' Nouveau : True ou False
  Private _candidats As String                   ' Nouveau : 123456789 ou 9blancs ou 1b3bb6b89
  Private ReadOnly _cellule_arrondie As Boolean

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
  ''' <summary>Le numéro de la cellule est-il valide ? </summary>
  Public ReadOnly Property IsValid As Boolean
    Get
      If Numéro < 0 OrElse Numéro > 80 Then Return False
      Return True
    End Get
  End Property

  ''' <summary>le Candidat est-il Unique ? </summary>
  Public ReadOnly Property Candidat_Unique As Boolean
    'Propriété dépendante de U
    Get
      _candidat_unique = False
      If (U(Numéro, 2) = " " AndAlso Trim(U(Numéro, 3)).Length = 1) Then _candidat_unique = True
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
      If U(Numéro, 1) = " " And U(Numéro, 2) = " " Then _typologie = "V"
      Return _typologie
    End Get
  End Property
  ''' <summary>Précise si la cellule est arrondie ou non.</summary>
  Public ReadOnly Property Cellule_Arrondie As Boolean
    ' Définition des cellules arrondies selon les formats DAB
    Get 'Propriété dépendante de U et de Plcy_Format_DAB
      Select Case Plcy_Format_DAB
        Case 1 : Return Format34.Contains(Numéro) OrElse Format56_Extra.Contains(Numéro)
        Case Else
          Return False
      End Select
    End Get
  End Property
#End Region

#Region "Méthodes"
  ''' <summary>Peint le Fond de la Cellule .</summary>
  Public Sub G2_Cellule_Paint_Fond(g As Graphics)
    'TODO Le fond de la cellule doit traiter 3 choses:
    '             La valeur saisie est différente de la solution
    '             la cellule est vide et n'a plus de candidats
    '             l'affichage de la cellule vise présente un quadrillage de saisie et/ou des chiffres
    'Concerne le fond d'une cellule quelque soit sa Typologie : Initiale, Remplie ou Vide ou une image
    'Plcy_Fond_Grille représente le n° de fond choisi dans la liste des fonds d'image
    '                 0 est le "Fond Standard", ie une couleur et non une photo
    If Not IsValid Then Exit Sub
    Using brsh_0 As New SolidBrush(U_Clr_Cell_Fond(Numéro)),
          brsh As New SolidBrush(Color_Frm_BackColor)
      If Plcy_Fond_Grille = 0 Then    ' Un fond standard est affiché
        If Cellule_Arrondie Then
          g.FillPath(brsh_0, Sqr_Pth(Numéro))
        Else
          g.FillRectangle(brsh_0, Sqr_Cel(Numéro))
        End If
      Else                            ' L'image de fond est affichée
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
  End Sub

  ''' <summary>Peint la valeur d'une cellule IR.</summary>
  Public Sub G5_Cellule_Paint_Valeur(g As Graphics)
    'Concerne l'ensemble des Cellules Initiales et Remplies
    'Les valeurs sont peintes dans une couleur différentes suivant leur typologie I/R
    If Not IsValid Then Exit Sub
    Using brsh As New SolidBrush(U_Clr_Cell_Val(Numéro)),
          fnt As New Font(Font_Name_ValCdd, Font_Val_Size, FontStyle.Regular)
      g.DrawString(Subst_Police(U(Numéro, 2)),
                   fnt,
                   brsh,
                   Position_Center.X, Position_Center.Y, Format_Center)
    End Using
  End Sub

  ''' <summary>Dessine UN Candidat de la Cellule.</summary>
  Public Sub G6_Cellule_Paint_Candidat(g As Graphics, Candidat As String, Couleur As Color)
    'Dessine UN Candidat d'une cellule dans un cercle de couleur
    'Un candidat a toujours la même couleur, puisqu'il ne peut être affiché que dans la typologie V
    Dim Coté_6 As Integer = (WH \ 6)
    If Not IsValid Then Exit Sub
    If Not Candidats.Contains(Candidat) Then Exit Sub
    Dim Cdd_n As Integer = (Numéro * 10) + CInt(Candidat)
    Dim Sqr_Cdd_n As Rectangle = Sqr_Cdd(Cdd_n)
    Sqr_Cdd_n.Inflate(-1, -1)    'Diminution du cercle du candidat  
    Using brsh_1 As New SolidBrush(Color.FromArgb(128, Couleur)),
          brsh_2 As New SolidBrush(Color_VCdd),
          font As New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Regular)
      g.FillPie(brsh_1, Sqr_Cdd_n, 0.0F, 360.0F)
      g.DrawString(Subst_Police(Candidat),
                   font,
                   brsh_2,
                   Sqr_Cdd(Cdd_n).X + Coté_6, Sqr_Cdd(Cdd_n).Y + Coté_6, Format_Center)
    End Using
  End Sub

  ''' <summary>Dessine les Candidats de la Cellule.</summary>
  Public Sub G6_Cellule_Paint_Candidats(g As Graphics, ByVal typeCdd As String)
    'Procédure utilisée pour dessiner le fond de sélection d'une cellule
    If Not IsValid Then Exit Sub
    If Typologie = "I" Or Typologie = "R" Then Exit Sub
    Dim Coté_6 As Integer = Coté \ 6
    Dim cdd_n As Integer
    Using font As New Font(Font_Name_ValCdd, Font_Cdd_Size, FontStyle.Regular),
          brsh As New SolidBrush(Color_VCdd)
      For cdd As Integer = 1 To 9
        If (typeCdd = "Les9Candidats") _
        Or (typeCdd = "LesCandidatsEligibles" And Candidats.Contains(cdd.ToString())) Then
          cdd_n = (Numéro * 10) + cdd
          g.DrawString(Subst_Police(CStr(cdd)),
                       font,
                       brsh,
                       Sqr_Cdd(cdd_n).X + Coté_6, Sqr_Cdd(cdd_n).Y + Coté_6, Format_Center)
        End If
      Next cdd
    End Using
  End Sub
#End Region
End Class


Public Class SDK_ColorDialog
  '05/04/2024 Personnalisation du Titre de la Boîte des couleurs
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