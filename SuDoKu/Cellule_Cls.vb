Imports System.Runtime.InteropServices   ' Nécessaire à <DllImport("user32.dll")>

Public Class Cellule_Cls
#Region "Propriétés"

  ' --- Propriété ---
  Private _numéro As Integer
  Private ReadOnly _isvalid As Boolean           ' Nouveau : True compris entre 0 et 80, sinon False
  Private _typologie As String
  Private _position As Point
  Private _position_Center As Point
  Private _valeur As Integer                     ' Nouveau : INTEGER 0 si rien ou 1 à 9
  Private _candidats As String                   ' Nouveau : 123456789 ou 9blancs ou 1b3bb6b89

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
  ''' <summary> Valeur de la cellule </summary>
  Public ReadOnly Property Valeur As Integer
    'Propriété dépendante de U
    Get
      _valeur = 0
      If U(Numéro, 2) <> " " Then _valeur = CInt(U(Numéro, 2))
      Return _valeur
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
      _position_Center.X = Sqr_Cel(Numéro).X + WHh
      _position_Center.Y = Sqr_Cel(Numéro).Y + WHh
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
#End Region

#Region "Méthodes"
  ''' <summary>Dessine UN Candidat de la Cellule.</summary>
  Public Sub G6_Cellule_Paint_Candidat_Eligible(g As Graphics, Candidat As String, Couleur As Color)
    'Dessine UN Candidat d'une cellule dans un cercle de couleur
    'Un candidat a toujours la même couleur, puisqu'il ne peut être affiché que dans la typologie V
    If Not IsValid Then Exit Sub
    If Not Candidats.Contains(Candidat) Then Exit Sub
    Dim Cdd_n As Integer = (Numéro * 10) + CInt(Candidat)
    Dim Sqr_Cdd_n As Rectangle = Sqr_Cdd_Inf(Cdd_n)
    Using brsh_1 As New SolidBrush(Color.FromArgb(128, Couleur))
      g.FillPie(brsh_1, Sqr_Cdd_n, 0.0F, 360.0F)
      g.DrawString(Subst_Police(Candidat),
                   Fnt_Cdd,
                   Brh_VCdd,
                   Sqr_Cdd(Cdd_n), Format_Center)
    End Using
  End Sub

  ''' <summary>Dessine les Candidats de la Cellule.</summary>
  Public Sub G6_Cellule_Paint_Candidats_Eligibles(g As Graphics)
    'Procédure utilisée pour dessiner le fond de sélection d'une cellule
    If Not IsValid Then Exit Sub
    If Typologie = "I" Or Typologie = "R" Then Exit Sub
    Dim cdd_n As Integer
    For cdd As Integer = 1 To 9
      If Candidats.Contains(cdd.ToString()) Then
        cdd_n = (Numéro * 10) + cdd
        g.DrawString(Subst_Police(CStr(cdd)),
                     Fnt_Cdd,
                     Brh_VCdd,
                     Sqr_Cdd(cdd_n), Format_Center)
      End If
    Next cdd
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
    ' Cette fonction modifie le titre de la boite de dialogue des couleurs 
    ' dans Préférences / Couleurs
    If Not _titleSet Then
      SetWindowText(hWnd, _title)
      _titleSet = True
    End If
    Return MyBase.HookProc(hWnd, msg, wparam, lparam)
  End Function
End Class