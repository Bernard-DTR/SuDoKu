'  Created By: Marc Cramer
'  Legal Copyright: Marc Cramer © 1/28/2003

Option Strict On
Option Explicit On

Imports System.ComponentModel

Public Class ColorCombobox
  Inherits UserControl
  '=====================================================================================
  ' VARIABLE DECLARATION
  '=====================================================================================
  ReadOnly Cbb_Color_K() As String = Split("Black,Gray,Silver,Brown,Maroon,Red,Salmon,Tomato,Coral,Sienna,Chocolate,Peru,Orange,Gold,Khaki,Olive,Yellow,Chartreuse,Green,Lime,Turquoise,Teal,Aqua,Cyan,Navy,Blue,Indigo,Thistle,Plum,Violet,Purple,Fuchsia,Magenta,Orchid", ",")
  ReadOnly Cbb_Color_K_french() As String = Split("Noir,Gris,Argent,Marron,Bordeaux,Rouge,Saumon,Tomate,Corail,Terre de Sienne,Chocolat,Pérou,Orange,Or,Kaki,Olive,Jaune,Chartreuse,Vert,Citron vert,Turquoise,Sarcelle,Aqua,Cyan,Marine,Bleu,Indigo,Chardon,Prune,Violet,Violet,Fuchsia,Magenta,Orchidée", ",")

  Public Class Cbb_Color_Cls 'Classe structurant les couleurs
    Public Property Color_en As String           ' Couleur anglais
    Public Property Color_fr As String           ' Couleur français
    Sub New(New_Color_en As String,
            New_Color_fr As String)
      Color_en = New_Color_en
      Color_fr = New_Color_fr
    End Sub
  End Class
  Public Cbb_Color_List As New List(Of Cbb_Color_Cls)  ' List comportant les couleurs en et fr
  '=====================================================================================
  ' PROPERTY VARIABLES
  '=====================================================================================
  Dim m_ColorSelected As String
  Dim m_ColorExclude As String
  '=====================================================================================
  ' WINDOWS FORM DESIGNER GENERATED CODE
  '=====================================================================================
  Public Sub New()
    MyBase.New()
    'This call is required by the Windows Form Designer.
    Cbb_Color_InitializeComponent()
  End Sub
  'UserControl overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(disposing As Boolean)
    If disposing Then
      components?.Dispose()
      'If Not (components Is Nothing) Then
      '    components.Dispose()
      'End If
    End If
    MyBase.Dispose(disposing)
  End Sub
  'Required by the Windows Form Designer
  Private ReadOnly components As IContainer
  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents Cbb_Color As ComboBox
  <DebuggerStepThrough()>
  Private Sub Cbb_Color_InitializeComponent()
    Me.Cbb_Color = New ComboBox()
    Me.SuspendLayout()
    '
    'Cbb_Color
    '
    Me.Cbb_Color.BackColor = Color.White
    Me.Cbb_Color.Dock = System.Windows.Forms.DockStyle.Fill
    Me.Cbb_Color.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed
    Me.Cbb_Color.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.Cbb_Color.Location = New Point(0, 0)
    Me.Cbb_Color.Name = "Cbb_Color"
    Me.Cbb_Color.Size = New Size(150, 21)
    Me.Cbb_Color.TabIndex = 0
    '
    'ColorCombobox
    '
    Me.Controls.Add(Me.Cbb_Color)
    Me.Name = "ColorCombobox"
    Me.Size = New Size(150, 24)
    Me.ResumeLayout(False)

  End Sub
  '=====================================================================================
  'EVENTS
  '=====================================================================================
  Public Event Cbb_Color_Changed_Selected(ColorSelected As Color, sender As Object)
  Public Event Cbb_Color_Index_Changed_Selected(sender As Object, e As EventArgs)
  Private Sub Cbb_Color_DropDown(sender As Object, e As EventArgs) Handles Cbb_Color.DropDown
    Dim sColor As String
    sColor = m_ColorSelected
    Cbb_Color_Include_Load()
    Cbb_Color.SelectedIndex = Cbb_Color.Items.IndexOf(sColor)
  End Sub
  Private Sub Cbb_Color_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Cbb_Color.SelectedIndexChanged
    ' a new value was picked so set variable and fire event
    m_ColorSelected = Cbb_Color.Text
    RaiseEvent Cbb_Color_Changed_Selected(Color.FromName(Cbb_Color.Text), sender)
    RaiseEvent Cbb_Color_Index_Changed_Selected(sender, e)
  End Sub
  '=====================================================================================
  ' PROPERTIES
  '=====================================================================================
  <Description("Gets and sets the currently selected color."), Category("Police_Display")>
  Public Property Cbb_Color_Selected() As String
    Get
      Return m_ColorSelected
    End Get
    Set(Value As String)
      m_ColorSelected = Value
      Dim ColorIndex As Integer
      ColorIndex = Cbb_Color.FindStringExact(m_ColorSelected)
      Cbb_Color.SelectedIndex = ColorIndex
    End Set
  End Property
  '=====================================================================================
  Public Property Cbb_Color_Exclude() As String
    Get
      Return m_ColorExclude
    End Get
    Set(Value As String)
      m_ColorExclude = Value
    End Set
  End Property
  '=====================================================================================
  ' METHODS
  '=====================================================================================
  Private Sub Cbb_Color_DrawItem(sender As System.Object, e As DrawItemEventArgs) Handles Cbb_Color.DrawItem
    ' if a valid value then draw the color entry
    If e.Index <> -1 Then
      Cbb_Color_Item(e.Graphics, e.Bounds, e.Index)
    End If
  End Sub

  Private Sub Cbb_Color_Include_Load()
    ' load our color entry list into the combobox
    Cbb_Color.Items.Clear()
    'Prise en compte de Cbb_Color_Exclude
    For i As Integer = 0 To Cbb_Color_List.Count - 1
      If Cbb_Color_List.Item(i).Color_en = Cbb_Color_Exclude Then
        'La couleur de VI   ne sera pas proposée pour VCdd
        'La couleur de VCdd ne sera pas proposée pour VI
      Else
        Nsd_i = Cbb_Color.Items.Add(Cbb_Color_List.Item(i).Color_en)
      End If
    Next i
  End Sub
  Private Sub Cbb_Color_Item(g As Graphics, ItemRectangle As Rectangle, ItemIndex As Integer)

    Dim itemText As String = Cbb_Color.Items.Item(ItemIndex).ToString()
    Dim colorName As String = Cbb_Color_Name_fr(itemText)

    Using brushText As New SolidBrush(Color.FromKnownColor(KnownColor.MenuText)),
          brushColor As New SolidBrush(Color.FromName(itemText)),
          penBorder As New Pen(Color.Black, 1)

      g.FillRectangle(brushColor,
                    ItemRectangle.Left + 2,
                    ItemRectangle.Top + 2,
                    20,
                    ItemRectangle.Height - 4)

      g.DrawRectangle(penBorder,
                           New Rectangle(ItemRectangle.Left + 1,
                                         ItemRectangle.Top + 1,
                                         21,
                                         ItemRectangle.Height - 3))

      g.DrawString(colorName,
                      Cbb_Color.Font,
                      brushText,
                      ItemRectangle.Left + 28,
                      ItemRectangle.Top + ((ItemRectangle.Height - Cbb_Color.Font.GetHeight()) / 2))

    End Using

  End Sub
  Public Function Cbb_Color_Name_fr(Clr_Name_en As String) As String
    'Retourne le nom fr de la couleur en
    Cbb_Color_Name_fr = "#"
    For i As Integer = 0 To Cbb_Color_List.Count - 1
      If Cbb_Color_List.Item(i).Color_en = Clr_Name_en Then
        Cbb_Color_Name_fr = Cbb_Color_List.Item(i).Color_fr
        Exit For
      End If
    Next i
    Return Cbb_Color_Name_fr
  End Function
  Private Sub Cbb_Color_Load(sender As Object, e As EventArgs) Handles Me.Load
    'Affiche la couleur lors du premier affichage
    'Remplissage de la List de couleur avec les noms en et fr
    'Il ne semble pas nécessaire de clearer la collection
    Cbb_Color_List.Clear()
    For Counter As Integer = 0 To Cbb_Color_K.GetUpperBound(0)
      Cbb_Color_List.Add(New Cbb_Color_Cls(Cbb_Color_K(Counter), Cbb_Color_K_french(Counter)))
    Next Counter
    Cbb_Color_Include_Load()
  End Sub
End Class