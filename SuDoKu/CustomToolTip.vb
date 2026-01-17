Option Strict On
Option Explicit On

' Création le 12/05/2025
Public Class CustomToolTip
  Inherits Form

  Private _text As String
  Private _font As Font

  Public Sub New(text As String, font As Font)
    _text = text
    _font = font

    ' Configurer la fenêtre
    FormBorderStyle = FormBorderStyle.None
    ShowInTaskbar = False
    StartPosition = FormStartPosition.Manual
    BackColor = Color_Frm_BackColor
    Opacity = 0.9
    Padding = New Padding(4)

    ' Calculer la taille de la fenêtre en fonction du texte et de la police
    Dim textSize As SizeF = TextRenderer.MeasureText(_text, _font)
    ClientSize = New Size(CInt(textSize.Width) + 8, CInt(textSize.Height) + 8)
  End Sub

  Protected Overrides Sub OnPaint(e As PaintEventArgs)
    MyBase.OnPaint(e)
    ' Dessiner le texte avec la police spécifiée
    'e.Graphics.DrawString(_text, _font, Brushes.Black, New PointF(4, 4))
    'Le TTT est cadré à droite
    'Dim sf As New StringFormat With {.Alignment = StringAlignment.Far}
    Dim sf As New StringFormat With {.Alignment = StringAlignment.Center}
    e.Graphics.DrawString(_text, _font, Brushes.Black, New RectangleF(0, 0, Me.ClientSize.Width, Me.ClientSize.Height), sf)
  End Sub

  Public Sub ShowTooltip(location As Point)
    Me.Location = location
    Show()
  End Sub

  Public Sub HideTooltip()
    Hide()
  End Sub
End Class