' Usage : Clic Middle pour afficher les candidats de l'unité
Public Class CustomToolTip
  Inherits Form
  Private _text As String
  Private _font As Font

  Public Sub New(text As String, font As Font)
    _text = text
    _font = font
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
    Using sf As New StringFormat With {.Alignment = StringAlignment.Center}
      e.Graphics.DrawString(_text, _font, Brushes.Black, New RectangleF(0, 0, Me.ClientSize.Width, Me.ClientSize.Height), sf)
    End Using
  End Sub
  Public Sub ShowTooltip(location As Point)
    Me.Location = location
    Show()
  End Sub
  Public Sub HideTooltip()
    Hide()
  End Sub
End Class