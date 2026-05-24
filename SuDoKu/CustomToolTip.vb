Public Class CustomToolTip
  Inherits ToolStripDropDown

  Private panel As ToolStripControlHost
  Private lbl As Label

  Public Sub New(text As String, font As Font)
    AutoClose = False
    DoubleBuffered = True
    Margin = Padding.Empty
    Padding = Padding.Empty

    lbl = New Label()
    lbl.Text = text
    lbl.Font = font
    lbl.BackColor = Color.Yellow
    lbl.AutoSize = True
    lbl.Padding = New Padding(4)

    panel = New ToolStripControlHost(lbl)
    panel.Margin = Padding.Empty
    panel.Padding = Padding.Empty

    Items.Add(panel)
  End Sub

  Public Sub ShowTooltip(location As Point)
    Show(location)
  End Sub
End Class
