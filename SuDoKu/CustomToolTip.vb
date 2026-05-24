Public Class CustomToolTip
  Inherits ToolStripDropDown

  Private pnl As ToolStripControlHost
  Private lbl As Label

  Public Sub New(text As String, font As Font)
    AutoClose = False
    DoubleBuffered = True
    Margin = Padding.Empty
    Padding = Padding.Empty

    Dim isEmpty As Boolean = (U(Pbl_Cell_Select, 1) = " ")
    lbl = New Label()
    lbl.Text = text
    lbl.Font = font
    lbl.BackColor = If(isEmpty, Clr_Fnd_VI, Clr_Fnd_VCdd)
    lbl.ForeColor = If(isEmpty, Clr_VI, Clr_VCdd)
    lbl.AutoSize = True
    lbl.Padding = New Padding(4)

    pnl = New ToolStripControlHost(lbl)
    pnl.Margin = Padding.Empty
    pnl.Padding = Padding.Empty

    Items.Add(pnl)
  End Sub

  Public Sub ShowTooltip(location As Point)
    Show(location)
  End Sub
End Class