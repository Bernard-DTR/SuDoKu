<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Frm_MySettings
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.PG = New System.Windows.Forms.PropertyGrid()
        Me.SuspendLayout()
        '
        'PG
        '
        Me.PG.Location = New System.Drawing.Point(22, 12)
        Me.PG.Name = "PG"
        Me.PG.PropertySort = System.Windows.Forms.PropertySort.NoSort
        Me.PG.Size = New System.Drawing.Size(755, 313)
        Me.PG.TabIndex = 0
        '
        'Frm_MySettings
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 344)
        Me.Controls.Add(Me.PG)
        Me.Name = "Frm_MySettings"
        Me.Text = "Frm_MySettings"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents PG As PropertyGrid
End Class
