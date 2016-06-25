<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CargaIntereses
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.gridCtaCte = New System.Windows.Forms.DataGridView()
        CType(Me.gridCtaCte, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gridCtaCte
        '
        Me.gridCtaCte.AllowUserToAddRows = False
        Me.gridCtaCte.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridCtaCte.Location = New System.Drawing.Point(12, 12)
        Me.gridCtaCte.Name = "gridCtaCte"
        Me.gridCtaCte.Size = New System.Drawing.Size(760, 350)
        Me.gridCtaCte.StandardTab = True
        Me.gridCtaCte.TabIndex = 1
        '
        'CargaIntereses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.gridCtaCte)
        Me.Name = "CargaIntereses"
        Me.Text = "Carga Intereses"
        CType(Me.gridCtaCte, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gridCtaCte As System.Windows.Forms.DataGridView
End Class
