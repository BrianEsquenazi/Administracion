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
        Me.btnCancela = New WindowsApplication1.CustomButton()
        Me.btnGraba = New WindowsApplication1.CustomButton()
        CType(Me.gridCtaCte, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gridCtaCte
        '
        Me.gridCtaCte.AllowUserToAddRows = False
        Me.gridCtaCte.AllowUserToDeleteRows = False
        Me.gridCtaCte.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader
        Me.gridCtaCte.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridCtaCte.Location = New System.Drawing.Point(12, 12)
        Me.gridCtaCte.Name = "gridCtaCte"
        Me.gridCtaCte.Size = New System.Drawing.Size(760, 468)
        Me.gridCtaCte.StandardTab = True
        Me.gridCtaCte.TabIndex = 1
        '
        'btnCancela
        '
        Me.btnCancela.Cleanable = False
        Me.btnCancela.EnterIndex = -1
        Me.btnCancela.LabelAssociationKey = -1
        Me.btnCancela.Location = New System.Drawing.Point(385, 496)
        Me.btnCancela.Name = "btnCancela"
        Me.btnCancela.Size = New System.Drawing.Size(122, 42)
        Me.btnCancela.TabIndex = 3
        Me.btnCancela.Text = "Cancela"
        Me.btnCancela.UseVisualStyleBackColor = True
        '
        'btnGraba
        '
        Me.btnGraba.Cleanable = False
        Me.btnGraba.EnterIndex = -1
        Me.btnGraba.LabelAssociationKey = -1
        Me.btnGraba.Location = New System.Drawing.Point(248, 496)
        Me.btnGraba.Name = "btnGraba"
        Me.btnGraba.Size = New System.Drawing.Size(122, 42)
        Me.btnGraba.TabIndex = 2
        Me.btnGraba.Text = "Graba"
        Me.btnGraba.UseVisualStyleBackColor = True
        '
        'CargaIntereses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.btnCancela)
        Me.Controls.Add(Me.btnGraba)
        Me.Controls.Add(Me.gridCtaCte)
        Me.Name = "CargaIntereses"
        Me.Text = "Actualizacion de Deuda de Pyme Nacion"
        CType(Me.gridCtaCte, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents gridCtaCte As System.Windows.Forms.DataGridView
    Friend WithEvents btnGraba As WindowsApplication1.CustomButton
    Friend WithEvents btnCancela As WindowsApplication1.CustomButton
End Class
