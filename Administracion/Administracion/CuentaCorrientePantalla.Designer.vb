<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CuentaCorrientePantalla
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
        Me.GRilla = New System.Windows.Forms.DataGridView()
        Me.Column = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustomLabel3 = New WindowsApplication1.CustomLabel()
        Me.Proveedor = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel1 = New WindowsApplication1.CustomLabel()
        Me.ProveedorRazon = New WindowsApplication1.CustomTextBox()
        CType(Me.GRilla, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GRilla
        '
        Me.GRilla.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GRilla.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Column})
        Me.GRilla.Location = New System.Drawing.Point(93, 99)
        Me.GRilla.Name = "GRilla"
        Me.GRilla.Size = New System.Drawing.Size(497, 289)
        Me.GRilla.StandardTab = True
        Me.GRilla.TabIndex = 1
        '
        'Column
        '
        Me.Column.HeaderText = "Hola"
        Me.Column.Name = "Column"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = -1
        Me.CustomLabel3.Location = New System.Drawing.Point(90, 39)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel3.TabIndex = 3
        Me.CustomLabel3.Text = "Proveedor"
        '
        'Proveedor
        '
        Me.Proveedor.Cleanable = False
        Me.Proveedor.Empty = True
        Me.Proveedor.EnterIndex = -1
        Me.Proveedor.LabelAssociationKey = -1
        Me.Proveedor.Location = New System.Drawing.Point(182, 36)
        Me.Proveedor.Name = "Proveedor"
        Me.Proveedor.Size = New System.Drawing.Size(108, 20)
        Me.Proveedor.TabIndex = 0
        Me.Proveedor.Validator = WindowsApplication1.ValidatorType.None
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Location = New System.Drawing.Point(330, 39)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(0, 13)
        Me.CustomLabel1.TabIndex = 5
        '
        'ProveedorRazon
        '
        Me.ProveedorRazon.BackColor = System.Drawing.Color.Silver
        Me.ProveedorRazon.Cleanable = False
        Me.ProveedorRazon.Empty = True
        Me.ProveedorRazon.EnterIndex = -1
        Me.ProveedorRazon.LabelAssociationKey = -1
        Me.ProveedorRazon.Location = New System.Drawing.Point(307, 36)
        Me.ProveedorRazon.Name = "ProveedorRazon"
        Me.ProveedorRazon.Size = New System.Drawing.Size(358, 20)
        Me.ProveedorRazon.TabIndex = 6
        Me.ProveedorRazon.Validator = WindowsApplication1.ValidatorType.None
        '
        'CuentaCorrientePantalla
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 434)
        Me.Controls.Add(Me.ProveedorRazon)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Controls.Add(Me.Proveedor)
        Me.Controls.Add(Me.CustomLabel3)
        Me.Controls.Add(Me.GRilla)
        Me.Name = "CuentaCorrientePantalla"
        Me.Text = "Form2"
        CType(Me.GRilla, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GRilla As System.Windows.Forms.DataGridView
    Friend WithEvents Column As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustomLabel3 As WindowsApplication1.CustomLabel
    Friend WithEvents Proveedor As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents ProveedorRazon As WindowsApplication1.CustomTextBox
End Class
