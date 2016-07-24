<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConsultaCheque
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
        Me.gridCheque = New System.Windows.Forms.DataGridView()
        Me.Cheque = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Banco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Importe = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.FechaComprobante = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fechaCheque = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.comprobante = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.observaciones = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.btnCerrar = New WindowsApplication1.CustomButton()
        Me.btnProceso = New WindowsApplication1.CustomButton()
        Me.cmbTipo = New System.Windows.Forms.ComboBox()
        Me.txtCheque = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel1 = New WindowsApplication1.CustomLabel()
        CType(Me.gridCheque, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gridCheque
        '
        Me.gridCheque.AllowUserToAddRows = False
        Me.gridCheque.AllowUserToDeleteRows = False
        Me.gridCheque.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader
        Me.gridCheque.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridCheque.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Cheque, Me.Banco, Me.Importe, Me.FechaComprobante, Me.fechaCheque, Me.comprobante, Me.observaciones})
        Me.gridCheque.Location = New System.Drawing.Point(12, 53)
        Me.gridCheque.Name = "gridCheque"
        Me.gridCheque.Size = New System.Drawing.Size(760, 497)
        Me.gridCheque.StandardTab = True
        Me.gridCheque.TabIndex = 2
        '
        'Cheque
        '
        Me.Cheque.HeaderText = "Cheque"
        Me.Cheque.Name = "Cheque"
        Me.Cheque.ReadOnly = True
        Me.Cheque.Width = 69
        '
        'Banco
        '
        Me.Banco.HeaderText = "Banco"
        Me.Banco.Name = "Banco"
        Me.Banco.ReadOnly = True
        Me.Banco.Width = 63
        '
        'Importe
        '
        Me.Importe.HeaderText = "Importe"
        Me.Importe.Name = "Importe"
        Me.Importe.ReadOnly = True
        Me.Importe.Width = 67
        '
        'FechaComprobante
        '
        Me.FechaComprobante.HeaderText = "FechaComprobante"
        Me.FechaComprobante.Name = "FechaComprobante"
        Me.FechaComprobante.ReadOnly = True
        Me.FechaComprobante.Width = 125
        '
        'fechaCheque
        '
        Me.fechaCheque.HeaderText = "Fecha Cheque"
        Me.fechaCheque.Name = "fechaCheque"
        Me.fechaCheque.ReadOnly = True
        Me.fechaCheque.Width = 102
        '
        'comprobante
        '
        Me.comprobante.HeaderText = "Comprobante"
        Me.comprobante.Name = "comprobante"
        Me.comprobante.ReadOnly = True
        Me.comprobante.Width = 95
        '
        'observaciones
        '
        Me.observaciones.HeaderText = "Observaciones"
        Me.observaciones.Name = "observaciones"
        Me.observaciones.ReadOnly = True
        Me.observaciones.Width = 103
        '
        'btnCerrar
        '
        Me.btnCerrar.Cleanable = False
        Me.btnCerrar.EnterIndex = -1
        Me.btnCerrar.LabelAssociationKey = -1
        Me.btnCerrar.Location = New System.Drawing.Point(276, 9)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(84, 23)
        Me.btnCerrar.TabIndex = 3
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnProceso
        '
        Me.btnProceso.Cleanable = False
        Me.btnProceso.EnterIndex = -1
        Me.btnProceso.LabelAssociationKey = -1
        Me.btnProceso.Location = New System.Drawing.Point(366, 9)
        Me.btnProceso.Name = "btnProceso"
        Me.btnProceso.Size = New System.Drawing.Size(84, 23)
        Me.btnProceso.TabIndex = 4
        Me.btnProceso.Text = "Proceso"
        Me.btnProceso.UseVisualStyleBackColor = True
        '
        'cmbTipo
        '
        Me.cmbTipo.FormattingEnabled = True
        Me.cmbTipo.Items.AddRange(New Object() {"Cheque Terceros", "Cheques Propios"})
        Me.cmbTipo.Location = New System.Drawing.Point(517, 12)
        Me.cmbTipo.Name = "cmbTipo"
        Me.cmbTipo.Size = New System.Drawing.Size(255, 21)
        Me.cmbTipo.TabIndex = 5
        '
        'txtCheque
        '
        Me.txtCheque.Cleanable = False
        Me.txtCheque.Empty = True
        Me.txtCheque.EnterIndex = -1
        Me.txtCheque.LabelAssociationKey = -1
        Me.txtCheque.Location = New System.Drawing.Point(79, 12)
        Me.txtCheque.MaxLength = 8
        Me.txtCheque.Name = "txtCheque"
        Me.txtCheque.Size = New System.Drawing.Size(121, 20)
        Me.txtCheque.TabIndex = 6
        Me.txtCheque.Validator = WindowsApplication1.ValidatorType.None
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Location = New System.Drawing.Point(12, 19)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(44, 13)
        Me.CustomLabel1.TabIndex = 7
        Me.CustomLabel1.Text = "Cheque"
        '
        'ConsultaCheque
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Controls.Add(Me.txtCheque)
        Me.Controls.Add(Me.cmbTipo)
        Me.Controls.Add(Me.btnProceso)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.gridCheque)
        Me.Name = "ConsultaCheque"
        Me.Text = "ConsultaCheque"
        CType(Me.gridCheque, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents gridCheque As System.Windows.Forms.DataGridView
    Friend WithEvents btnCerrar As WindowsApplication1.CustomButton
    Friend WithEvents btnProceso As WindowsApplication1.CustomButton
    Friend WithEvents cmbTipo As System.Windows.Forms.ComboBox
    Friend WithEvents txtCheque As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents Cheque As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Banco As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Importe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents FechaComprobante As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents fechaCheque As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents comprobante As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents observaciones As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
