<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RecibosProvisorios
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
        Me.gridRecibos = New System.Windows.Forms.DataGridView()
        Me.Tipo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.numero = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.banco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.importe = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustomLabel10 = New Administracion.CustomLabel()
        Me.lstSeleccion = New Administracion.CustomListBox()
        Me.txtConsulta = New Administracion.CustomTextBox()
        Me.lstConsulta = New Administracion.CustomListBox()
        Me.lblTotal = New Administracion.CustomLabel()
        Me.btnLimpiar = New Administracion.CustomButton()
        Me.btnCerrar = New Administracion.CustomButton()
        Me.btnIntereses = New Administracion.CustomButton()
        Me.btnConsulta = New Administracion.CustomButton()
        Me.btnAgregar = New Administracion.CustomButton()
        Me.txtTotal = New Administracion.CustomTextBox()
        Me.CustomLabel8 = New Administracion.CustomLabel()
        Me.CustomLabel9 = New Administracion.CustomLabel()
        Me.txtParidad = New Administracion.CustomTextBox()
        Me.txtRetSuss = New Administracion.CustomTextBox()
        Me.CustomLabel6 = New Administracion.CustomLabel()
        Me.CustomLabel7 = New Administracion.CustomLabel()
        Me.txtRetIB = New Administracion.CustomTextBox()
        Me.txtRetIva = New Administracion.CustomTextBox()
        Me.CustomLabel5 = New Administracion.CustomLabel()
        Me.CustomLabel4 = New Administracion.CustomLabel()
        Me.txtRetGanancias = New Administracion.CustomTextBox()
        Me.txtNombre = New Administracion.CustomTextBox()
        Me.txtCliente = New Administracion.CustomTextBox()
        Me.txtRecibo = New Administracion.CustomTextBox()
        Me.txtFecha = New Administracion.CustomTextBox()
        Me.CustomLabel3 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.CustomLabel1 = New Administracion.CustomLabel()
        CType(Me.gridRecibos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gridRecibos
        '
        Me.gridRecibos.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.gridRecibos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridRecibos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tipo, Me.numero, Me.fecha, Me.banco, Me.importe})
        Me.gridRecibos.Location = New System.Drawing.Point(22, 174)
        Me.gridRecibos.Name = "gridRecibos"
        Me.gridRecibos.Size = New System.Drawing.Size(756, 363)
        Me.gridRecibos.TabIndex = 13
        '
        'Tipo
        '
        Me.Tipo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.Tipo.FillWeight = 80.0!
        Me.Tipo.HeaderText = "Tipo"
        Me.Tipo.Name = "Tipo"
        Me.Tipo.Width = 53
        '
        'numero
        '
        Me.numero.FillWeight = 120.0!
        Me.numero.HeaderText = "Numero/Cta"
        Me.numero.Name = "numero"
        '
        'fecha
        '
        Me.fecha.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.fecha.FillWeight = 120.0!
        Me.fecha.HeaderText = "Fecha"
        Me.fecha.Name = "fecha"
        Me.fecha.Width = 62
        '
        'banco
        '
        Me.banco.FillWeight = 150.0!
        Me.banco.HeaderText = "Banco"
        Me.banco.Name = "banco"
        '
        'importe
        '
        Me.importe.FillWeight = 80.0!
        Me.importe.HeaderText = "Importe"
        Me.importe.Name = "importe"
        '
        'CustomLabel10
        '
        Me.CustomLabel10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.CustomLabel10.ControlAssociationKey = -1
        Me.CustomLabel10.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomLabel10.Location = New System.Drawing.Point(379, 540)
        Me.CustomLabel10.Name = "CustomLabel10"
        Me.CustomLabel10.Size = New System.Drawing.Size(247, 22)
        Me.CustomLabel10.TabIndex = 118
        Me.CustomLabel10.Text = "Tipo de Doc.:   1) Efectivo   2) Cheque"
        Me.CustomLabel10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lstSeleccion
        '
        Me.lstSeleccion.Cleanable = False
        Me.lstSeleccion.EnterIndex = -1
        Me.lstSeleccion.FormattingEnabled = True
        Me.lstSeleccion.LabelAssociationKey = -1
        Me.lstSeleccion.Location = New System.Drawing.Point(492, 5)
        Me.lstSeleccion.Name = "lstSeleccion"
        Me.lstSeleccion.Size = New System.Drawing.Size(286, 134)
        Me.lstSeleccion.TabIndex = 78
        Me.lstSeleccion.Visible = False
        '
        'txtConsulta
        '
        Me.txtConsulta.Cleanable = False
        Me.txtConsulta.Empty = True
        Me.txtConsulta.EnterIndex = -1
        Me.txtConsulta.LabelAssociationKey = -1
        Me.txtConsulta.Location = New System.Drawing.Point(492, 6)
        Me.txtConsulta.Name = "txtConsulta"
        Me.txtConsulta.Size = New System.Drawing.Size(286, 20)
        Me.txtConsulta.TabIndex = 77
        Me.txtConsulta.Validator = Administracion.ValidatorType.None
        Me.txtConsulta.Visible = False
        '
        'lstConsulta
        '
        Me.lstConsulta.Cleanable = False
        Me.lstConsulta.EnterIndex = -1
        Me.lstConsulta.FormattingEnabled = True
        Me.lstConsulta.LabelAssociationKey = -1
        Me.lstConsulta.Location = New System.Drawing.Point(492, 32)
        Me.lstConsulta.Name = "lstConsulta"
        Me.lstConsulta.Size = New System.Drawing.Size(286, 108)
        Me.lstConsulta.TabIndex = 76
        Me.lstConsulta.Visible = False
        '
        'lblTotal
        '
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.ControlAssociationKey = -1
        Me.lblTotal.Location = New System.Drawing.Point(632, 540)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(146, 22)
        Me.lblTotal.TabIndex = 74
        Me.lblTotal.Text = "0,00"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Cleanable = False
        Me.btnLimpiar.EnterIndex = -1
        Me.btnLimpiar.LabelAssociationKey = -1
        Me.btnLimpiar.Location = New System.Drawing.Point(398, 145)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(88, 23)
        Me.btnLimpiar.TabIndex = 30
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.Cleanable = False
        Me.btnCerrar.EnterIndex = -1
        Me.btnCerrar.LabelAssociationKey = -1
        Me.btnCerrar.Location = New System.Drawing.Point(304, 145)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(88, 23)
        Me.btnCerrar.TabIndex = 29
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnIntereses
        '
        Me.btnIntereses.Cleanable = False
        Me.btnIntereses.EnterIndex = -1
        Me.btnIntereses.LabelAssociationKey = -1
        Me.btnIntereses.Location = New System.Drawing.Point(210, 145)
        Me.btnIntereses.Name = "btnIntereses"
        Me.btnIntereses.Size = New System.Drawing.Size(88, 23)
        Me.btnIntereses.TabIndex = 28
        Me.btnIntereses.Text = "Intereses"
        Me.btnIntereses.UseVisualStyleBackColor = True
        '
        'btnConsulta
        '
        Me.btnConsulta.Cleanable = False
        Me.btnConsulta.EnterIndex = -1
        Me.btnConsulta.LabelAssociationKey = -1
        Me.btnConsulta.Location = New System.Drawing.Point(116, 145)
        Me.btnConsulta.Name = "btnConsulta"
        Me.btnConsulta.Size = New System.Drawing.Size(88, 23)
        Me.btnConsulta.TabIndex = 27
        Me.btnConsulta.Text = "Consulta"
        Me.btnConsulta.UseVisualStyleBackColor = True
        '
        'btnAgregar
        '
        Me.btnAgregar.Cleanable = False
        Me.btnAgregar.EnterIndex = -1
        Me.btnAgregar.LabelAssociationKey = -1
        Me.btnAgregar.Location = New System.Drawing.Point(22, 145)
        Me.btnAgregar.Name = "btnAgregar"
        Me.btnAgregar.Size = New System.Drawing.Size(88, 23)
        Me.btnAgregar.TabIndex = 26
        Me.btnAgregar.Text = "Agregar"
        Me.btnAgregar.UseVisualStyleBackColor = True
        '
        'txtTotal
        '
        Me.txtTotal.Cleanable = True
        Me.txtTotal.Empty = False
        Me.txtTotal.EnterIndex = 8
        Me.txtTotal.LabelAssociationKey = 9
        Me.txtTotal.Location = New System.Drawing.Point(703, 147)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.Size = New System.Drawing.Size(75, 20)
        Me.txtTotal.TabIndex = 25
        Me.txtTotal.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel8
        '
        Me.CustomLabel8.AutoSize = True
        Me.CustomLabel8.ControlAssociationKey = 9
        Me.CustomLabel8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomLabel8.Location = New System.Drawing.Point(619, 150)
        Me.CustomLabel8.Name = "CustomLabel8"
        Me.CustomLabel8.Size = New System.Drawing.Size(80, 13)
        Me.CustomLabel8.TabIndex = 24
        Me.CustomLabel8.Text = "Total Recibo"
        '
        'CustomLabel9
        '
        Me.CustomLabel9.AutoSize = True
        Me.CustomLabel9.ControlAssociationKey = 8
        Me.CustomLabel9.Location = New System.Drawing.Point(19, 124)
        Me.CustomLabel9.Name = "CustomLabel9"
        Me.CustomLabel9.Size = New System.Drawing.Size(43, 13)
        Me.CustomLabel9.TabIndex = 23
        Me.CustomLabel9.Text = "Paridad"
        '
        'txtParidad
        '
        Me.txtParidad.Cleanable = True
        Me.txtParidad.Empty = True
        Me.txtParidad.EnterIndex = 9
        Me.txtParidad.LabelAssociationKey = 8
        Me.txtParidad.Location = New System.Drawing.Point(103, 121)
        Me.txtParidad.Name = "txtParidad"
        Me.txtParidad.Size = New System.Drawing.Size(75, 20)
        Me.txtParidad.TabIndex = 22
        Me.txtParidad.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'txtRetSuss
        '
        Me.txtRetSuss.Cleanable = True
        Me.txtRetSuss.Empty = False
        Me.txtRetSuss.EnterIndex = 7
        Me.txtRetSuss.LabelAssociationKey = 7
        Me.txtRetSuss.Location = New System.Drawing.Point(246, 95)
        Me.txtRetSuss.Name = "txtRetSuss"
        Me.txtRetSuss.Size = New System.Drawing.Size(75, 20)
        Me.txtRetSuss.TabIndex = 21
        Me.txtRetSuss.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = 7
        Me.CustomLabel6.Location = New System.Drawing.Point(181, 102)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel6.TabIndex = 20
        Me.CustomLabel6.Text = "Ret. Suss."
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = 5
        Me.CustomLabel7.Location = New System.Drawing.Point(184, 72)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel7.TabIndex = 19
        Me.CustomLabel7.Text = "Ret. I.B."
        '
        'txtRetIB
        '
        Me.txtRetIB.Cleanable = True
        Me.txtRetIB.Empty = False
        Me.txtRetIB.EnterIndex = 5
        Me.txtRetIB.LabelAssociationKey = 5
        Me.txtRetIB.Location = New System.Drawing.Point(246, 69)
        Me.txtRetIB.Name = "txtRetIB"
        Me.txtRetIB.Size = New System.Drawing.Size(75, 20)
        Me.txtRetIB.TabIndex = 18
        Me.txtRetIB.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'txtRetIva
        '
        Me.txtRetIva.Cleanable = True
        Me.txtRetIva.Empty = False
        Me.txtRetIva.EnterIndex = 6
        Me.txtRetIva.LabelAssociationKey = 6
        Me.txtRetIva.Location = New System.Drawing.Point(103, 95)
        Me.txtRetIva.Name = "txtRetIva"
        Me.txtRetIva.Size = New System.Drawing.Size(75, 20)
        Me.txtRetIva.TabIndex = 17
        Me.txtRetIva.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = 6
        Me.CustomLabel5.Location = New System.Drawing.Point(19, 98)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(47, 13)
        Me.CustomLabel5.TabIndex = 16
        Me.CustomLabel5.Text = "Ret. IVA"
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 4
        Me.CustomLabel4.Location = New System.Drawing.Point(19, 72)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(81, 13)
        Me.CustomLabel4.TabIndex = 15
        Me.CustomLabel4.Text = "Ret. Ganancias"
        '
        'txtRetGanancias
        '
        Me.txtRetGanancias.Cleanable = True
        Me.txtRetGanancias.Empty = False
        Me.txtRetGanancias.EnterIndex = 4
        Me.txtRetGanancias.LabelAssociationKey = 4
        Me.txtRetGanancias.Location = New System.Drawing.Point(103, 69)
        Me.txtRetGanancias.Name = "txtRetGanancias"
        Me.txtRetGanancias.Size = New System.Drawing.Size(75, 20)
        Me.txtRetGanancias.TabIndex = 3
        Me.txtRetGanancias.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'txtNombre
        '
        Me.txtNombre.Cleanable = True
        Me.txtNombre.Empty = False
        Me.txtNombre.Enabled = False
        Me.txtNombre.EnterIndex = -1
        Me.txtNombre.LabelAssociationKey = 3
        Me.txtNombre.Location = New System.Drawing.Point(184, 43)
        Me.txtNombre.MaxLength = 1000
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.Size = New System.Drawing.Size(302, 20)
        Me.txtNombre.TabIndex = 12
        Me.txtNombre.Validator = Administracion.ValidatorType.None
        '
        'txtCliente
        '
        Me.txtCliente.Cleanable = True
        Me.txtCliente.Empty = False
        Me.txtCliente.EnterIndex = 3
        Me.txtCliente.LabelAssociationKey = 3
        Me.txtCliente.Location = New System.Drawing.Point(103, 43)
        Me.txtCliente.MaxLength = 10
        Me.txtCliente.Name = "txtCliente"
        Me.txtCliente.Size = New System.Drawing.Size(75, 20)
        Me.txtCliente.TabIndex = 2
        Me.txtCliente.Validator = Administracion.ValidatorType.None
        '
        'txtRecibo
        '
        Me.txtRecibo.Cleanable = True
        Me.txtRecibo.Empty = False
        Me.txtRecibo.EnterIndex = 1
        Me.txtRecibo.LabelAssociationKey = 1
        Me.txtRecibo.Location = New System.Drawing.Point(103, 17)
        Me.txtRecibo.MaxLength = 10
        Me.txtRecibo.Name = "txtRecibo"
        Me.txtRecibo.Size = New System.Drawing.Size(75, 20)
        Me.txtRecibo.TabIndex = 0
        Me.txtRecibo.Validator = Administracion.ValidatorType.Numeric
        '
        'txtFecha
        '
        Me.txtFecha.Cleanable = True
        Me.txtFecha.Empty = False
        Me.txtFecha.EnterIndex = 2
        Me.txtFecha.LabelAssociationKey = 2
        Me.txtFecha.Location = New System.Drawing.Point(246, 17)
        Me.txtFecha.MaxLength = 10
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(75, 20)
        Me.txtFecha.TabIndex = 1
        Me.txtFecha.Validator = Administracion.ValidatorType.DateFormat
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = 2
        Me.CustomLabel3.Location = New System.Drawing.Point(184, 20)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(37, 13)
        Me.CustomLabel3.TabIndex = 3
        Me.CustomLabel3.Text = "Fecha"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = 3
        Me.CustomLabel2.Location = New System.Drawing.Point(19, 46)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(64, 13)
        Me.CustomLabel2.TabIndex = 1
        Me.CustomLabel2.Text = "Cod. Cliente"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = 1
        Me.CustomLabel1.Location = New System.Drawing.Point(19, 20)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(64, 13)
        Me.CustomLabel1.TabIndex = 0
        Me.CustomLabel1.Text = "Nro. Recibo"
        '
        'RecibosProvisorios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(790, 568)
        Me.Controls.Add(Me.CustomLabel10)
        Me.Controls.Add(Me.lstSeleccion)
        Me.Controls.Add(Me.txtConsulta)
        Me.Controls.Add(Me.lstConsulta)
        Me.Controls.Add(Me.lblTotal)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.btnIntereses)
        Me.Controls.Add(Me.btnConsulta)
        Me.Controls.Add(Me.btnAgregar)
        Me.Controls.Add(Me.txtTotal)
        Me.Controls.Add(Me.CustomLabel8)
        Me.Controls.Add(Me.CustomLabel9)
        Me.Controls.Add(Me.txtParidad)
        Me.Controls.Add(Me.txtRetSuss)
        Me.Controls.Add(Me.CustomLabel6)
        Me.Controls.Add(Me.CustomLabel7)
        Me.Controls.Add(Me.txtRetIB)
        Me.Controls.Add(Me.txtRetIva)
        Me.Controls.Add(Me.CustomLabel5)
        Me.Controls.Add(Me.CustomLabel4)
        Me.Controls.Add(Me.txtRetGanancias)
        Me.Controls.Add(Me.gridRecibos)
        Me.Controls.Add(Me.txtNombre)
        Me.Controls.Add(Me.txtCliente)
        Me.Controls.Add(Me.txtRecibo)
        Me.Controls.Add(Me.txtFecha)
        Me.Controls.Add(Me.CustomLabel3)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "RecibosProvisorios"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ingreso de Recibos Provisorios"
        CType(Me.gridRecibos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents CustomLabel3 As Administracion.CustomLabel
    Friend WithEvents txtFecha As Administracion.CustomTextBox
    Friend WithEvents txtRecibo As Administracion.CustomTextBox
    Friend WithEvents txtCliente As Administracion.CustomTextBox
    Friend WithEvents txtNombre As Administracion.CustomTextBox
    Friend WithEvents gridRecibos As System.Windows.Forms.DataGridView
    Friend WithEvents txtRetGanancias As Administracion.CustomTextBox
    Friend WithEvents CustomLabel4 As Administracion.CustomLabel
    Friend WithEvents CustomLabel5 As Administracion.CustomLabel
    Friend WithEvents txtRetIva As Administracion.CustomTextBox
    Friend WithEvents txtRetSuss As Administracion.CustomTextBox
    Friend WithEvents CustomLabel6 As Administracion.CustomLabel
    Friend WithEvents CustomLabel7 As Administracion.CustomLabel
    Friend WithEvents txtRetIB As Administracion.CustomTextBox
    Friend WithEvents txtTotal As Administracion.CustomTextBox
    Friend WithEvents CustomLabel8 As Administracion.CustomLabel
    Friend WithEvents CustomLabel9 As Administracion.CustomLabel
    Friend WithEvents txtParidad As Administracion.CustomTextBox
    Friend WithEvents btnAgregar As Administracion.CustomButton
    Friend WithEvents btnConsulta As Administracion.CustomButton
    Friend WithEvents btnCerrar As Administracion.CustomButton
    Friend WithEvents btnIntereses As Administracion.CustomButton
    Friend WithEvents btnLimpiar As Administracion.CustomButton
    Friend WithEvents lblTotal As Administracion.CustomLabel
    Friend WithEvents lstSeleccion As Administracion.CustomListBox
    Friend WithEvents txtConsulta As Administracion.CustomTextBox
    Friend WithEvents lstConsulta As Administracion.CustomListBox
    Friend WithEvents Tipo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents numero As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents fecha As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents banco As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents importe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustomLabel10 As Administracion.CustomLabel
End Class
