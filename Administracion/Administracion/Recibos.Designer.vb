<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Recibos
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
        Me.lstSeleccion = New Administracion.CustomListBox()
        Me.txtConsulta = New Administracion.CustomTextBox()
        Me.lstConsulta = New Administracion.CustomListBox()
        Me.lblTotalFormasPago = New Administracion.CustomLabel()
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
        Me.gridPagos = New System.Windows.Forms.DataGridView()
        Me.TipoCC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.LetraCC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.PuntoCC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.NumeroCC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ImporteCC = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.txtNombre = New Administracion.CustomTextBox()
        Me.txtCliente = New Administracion.CustomTextBox()
        Me.txtRecibo = New Administracion.CustomTextBox()
        Me.txtFecha = New Administracion.CustomTextBox()
        Me.CustomLabel3 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.CustomLabel1 = New Administracion.CustomLabel()
        Me.txtProvi = New Administracion.CustomTextBox()
        Me.CustomLabel10 = New Administracion.CustomLabel()
        Me.txtNombreCuenta = New Administracion.CustomTextBox()
        Me.txtCuenta = New Administracion.CustomTextBox()
        Me.CustomLabel11 = New Administracion.CustomLabel()
        Me.btnDias = New Administracion.CustomButton()
        Me.btnImpresion = New Administracion.CustomButton()
        Me.CustomLabel12 = New Administracion.CustomLabel()
        Me.gridFormasPago = New System.Windows.Forms.DataGridView()
        Me.Tipo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.numero = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.banco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.importe = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.lblTotalPagos = New Administracion.CustomLabel()
        CType(Me.gridPagos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridFormasPago, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lstSeleccion
        '
        Me.lstSeleccion.Cleanable = False
        Me.lstSeleccion.EnterIndex = -1
        Me.lstSeleccion.FormattingEnabled = True
        Me.lstSeleccion.LabelAssociationKey = -1
        Me.lstSeleccion.Location = New System.Drawing.Point(486, 3)
        Me.lstSeleccion.Name = "lstSeleccion"
        Me.lstSeleccion.Size = New System.Drawing.Size(286, 134)
        Me.lstSeleccion.TabIndex = 108
        Me.lstSeleccion.Visible = False
        '
        'txtConsulta
        '
        Me.txtConsulta.Cleanable = False
        Me.txtConsulta.Empty = True
        Me.txtConsulta.EnterIndex = -1
        Me.txtConsulta.LabelAssociationKey = -1
        Me.txtConsulta.Location = New System.Drawing.Point(486, 4)
        Me.txtConsulta.Name = "txtConsulta"
        Me.txtConsulta.Size = New System.Drawing.Size(286, 20)
        Me.txtConsulta.TabIndex = 107
        Me.txtConsulta.Validator = Administracion.ValidatorType.None
        Me.txtConsulta.Visible = False
        '
        'lstConsulta
        '
        Me.lstConsulta.Cleanable = False
        Me.lstConsulta.EnterIndex = -1
        Me.lstConsulta.FormattingEnabled = True
        Me.lstConsulta.LabelAssociationKey = -1
        Me.lstConsulta.Location = New System.Drawing.Point(486, 30)
        Me.lstConsulta.Name = "lstConsulta"
        Me.lstConsulta.Size = New System.Drawing.Size(286, 108)
        Me.lstConsulta.TabIndex = 106
        Me.lstConsulta.Visible = False
        '
        'lblTotalFormasPago
        '
        Me.lblTotalFormasPago.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalFormasPago.ControlAssociationKey = -1
        Me.lblTotalFormasPago.Location = New System.Drawing.Point(672, 539)
        Me.lblTotalFormasPago.Name = "lblTotalFormasPago"
        Me.lblTotalFormasPago.Size = New System.Drawing.Size(100, 22)
        Me.lblTotalFormasPago.TabIndex = 104
        Me.lblTotalFormasPago.Text = "0,00"
        Me.lblTotalFormasPago.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Cleanable = False
        Me.btnLimpiar.EnterIndex = -1
        Me.btnLimpiar.LabelAssociationKey = -1
        Me.btnLimpiar.Location = New System.Drawing.Point(684, 171)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(88, 23)
        Me.btnLimpiar.TabIndex = 103
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.Cleanable = False
        Me.btnCerrar.EnterIndex = -1
        Me.btnCerrar.LabelAssociationKey = -1
        Me.btnCerrar.Location = New System.Drawing.Point(590, 171)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(88, 23)
        Me.btnCerrar.TabIndex = 102
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnIntereses
        '
        Me.btnIntereses.Cleanable = False
        Me.btnIntereses.EnterIndex = -1
        Me.btnIntereses.LabelAssociationKey = -1
        Me.btnIntereses.Location = New System.Drawing.Point(496, 171)
        Me.btnIntereses.Name = "btnIntereses"
        Me.btnIntereses.Size = New System.Drawing.Size(88, 23)
        Me.btnIntereses.TabIndex = 101
        Me.btnIntereses.Text = "Intereses"
        Me.btnIntereses.UseVisualStyleBackColor = True
        '
        'btnConsulta
        '
        Me.btnConsulta.Cleanable = False
        Me.btnConsulta.EnterIndex = -1
        Me.btnConsulta.LabelAssociationKey = -1
        Me.btnConsulta.Location = New System.Drawing.Point(402, 171)
        Me.btnConsulta.Name = "btnConsulta"
        Me.btnConsulta.Size = New System.Drawing.Size(88, 23)
        Me.btnConsulta.TabIndex = 100
        Me.btnConsulta.Text = "Consulta"
        Me.btnConsulta.UseVisualStyleBackColor = True
        '
        'btnAgregar
        '
        Me.btnAgregar.Cleanable = False
        Me.btnAgregar.EnterIndex = -1
        Me.btnAgregar.LabelAssociationKey = -1
        Me.btnAgregar.Location = New System.Drawing.Point(308, 171)
        Me.btnAgregar.Name = "btnAgregar"
        Me.btnAgregar.Size = New System.Drawing.Size(88, 23)
        Me.btnAgregar.TabIndex = 99
        Me.btnAgregar.Text = "Agregar"
        Me.btnAgregar.UseVisualStyleBackColor = True
        '
        'txtTotal
        '
        Me.txtTotal.Cleanable = True
        Me.txtTotal.Empty = False
        Me.txtTotal.EnterIndex = 8
        Me.txtTotal.LabelAssociationKey = 9
        Me.txtTotal.Location = New System.Drawing.Point(697, 145)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.Size = New System.Drawing.Size(75, 20)
        Me.txtTotal.TabIndex = 98
        Me.txtTotal.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel8
        '
        Me.CustomLabel8.AutoSize = True
        Me.CustomLabel8.ControlAssociationKey = 9
        Me.CustomLabel8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomLabel8.Location = New System.Drawing.Point(613, 148)
        Me.CustomLabel8.Name = "CustomLabel8"
        Me.CustomLabel8.Size = New System.Drawing.Size(80, 13)
        Me.CustomLabel8.TabIndex = 97
        Me.CustomLabel8.Text = "Total Recibo"
        '
        'CustomLabel9
        '
        Me.CustomLabel9.AutoSize = True
        Me.CustomLabel9.ControlAssociationKey = 8
        Me.CustomLabel9.Location = New System.Drawing.Point(369, 74)
        Me.CustomLabel9.Name = "CustomLabel9"
        Me.CustomLabel9.Size = New System.Drawing.Size(43, 13)
        Me.CustomLabel9.TabIndex = 96
        Me.CustomLabel9.Text = "Paridad"
        '
        'txtParidad
        '
        Me.txtParidad.Cleanable = True
        Me.txtParidad.Empty = True
        Me.txtParidad.EnterIndex = 10
        Me.txtParidad.LabelAssociationKey = 8
        Me.txtParidad.Location = New System.Drawing.Point(358, 93)
        Me.txtParidad.Name = "txtParidad"
        Me.txtParidad.Size = New System.Drawing.Size(75, 20)
        Me.txtParidad.TabIndex = 95
        Me.txtParidad.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'txtRetSuss
        '
        Me.txtRetSuss.Cleanable = True
        Me.txtRetSuss.Empty = False
        Me.txtRetSuss.EnterIndex = 8
        Me.txtRetSuss.LabelAssociationKey = 7
        Me.txtRetSuss.Location = New System.Drawing.Point(230, 93)
        Me.txtRetSuss.Name = "txtRetSuss"
        Me.txtRetSuss.Size = New System.Drawing.Size(75, 20)
        Me.txtRetSuss.TabIndex = 94
        Me.txtRetSuss.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = 7
        Me.CustomLabel6.Location = New System.Drawing.Point(178, 100)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel6.TabIndex = 93
        Me.CustomLabel6.Text = "Ret. Suss."
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = 5
        Me.CustomLabel7.Location = New System.Drawing.Point(178, 74)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel7.TabIndex = 92
        Me.CustomLabel7.Text = "Ret. I.B."
        '
        'txtRetIB
        '
        Me.txtRetIB.Cleanable = True
        Me.txtRetIB.Empty = False
        Me.txtRetIB.EnterIndex = 6
        Me.txtRetIB.LabelAssociationKey = 5
        Me.txtRetIB.Location = New System.Drawing.Point(230, 67)
        Me.txtRetIB.Name = "txtRetIB"
        Me.txtRetIB.Size = New System.Drawing.Size(75, 20)
        Me.txtRetIB.TabIndex = 91
        Me.txtRetIB.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'txtRetIva
        '
        Me.txtRetIva.Cleanable = True
        Me.txtRetIva.Empty = False
        Me.txtRetIva.EnterIndex = 7
        Me.txtRetIva.LabelAssociationKey = 6
        Me.txtRetIva.Location = New System.Drawing.Point(97, 93)
        Me.txtRetIva.Name = "txtRetIva"
        Me.txtRetIva.Size = New System.Drawing.Size(75, 20)
        Me.txtRetIva.TabIndex = 90
        Me.txtRetIva.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = 6
        Me.CustomLabel5.Location = New System.Drawing.Point(12, 100)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(47, 13)
        Me.CustomLabel5.TabIndex = 89
        Me.CustomLabel5.Text = "Ret. IVA"
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 4
        Me.CustomLabel4.Location = New System.Drawing.Point(13, 74)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(81, 13)
        Me.CustomLabel4.TabIndex = 88
        Me.CustomLabel4.Text = "Ret. Ganancias"
        '
        'txtRetGanancias
        '
        Me.txtRetGanancias.Cleanable = True
        Me.txtRetGanancias.Empty = False
        Me.txtRetGanancias.EnterIndex = 5
        Me.txtRetGanancias.LabelAssociationKey = 4
        Me.txtRetGanancias.Location = New System.Drawing.Point(97, 67)
        Me.txtRetGanancias.Name = "txtRetGanancias"
        Me.txtRetGanancias.Size = New System.Drawing.Size(75, 20)
        Me.txtRetGanancias.TabIndex = 84
        Me.txtRetGanancias.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'gridPagos
        '
        Me.gridPagos.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.gridPagos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridPagos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.TipoCC, Me.LetraCC, Me.PuntoCC, Me.NumeroCC, Me.ImporteCC})
        Me.gridPagos.Location = New System.Drawing.Point(11, 200)
        Me.gridPagos.Name = "gridPagos"
        Me.gridPagos.RowHeadersWidth = 20
        Me.gridPagos.Size = New System.Drawing.Size(362, 335)
        Me.gridPagos.TabIndex = 20
        '
        'TipoCC
        '
        Me.TipoCC.FillWeight = 61.20156!
        Me.TipoCC.HeaderText = "Tipo"
        Me.TipoCC.Name = "TipoCC"
        Me.TipoCC.ReadOnly = True
        '
        'LetraCC
        '
        Me.LetraCC.FillWeight = 59.29222!
        Me.LetraCC.HeaderText = "Letra"
        Me.LetraCC.Name = "LetraCC"
        Me.LetraCC.ReadOnly = True
        '
        'PuntoCC
        '
        Me.PuntoCC.FillWeight = 55.11111!
        Me.PuntoCC.HeaderText = "Punto"
        Me.PuntoCC.Name = "PuntoCC"
        Me.PuntoCC.ReadOnly = True
        '
        'NumeroCC
        '
        Me.NumeroCC.FillWeight = 125.6672!
        Me.NumeroCC.HeaderText = "Numero"
        Me.NumeroCC.Name = "NumeroCC"
        Me.NumeroCC.ReadOnly = True
        '
        'ImporteCC
        '
        Me.ImporteCC.FillWeight = 125.6672!
        Me.ImporteCC.HeaderText = "Importe"
        Me.ImporteCC.Name = "ImporteCC"
        '
        'txtNombre
        '
        Me.txtNombre.Cleanable = True
        Me.txtNombre.Empty = False
        Me.txtNombre.Enabled = False
        Me.txtNombre.EnterIndex = -1
        Me.txtNombre.LabelAssociationKey = 3
        Me.txtNombre.Location = New System.Drawing.Point(178, 41)
        Me.txtNombre.MaxLength = 1000
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.Size = New System.Drawing.Size(302, 20)
        Me.txtNombre.TabIndex = 86
        Me.txtNombre.Validator = Administracion.ValidatorType.None
        '
        'txtCliente
        '
        Me.txtCliente.Cleanable = True
        Me.txtCliente.Empty = False
        Me.txtCliente.EnterIndex = 4
        Me.txtCliente.LabelAssociationKey = 3
        Me.txtCliente.Location = New System.Drawing.Point(97, 41)
        Me.txtCliente.MaxLength = 6
        Me.txtCliente.Name = "txtCliente"
        Me.txtCliente.Size = New System.Drawing.Size(75, 20)
        Me.txtCliente.TabIndex = 83
        Me.txtCliente.Validator = Administracion.ValidatorType.None
        '
        'txtRecibo
        '
        Me.txtRecibo.Cleanable = True
        Me.txtRecibo.Empty = False
        Me.txtRecibo.EnterIndex = 2
        Me.txtRecibo.LabelAssociationKey = 1
        Me.txtRecibo.Location = New System.Drawing.Point(97, 15)
        Me.txtRecibo.MaxLength = 6
        Me.txtRecibo.Name = "txtRecibo"
        Me.txtRecibo.Size = New System.Drawing.Size(75, 20)
        Me.txtRecibo.TabIndex = 79
        Me.txtRecibo.Validator = Administracion.ValidatorType.Numeric
        '
        'txtFecha
        '
        Me.txtFecha.Cleanable = True
        Me.txtFecha.Empty = False
        Me.txtFecha.EnterIndex = 3
        Me.txtFecha.LabelAssociationKey = 2
        Me.txtFecha.Location = New System.Drawing.Point(230, 15)
        Me.txtFecha.MaxLength = 10
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(75, 20)
        Me.txtFecha.TabIndex = 81
        Me.txtFecha.Validator = Administracion.ValidatorType.DateFormat
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = 2
        Me.CustomLabel3.Location = New System.Drawing.Point(178, 18)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(37, 13)
        Me.CustomLabel3.TabIndex = 85
        Me.CustomLabel3.Text = "Fecha"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = 3
        Me.CustomLabel2.Location = New System.Drawing.Point(12, 48)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(64, 13)
        Me.CustomLabel2.TabIndex = 82
        Me.CustomLabel2.Text = "Cod. Cliente"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = 1
        Me.CustomLabel1.Location = New System.Drawing.Point(13, 18)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(64, 13)
        Me.CustomLabel1.TabIndex = 80
        Me.CustomLabel1.Text = "Nro. Recibo"
        '
        'txtProvi
        '
        Me.txtProvi.Cleanable = True
        Me.txtProvi.Empty = True
        Me.txtProvi.EnterIndex = 1
        Me.txtProvi.LabelAssociationKey = 2
        Me.txtProvi.Location = New System.Drawing.Point(397, 15)
        Me.txtProvi.MaxLength = 6
        Me.txtProvi.Name = "txtProvi"
        Me.txtProvi.Size = New System.Drawing.Size(75, 20)
        Me.txtProvi.TabIndex = 109
        Me.txtProvi.Validator = Administracion.ValidatorType.Numeric
        '
        'CustomLabel10
        '
        Me.CustomLabel10.AutoSize = True
        Me.CustomLabel10.ControlAssociationKey = 2
        Me.CustomLabel10.Location = New System.Drawing.Point(312, 18)
        Me.CustomLabel10.Name = "CustomLabel10"
        Me.CustomLabel10.Size = New System.Drawing.Size(79, 13)
        Me.CustomLabel10.TabIndex = 110
        Me.CustomLabel10.Text = "Rec. Provisorio"
        '
        'txtNombreCuenta
        '
        Me.txtNombreCuenta.Cleanable = True
        Me.txtNombreCuenta.Empty = False
        Me.txtNombreCuenta.Enabled = False
        Me.txtNombreCuenta.EnterIndex = -1
        Me.txtNombreCuenta.LabelAssociationKey = 3
        Me.txtNombreCuenta.Location = New System.Drawing.Point(178, 119)
        Me.txtNombreCuenta.MaxLength = 1000
        Me.txtNombreCuenta.Name = "txtNombreCuenta"
        Me.txtNombreCuenta.Size = New System.Drawing.Size(302, 20)
        Me.txtNombreCuenta.TabIndex = 113
        Me.txtNombreCuenta.Validator = Administracion.ValidatorType.None
        '
        'txtCuenta
        '
        Me.txtCuenta.Cleanable = True
        Me.txtCuenta.Empty = False
        Me.txtCuenta.EnterIndex = 9
        Me.txtCuenta.LabelAssociationKey = 3
        Me.txtCuenta.Location = New System.Drawing.Point(97, 119)
        Me.txtCuenta.MaxLength = 6
        Me.txtCuenta.Name = "txtCuenta"
        Me.txtCuenta.Size = New System.Drawing.Size(75, 20)
        Me.txtCuenta.TabIndex = 112
        Me.txtCuenta.Validator = Administracion.ValidatorType.None
        '
        'CustomLabel11
        '
        Me.CustomLabel11.AutoSize = True
        Me.CustomLabel11.ControlAssociationKey = 3
        Me.CustomLabel11.Location = New System.Drawing.Point(12, 124)
        Me.CustomLabel11.Name = "CustomLabel11"
        Me.CustomLabel11.Size = New System.Drawing.Size(86, 13)
        Me.CustomLabel11.TabIndex = 111
        Me.CustomLabel11.Text = "Cuenta Contable"
        '
        'btnDias
        '
        Me.btnDias.Cleanable = False
        Me.btnDias.EnterIndex = -1
        Me.btnDias.LabelAssociationKey = -1
        Me.btnDias.Location = New System.Drawing.Point(402, 144)
        Me.btnDias.Name = "btnDias"
        Me.btnDias.Size = New System.Drawing.Size(88, 23)
        Me.btnDias.TabIndex = 114
        Me.btnDias.Text = "Dias"
        Me.btnDias.UseVisualStyleBackColor = True
        '
        'btnImpresion
        '
        Me.btnImpresion.Cleanable = False
        Me.btnImpresion.EnterIndex = -1
        Me.btnImpresion.LabelAssociationKey = -1
        Me.btnImpresion.Location = New System.Drawing.Point(308, 145)
        Me.btnImpresion.Name = "btnImpresion"
        Me.btnImpresion.Size = New System.Drawing.Size(88, 23)
        Me.btnImpresion.TabIndex = 115
        Me.btnImpresion.Text = "Impresion"
        Me.btnImpresion.UseVisualStyleBackColor = True
        '
        'CustomLabel12
        '
        Me.CustomLabel12.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.CustomLabel12.ControlAssociationKey = -1
        Me.CustomLabel12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomLabel12.Location = New System.Drawing.Point(379, 539)
        Me.CustomLabel12.Name = "CustomLabel12"
        Me.CustomLabel12.Size = New System.Drawing.Size(287, 22)
        Me.CustomLabel12.TabIndex = 118
        Me.CustomLabel12.Text = "Tipo Doc.: 1) Ef. 2) Ch. 4) Varios "
        Me.CustomLabel12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'gridFormasPago
        '
        Me.gridFormasPago.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.gridFormasPago.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridFormasPago.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tipo, Me.numero, Me.fecha, Me.banco, Me.importe})
        Me.gridFormasPago.Location = New System.Drawing.Point(379, 200)
        Me.gridFormasPago.Name = "gridFormasPago"
        Me.gridFormasPago.RowHeadersWidth = 20
        Me.gridFormasPago.Size = New System.Drawing.Size(393, 336)
        Me.gridFormasPago.TabIndex = 119
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
        Me.importe.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.DisplayedCells
        Me.importe.FillWeight = 80.0!
        Me.importe.HeaderText = "Importe"
        Me.importe.Name = "importe"
        Me.importe.Width = 67
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Location = New System.Drawing.Point(33, 144)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(245, 49)
        Me.GroupBox1.TabIndex = 120
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tipo"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(6, 19)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(97, 17)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Cobro Cta. Cte."
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(109, 19)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(68, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Anticipos"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(183, 19)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(54, 17)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.Text = "Varios"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'lblTotalPagos
        '
        Me.lblTotalPagos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPagos.ControlAssociationKey = -1
        Me.lblTotalPagos.Location = New System.Drawing.Point(280, 538)
        Me.lblTotalPagos.Name = "lblTotalPagos"
        Me.lblTotalPagos.Size = New System.Drawing.Size(93, 22)
        Me.lblTotalPagos.TabIndex = 121
        Me.lblTotalPagos.Text = "0,00"
        Me.lblTotalPagos.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Recibos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.lblTotalPagos)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.gridFormasPago)
        Me.Controls.Add(Me.CustomLabel12)
        Me.Controls.Add(Me.btnImpresion)
        Me.Controls.Add(Me.btnDias)
        Me.Controls.Add(Me.txtNombreCuenta)
        Me.Controls.Add(Me.txtCuenta)
        Me.Controls.Add(Me.CustomLabel11)
        Me.Controls.Add(Me.txtProvi)
        Me.Controls.Add(Me.CustomLabel10)
        Me.Controls.Add(Me.lstSeleccion)
        Me.Controls.Add(Me.txtConsulta)
        Me.Controls.Add(Me.lstConsulta)
        Me.Controls.Add(Me.lblTotalFormasPago)
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
        Me.Controls.Add(Me.gridPagos)
        Me.Controls.Add(Me.txtNombre)
        Me.Controls.Add(Me.txtCliente)
        Me.Controls.Add(Me.txtRecibo)
        Me.Controls.Add(Me.txtFecha)
        Me.Controls.Add(Me.CustomLabel3)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Name = "Recibos"
        Me.Text = "Ingreso de Recibos"
        CType(Me.gridPagos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridFormasPago, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lstSeleccion As Administracion.CustomListBox
    Friend WithEvents txtConsulta As Administracion.CustomTextBox
    Friend WithEvents lstConsulta As Administracion.CustomListBox
    Friend WithEvents lblTotalFormasPago As Administracion.CustomLabel
    Friend WithEvents btnLimpiar As Administracion.CustomButton
    Friend WithEvents btnCerrar As Administracion.CustomButton
    Friend WithEvents btnIntereses As Administracion.CustomButton
    Friend WithEvents btnConsulta As Administracion.CustomButton
    Friend WithEvents btnAgregar As Administracion.CustomButton
    Friend WithEvents txtTotal As Administracion.CustomTextBox
    Friend WithEvents CustomLabel8 As Administracion.CustomLabel
    Friend WithEvents CustomLabel9 As Administracion.CustomLabel
    Friend WithEvents txtParidad As Administracion.CustomTextBox
    Friend WithEvents txtRetSuss As Administracion.CustomTextBox
    Friend WithEvents CustomLabel6 As Administracion.CustomLabel
    Friend WithEvents CustomLabel7 As Administracion.CustomLabel
    Friend WithEvents txtRetIB As Administracion.CustomTextBox
    Friend WithEvents txtRetIva As Administracion.CustomTextBox
    Friend WithEvents CustomLabel5 As Administracion.CustomLabel
    Friend WithEvents CustomLabel4 As Administracion.CustomLabel
    Friend WithEvents txtRetGanancias As Administracion.CustomTextBox
    Friend WithEvents gridPagos As System.Windows.Forms.DataGridView
    Friend WithEvents txtNombre As Administracion.CustomTextBox
    Friend WithEvents txtCliente As Administracion.CustomTextBox
    Friend WithEvents txtRecibo As Administracion.CustomTextBox
    Friend WithEvents txtFecha As Administracion.CustomTextBox
    Friend WithEvents CustomLabel3 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents txtProvi As Administracion.CustomTextBox
    Friend WithEvents CustomLabel10 As Administracion.CustomLabel
    Friend WithEvents txtNombreCuenta As Administracion.CustomTextBox
    Friend WithEvents txtCuenta As Administracion.CustomTextBox
    Friend WithEvents CustomLabel11 As Administracion.CustomLabel
    Friend WithEvents btnDias As Administracion.CustomButton
    Friend WithEvents btnImpresion As Administracion.CustomButton
    Friend WithEvents CustomLabel12 As Administracion.CustomLabel
    Friend WithEvents TipoCC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents LetraCC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents PuntoCC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents NumeroCC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ImporteCC As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents gridFormasPago As System.Windows.Forms.DataGridView
    Friend WithEvents Tipo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents numero As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents fecha As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents banco As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents importe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents lblTotalPagos As Administracion.CustomLabel
End Class
