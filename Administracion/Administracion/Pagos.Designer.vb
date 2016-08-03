<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Pagos
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optTransferencias = New System.Windows.Forms.RadioButton()
        Me.optAnticipos = New System.Windows.Forms.RadioButton()
        Me.optChequeRechazado = New System.Windows.Forms.RadioButton()
        Me.optVarios = New System.Windows.Forms.RadioButton()
        Me.optCtaCte = New System.Windows.Forms.RadioButton()
        Me.gridPagos = New System.Windows.Forms.DataGridView()
        Me.gridFormaPagos = New System.Windows.Forms.DataGridView()
        Me.Tipo2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Numero2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Banco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nombre = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Importe2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CustomLabel13 = New Administracion.CustomLabel()
        Me.lblDiferencia = New Administracion.CustomLabel()
        Me.lblFormaPagos = New Administracion.CustomLabel()
        Me.lblPagos = New Administracion.CustomLabel()
        Me.lstSeleccion = New Administracion.CustomListBox()
        Me.btnCarpetas = New Administracion.CustomButton()
        Me.btnImprimir = New Administracion.CustomButton()
        Me.btnCalcular = New Administracion.CustomButton()
        Me.btnConsulta = New Administracion.CustomButton()
        Me.btnCerrar = New Administracion.CustomButton()
        Me.btnLimpiar = New Administracion.CustomButton()
        Me.btnAgregar = New Administracion.CustomButton()
        Me.txtConsulta = New Administracion.CustomTextBox()
        Me.lstConsulta = New Administracion.CustomListBox()
        Me.txtTotal = New Administracion.CustomTextBox()
        Me.CustomLabel11 = New Administracion.CustomLabel()
        Me.txtIVA = New Administracion.CustomTextBox()
        Me.txtIngresosBrutos = New Administracion.CustomTextBox()
        Me.CustomLabel10 = New Administracion.CustomLabel()
        Me.txtIBCiudad = New Administracion.CustomTextBox()
        Me.CustomLabel9 = New Administracion.CustomLabel()
        Me.txtGanancias = New Administracion.CustomTextBox()
        Me.CustomLabel8 = New Administracion.CustomLabel()
        Me.lblGanancias = New Administracion.CustomLabel()
        Me.cmbTipo = New Administracion.CustomComboBox()
        Me.txtParidad = New Administracion.CustomTextBox()
        Me.txtFechaParidad = New Administracion.CustomTextBox()
        Me.CustomLabel7 = New Administracion.CustomLabel()
        Me.CustomLabel6 = New Administracion.CustomLabel()
        Me.txtNombreBanco = New Administracion.CustomTextBox()
        Me.txtBanco = New Administracion.CustomTextBox()
        Me.txtObservaciones = New Administracion.CustomTextBox()
        Me.txtRazonSocial = New Administracion.CustomTextBox()
        Me.txtProveedor = New Administracion.CustomTextBox()
        Me.txtFecha = New Administracion.CustomTextBox()
        Me.txtOrdenPago = New Administracion.CustomTextBox()
        Me.CustomLabel5 = New Administracion.CustomLabel()
        Me.CustomLabel4 = New Administracion.CustomLabel()
        Me.CustomLabel3 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.CustomLabel1 = New Administracion.CustomLabel()
        Me.CustomLabel12 = New Administracion.CustomLabel()
        Me.Tipo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Letra = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Punto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Numero = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Importe = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Descripcion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        CType(Me.gridPagos, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.gridFormaPagos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.optTransferencias)
        Me.GroupBox1.Controls.Add(Me.optAnticipos)
        Me.GroupBox1.Controls.Add(Me.optChequeRechazado)
        Me.GroupBox1.Controls.Add(Me.optVarios)
        Me.GroupBox1.Controls.Add(Me.optCtaCte)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 129)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(239, 99)
        Me.GroupBox1.TabIndex = 34
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tipo de Orden de Pago"
        '
        'optTransferencias
        '
        Me.optTransferencias.AutoSize = True
        Me.optTransferencias.Location = New System.Drawing.Point(134, 46)
        Me.optTransferencias.Name = "optTransferencias"
        Me.optTransferencias.Size = New System.Drawing.Size(95, 17)
        Me.optTransferencias.TabIndex = 4
        Me.optTransferencias.Text = "Transferencias"
        Me.optTransferencias.UseVisualStyleBackColor = True
        '
        'optAnticipos
        '
        Me.optAnticipos.AutoSize = True
        Me.optAnticipos.Location = New System.Drawing.Point(134, 23)
        Me.optAnticipos.Name = "optAnticipos"
        Me.optAnticipos.Size = New System.Drawing.Size(68, 17)
        Me.optAnticipos.TabIndex = 3
        Me.optAnticipos.Text = "Anticipos"
        Me.optAnticipos.UseVisualStyleBackColor = True
        '
        'optChequeRechazado
        '
        Me.optChequeRechazado.AutoSize = True
        Me.optChequeRechazado.Location = New System.Drawing.Point(6, 69)
        Me.optChequeRechazado.Name = "optChequeRechazado"
        Me.optChequeRechazado.Size = New System.Drawing.Size(104, 17)
        Me.optChequeRechazado.TabIndex = 2
        Me.optChequeRechazado.Text = "Ch. Rechazados"
        Me.optChequeRechazado.UseVisualStyleBackColor = True
        '
        'optVarios
        '
        Me.optVarios.AutoSize = True
        Me.optVarios.Location = New System.Drawing.Point(6, 46)
        Me.optVarios.Name = "optVarios"
        Me.optVarios.Size = New System.Drawing.Size(87, 17)
        Me.optVarios.TabIndex = 1
        Me.optVarios.Text = "Pagos Varios"
        Me.optVarios.UseVisualStyleBackColor = True
        '
        'optCtaCte
        '
        Me.optCtaCte.AutoSize = True
        Me.optCtaCte.Checked = True
        Me.optCtaCte.Location = New System.Drawing.Point(6, 23)
        Me.optCtaCte.Name = "optCtaCte"
        Me.optCtaCte.Size = New System.Drawing.Size(99, 17)
        Me.optCtaCte.TabIndex = 0
        Me.optCtaCte.TabStop = True
        Me.optCtaCte.Text = "Pagos Cta. Cte."
        Me.optCtaCte.UseVisualStyleBackColor = True
        '
        'gridPagos
        '
        Me.gridPagos.AllowUserToAddRows = False
        Me.gridPagos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridPagos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tipo, Me.Letra, Me.Punto, Me.Numero, Me.Importe, Me.Descripcion})
        Me.gridPagos.Location = New System.Drawing.Point(0, 272)
        Me.gridPagos.Name = "gridPagos"
        Me.gridPagos.RowHeadersWidth = 10
        Me.gridPagos.Size = New System.Drawing.Size(396, 273)
        Me.gridPagos.TabIndex = 56
        '
        'gridFormaPagos
        '
        Me.gridFormaPagos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridFormaPagos.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tipo2, Me.Numero2, Me.Fecha, Me.Banco, Me.Nombre, Me.Importe2})
        Me.gridFormaPagos.Location = New System.Drawing.Point(396, 272)
        Me.gridFormaPagos.Name = "gridFormaPagos"
        Me.gridFormaPagos.RowHeadersWidth = 10
        Me.gridFormaPagos.Size = New System.Drawing.Size(398, 273)
        Me.gridFormaPagos.TabIndex = 57
        '
        'Tipo2
        '
        Me.Tipo2.HeaderText = "Tipo"
        Me.Tipo2.Name = "Tipo2"
        Me.Tipo2.Width = 35
        '
        'Numero2
        '
        Me.Numero2.HeaderText = "Número"
        Me.Numero2.Name = "Numero2"
        Me.Numero2.Width = 70
        '
        'Fecha
        '
        Me.Fecha.HeaderText = "Fecha"
        Me.Fecha.Name = "Fecha"
        Me.Fecha.Width = 75
        '
        'Banco
        '
        Me.Banco.HeaderText = "Banco"
        Me.Banco.Name = "Banco"
        Me.Banco.Width = 45
        '
        'Nombre
        '
        Me.Nombre.HeaderText = "Nombre"
        Me.Nombre.Name = "Nombre"
        Me.Nombre.Width = 80
        '
        'Importe2
        '
        Me.Importe2.HeaderText = "Importe"
        Me.Importe2.Name = "Importe2"
        Me.Importe2.Width = 80
        '
        'CustomLabel13
        '
        Me.CustomLabel13.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.CustomLabel13.ControlAssociationKey = -1
        Me.CustomLabel13.Location = New System.Drawing.Point(284, 548)
        Me.CustomLabel13.Name = "CustomLabel13"
        Me.CustomLabel13.Size = New System.Drawing.Size(342, 22)
        Me.CustomLabel13.TabIndex = 71
        Me.CustomLabel13.Text = "Tipo de Doc.:   1) Ef.   2) Bco.   3) Ch. Terceros   5) US$   6) Varios"
        Me.CustomLabel13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblDiferencia
        '
        Me.lblDiferencia.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblDiferencia.ControlAssociationKey = -1
        Me.lblDiferencia.Location = New System.Drawing.Point(632, 548)
        Me.lblDiferencia.Name = "lblDiferencia"
        Me.lblDiferencia.Size = New System.Drawing.Size(70, 22)
        Me.lblDiferencia.TabIndex = 70
        Me.lblDiferencia.Text = "0,00"
        Me.lblDiferencia.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFormaPagos
        '
        Me.lblFormaPagos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFormaPagos.ControlAssociationKey = -1
        Me.lblFormaPagos.Location = New System.Drawing.Point(708, 548)
        Me.lblFormaPagos.Name = "lblFormaPagos"
        Me.lblFormaPagos.Size = New System.Drawing.Size(70, 22)
        Me.lblFormaPagos.TabIndex = 69
        Me.lblFormaPagos.Text = "0,00"
        Me.lblFormaPagos.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPagos
        '
        Me.lblPagos.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPagos.ControlAssociationKey = -1
        Me.lblPagos.Location = New System.Drawing.Point(208, 548)
        Me.lblPagos.Name = "lblPagos"
        Me.lblPagos.Size = New System.Drawing.Size(70, 22)
        Me.lblPagos.TabIndex = 68
        Me.lblPagos.Text = "0,00"
        Me.lblPagos.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lstSeleccion
        '
        Me.lstSeleccion.Cleanable = False
        Me.lstSeleccion.EnterIndex = -1
        Me.lstSeleccion.FormattingEnabled = True
        Me.lstSeleccion.LabelAssociationKey = -1
        Me.lstSeleccion.Location = New System.Drawing.Point(437, 11)
        Me.lstSeleccion.Name = "lstSeleccion"
        Me.lstSeleccion.Size = New System.Drawing.Size(333, 134)
        Me.lstSeleccion.TabIndex = 66
        Me.lstSeleccion.Visible = False
        '
        'btnCarpetas
        '
        Me.btnCarpetas.Cleanable = False
        Me.btnCarpetas.EnterIndex = -1
        Me.btnCarpetas.LabelAssociationKey = -1
        Me.btnCarpetas.Location = New System.Drawing.Point(159, 234)
        Me.btnCarpetas.Name = "btnCarpetas"
        Me.btnCarpetas.Size = New System.Drawing.Size(100, 25)
        Me.btnCarpetas.TabIndex = 65
        Me.btnCarpetas.Text = "Carpetas"
        Me.btnCarpetas.UseVisualStyleBackColor = True
        '
        'btnImprimir
        '
        Me.btnImprimir.Cleanable = False
        Me.btnImprimir.EnterIndex = -1
        Me.btnImprimir.LabelAssociationKey = -1
        Me.btnImprimir.Location = New System.Drawing.Point(265, 234)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(100, 25)
        Me.btnImprimir.TabIndex = 64
        Me.btnImprimir.Text = "Impresión"
        Me.btnImprimir.UseVisualStyleBackColor = True
        '
        'btnCalcular
        '
        Me.btnCalcular.Cleanable = False
        Me.btnCalcular.EnterIndex = -1
        Me.btnCalcular.LabelAssociationKey = -1
        Me.btnCalcular.Location = New System.Drawing.Point(583, 234)
        Me.btnCalcular.Name = "btnCalcular"
        Me.btnCalcular.Size = New System.Drawing.Size(100, 25)
        Me.btnCalcular.TabIndex = 63
        Me.btnCalcular.Text = "Calc. Ret."
        Me.btnCalcular.UseVisualStyleBackColor = True
        '
        'btnConsulta
        '
        Me.btnConsulta.Cleanable = False
        Me.btnConsulta.EnterIndex = -1
        Me.btnConsulta.LabelAssociationKey = -1
        Me.btnConsulta.Location = New System.Drawing.Point(477, 234)
        Me.btnConsulta.Name = "btnConsulta"
        Me.btnConsulta.Size = New System.Drawing.Size(100, 25)
        Me.btnConsulta.TabIndex = 62
        Me.btnConsulta.Text = "Consulta"
        Me.btnConsulta.UseVisualStyleBackColor = True
        '
        'btnCerrar
        '
        Me.btnCerrar.Cleanable = False
        Me.btnCerrar.EnterIndex = -1
        Me.btnCerrar.LabelAssociationKey = -1
        Me.btnCerrar.Location = New System.Drawing.Point(371, 203)
        Me.btnCerrar.Name = "btnCerrar"
        Me.btnCerrar.Size = New System.Drawing.Size(100, 25)
        Me.btnCerrar.TabIndex = 61
        Me.btnCerrar.Text = "Cerrar"
        Me.btnCerrar.UseVisualStyleBackColor = True
        '
        'btnLimpiar
        '
        Me.btnLimpiar.Cleanable = False
        Me.btnLimpiar.EnterIndex = -1
        Me.btnLimpiar.LabelAssociationKey = -1
        Me.btnLimpiar.Location = New System.Drawing.Point(371, 234)
        Me.btnLimpiar.Name = "btnLimpiar"
        Me.btnLimpiar.Size = New System.Drawing.Size(100, 25)
        Me.btnLimpiar.TabIndex = 60
        Me.btnLimpiar.Text = "Limpiar"
        Me.btnLimpiar.UseVisualStyleBackColor = True
        '
        'btnAgregar
        '
        Me.btnAgregar.Cleanable = False
        Me.btnAgregar.EnterIndex = -1
        Me.btnAgregar.LabelAssociationKey = -1
        Me.btnAgregar.Location = New System.Drawing.Point(265, 203)
        Me.btnAgregar.Name = "btnAgregar"
        Me.btnAgregar.Size = New System.Drawing.Size(100, 25)
        Me.btnAgregar.TabIndex = 58
        Me.btnAgregar.Text = "Grabar"
        Me.btnAgregar.UseVisualStyleBackColor = True
        '
        'txtConsulta
        '
        Me.txtConsulta.Cleanable = False
        Me.txtConsulta.Empty = True
        Me.txtConsulta.EnterIndex = -1
        Me.txtConsulta.LabelAssociationKey = -1
        Me.txtConsulta.Location = New System.Drawing.Point(437, 12)
        Me.txtConsulta.Name = "txtConsulta"
        Me.txtConsulta.Size = New System.Drawing.Size(333, 20)
        Me.txtConsulta.TabIndex = 55
        Me.txtConsulta.Validator = Administracion.ValidatorType.None
        Me.txtConsulta.Visible = False
        '
        'lstConsulta
        '
        Me.lstConsulta.Cleanable = False
        Me.lstConsulta.EnterIndex = -1
        Me.lstConsulta.FormattingEnabled = True
        Me.lstConsulta.LabelAssociationKey = -1
        Me.lstConsulta.Location = New System.Drawing.Point(437, 38)
        Me.lstConsulta.Name = "lstConsulta"
        Me.lstConsulta.Size = New System.Drawing.Size(333, 108)
        Me.lstConsulta.TabIndex = 54
        Me.lstConsulta.Visible = False
        '
        'txtTotal
        '
        Me.txtTotal.Cleanable = True
        Me.txtTotal.Empty = False
        Me.txtTotal.EnterIndex = -1
        Me.txtTotal.LabelAssociationKey = 12
        Me.txtTotal.Location = New System.Drawing.Point(603, 208)
        Me.txtTotal.Name = "txtTotal"
        Me.txtTotal.ReadOnly = True
        Me.txtTotal.Size = New System.Drawing.Size(91, 20)
        Me.txtTotal.TabIndex = 53
        Me.txtTotal.Validator = Administracion.ValidatorType.Float
        '
        'CustomLabel11
        '
        Me.CustomLabel11.AutoSize = True
        Me.CustomLabel11.ControlAssociationKey = 12
        Me.CustomLabel11.Location = New System.Drawing.Point(503, 211)
        Me.CustomLabel11.Name = "CustomLabel11"
        Me.CustomLabel11.Size = New System.Drawing.Size(94, 13)
        Me.CustomLabel11.TabIndex = 52
        Me.CustomLabel11.Text = "Total Retenciones"
        '
        'txtIVA
        '
        Me.txtIVA.Cleanable = True
        Me.txtIVA.Empty = False
        Me.txtIVA.EnterIndex = 11
        Me.txtIVA.LabelAssociationKey = 11
        Me.txtIVA.Location = New System.Drawing.Point(685, 178)
        Me.txtIVA.Name = "txtIVA"
        Me.txtIVA.ReadOnly = True
        Me.txtIVA.Size = New System.Drawing.Size(75, 20)
        Me.txtIVA.TabIndex = 51
        Me.txtIVA.Text = "0,00"
        Me.txtIVA.Validator = Administracion.ValidatorType.Float
        '
        'txtIngresosBrutos
        '
        Me.txtIngresosBrutos.Cleanable = True
        Me.txtIngresosBrutos.Empty = False
        Me.txtIngresosBrutos.EnterIndex = 9
        Me.txtIngresosBrutos.LabelAssociationKey = 9
        Me.txtIngresosBrutos.Location = New System.Drawing.Point(685, 151)
        Me.txtIngresosBrutos.Name = "txtIngresosBrutos"
        Me.txtIngresosBrutos.ReadOnly = True
        Me.txtIngresosBrutos.Size = New System.Drawing.Size(75, 20)
        Me.txtIngresosBrutos.TabIndex = 50
        Me.txtIngresosBrutos.Text = "0,00"
        Me.txtIngresosBrutos.Validator = Administracion.ValidatorType.Float
        '
        'CustomLabel10
        '
        Me.CustomLabel10.AutoSize = True
        Me.CustomLabel10.ControlAssociationKey = 11
        Me.CustomLabel10.Location = New System.Drawing.Point(597, 181)
        Me.CustomLabel10.Name = "CustomLabel10"
        Me.CustomLabel10.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel10.TabIndex = 49
        Me.CustomLabel10.Text = "Ret. I.V.A."
        '
        'txtIBCiudad
        '
        Me.txtIBCiudad.Cleanable = True
        Me.txtIBCiudad.Empty = False
        Me.txtIBCiudad.EnterIndex = 10
        Me.txtIBCiudad.LabelAssociationKey = 10
        Me.txtIBCiudad.Location = New System.Drawing.Point(516, 178)
        Me.txtIBCiudad.Name = "txtIBCiudad"
        Me.txtIBCiudad.ReadOnly = True
        Me.txtIBCiudad.Size = New System.Drawing.Size(75, 20)
        Me.txtIBCiudad.TabIndex = 48
        Me.txtIBCiudad.Text = "0,00"
        Me.txtIBCiudad.Validator = Administracion.ValidatorType.Float
        '
        'CustomLabel9
        '
        Me.CustomLabel9.AutoSize = True
        Me.CustomLabel9.ControlAssociationKey = 10
        Me.CustomLabel9.Location = New System.Drawing.Point(434, 181)
        Me.CustomLabel9.Name = "CustomLabel9"
        Me.CustomLabel9.Size = New System.Drawing.Size(76, 13)
        Me.CustomLabel9.TabIndex = 47
        Me.CustomLabel9.Text = "Ret. IB Ciudad"
        '
        'txtGanancias
        '
        Me.txtGanancias.Cleanable = True
        Me.txtGanancias.Empty = False
        Me.txtGanancias.EnterIndex = 8
        Me.txtGanancias.LabelAssociationKey = 8
        Me.txtGanancias.Location = New System.Drawing.Point(516, 151)
        Me.txtGanancias.Name = "txtGanancias"
        Me.txtGanancias.ReadOnly = True
        Me.txtGanancias.Size = New System.Drawing.Size(75, 20)
        Me.txtGanancias.TabIndex = 46
        Me.txtGanancias.Text = "0,00"
        Me.txtGanancias.Validator = Administracion.ValidatorType.Float
        '
        'CustomLabel8
        '
        Me.CustomLabel8.AutoSize = True
        Me.CustomLabel8.ControlAssociationKey = 9
        Me.CustomLabel8.Location = New System.Drawing.Point(598, 154)
        Me.CustomLabel8.Name = "CustomLabel8"
        Me.CustomLabel8.Size = New System.Drawing.Size(81, 13)
        Me.CustomLabel8.TabIndex = 41
        Me.CustomLabel8.Text = "Ret. Ing. Brutos"
        '
        'lblGanancias
        '
        Me.lblGanancias.AutoSize = True
        Me.lblGanancias.ControlAssociationKey = 8
        Me.lblGanancias.Location = New System.Drawing.Point(434, 154)
        Me.lblGanancias.Name = "lblGanancias"
        Me.lblGanancias.Size = New System.Drawing.Size(76, 13)
        Me.lblGanancias.TabIndex = 40
        Me.lblGanancias.Text = "Ret. Ganancia"
        '
        'cmbTipo
        '
        Me.cmbTipo.Cleanable = True
        Me.cmbTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTipo.Empty = False
        Me.cmbTipo.EnterIndex = -1
        Me.cmbTipo.FormattingEnabled = True
        Me.cmbTipo.Items.AddRange(New Object() {"Normal", "Cheque Rechazado"})
        Me.cmbTipo.LabelAssociationKey = 13
        Me.cmbTipo.Location = New System.Drawing.Point(268, 177)
        Me.cmbTipo.Name = "cmbTipo"
        Me.cmbTipo.Size = New System.Drawing.Size(160, 21)
        Me.cmbTipo.TabIndex = 39
        Me.cmbTipo.Validator = Administracion.ValidatorType.None
        Me.cmbTipo.Visible = False
        '
        'txtParidad
        '
        Me.txtParidad.Cleanable = True
        Me.txtParidad.Empty = True
        Me.txtParidad.EnterIndex = 7
        Me.txtParidad.LabelAssociationKey = 7
        Me.txtParidad.Location = New System.Drawing.Point(347, 151)
        Me.txtParidad.Name = "txtParidad"
        Me.txtParidad.Size = New System.Drawing.Size(81, 20)
        Me.txtParidad.TabIndex = 38
        Me.txtParidad.Validator = Administracion.ValidatorType.StrictlyPositiveFloat
        '
        'txtFechaParidad
        '
        Me.txtFechaParidad.Cleanable = True
        Me.txtFechaParidad.Empty = True
        Me.txtFechaParidad.EnterIndex = 6
        Me.txtFechaParidad.LabelAssociationKey = 6
        Me.txtFechaParidad.Location = New System.Drawing.Point(347, 126)
        Me.txtFechaParidad.MaxLength = 10
        Me.txtFechaParidad.Name = "txtFechaParidad"
        Me.txtFechaParidad.Size = New System.Drawing.Size(81, 20)
        Me.txtFechaParidad.TabIndex = 37
        Me.txtFechaParidad.Validator = Administracion.ValidatorType.DateFormat
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = 7
        Me.CustomLabel7.Location = New System.Drawing.Point(265, 156)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(43, 13)
        Me.CustomLabel7.TabIndex = 36
        Me.CustomLabel7.Text = "Paridad"
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = 6
        Me.CustomLabel6.Location = New System.Drawing.Point(265, 129)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(76, 13)
        Me.CustomLabel6.TabIndex = 35
        Me.CustomLabel6.Text = "Fecha Paridad"
        '
        'txtNombreBanco
        '
        Me.txtNombreBanco.Cleanable = True
        Me.txtNombreBanco.Empty = True
        Me.txtNombreBanco.Enabled = False
        Me.txtNombreBanco.EnterIndex = -1
        Me.txtNombreBanco.LabelAssociationKey = 5
        Me.txtNombreBanco.Location = New System.Drawing.Point(187, 98)
        Me.txtNombreBanco.Name = "txtNombreBanco"
        Me.txtNombreBanco.Size = New System.Drawing.Size(241, 20)
        Me.txtNombreBanco.TabIndex = 33
        Me.txtNombreBanco.Validator = Administracion.ValidatorType.None
        '
        'txtBanco
        '
        Me.txtBanco.Cleanable = True
        Me.txtBanco.Empty = True
        Me.txtBanco.Enabled = False
        Me.txtBanco.EnterIndex = 5
        Me.txtBanco.LabelAssociationKey = 5
        Me.txtBanco.Location = New System.Drawing.Point(105, 98)
        Me.txtBanco.MaxLength = 8
        Me.txtBanco.Name = "txtBanco"
        Me.txtBanco.Size = New System.Drawing.Size(76, 20)
        Me.txtBanco.TabIndex = 32
        Me.txtBanco.Validator = Administracion.ValidatorType.Numeric
        '
        'txtObservaciones
        '
        Me.txtObservaciones.Cleanable = True
        Me.txtObservaciones.Empty = True
        Me.txtObservaciones.EnterIndex = 4
        Me.txtObservaciones.LabelAssociationKey = 4
        Me.txtObservaciones.Location = New System.Drawing.Point(105, 71)
        Me.txtObservaciones.MaxLength = 50
        Me.txtObservaciones.Name = "txtObservaciones"
        Me.txtObservaciones.Size = New System.Drawing.Size(323, 20)
        Me.txtObservaciones.TabIndex = 31
        Me.txtObservaciones.Validator = Administracion.ValidatorType.None
        '
        'txtRazonSocial
        '
        Me.txtRazonSocial.Cleanable = True
        Me.txtRazonSocial.Empty = False
        Me.txtRazonSocial.Enabled = False
        Me.txtRazonSocial.EnterIndex = -1
        Me.txtRazonSocial.LabelAssociationKey = 3
        Me.txtRazonSocial.Location = New System.Drawing.Point(187, 44)
        Me.txtRazonSocial.Name = "txtRazonSocial"
        Me.txtRazonSocial.Size = New System.Drawing.Size(241, 20)
        Me.txtRazonSocial.TabIndex = 30
        Me.txtRazonSocial.Validator = Administracion.ValidatorType.None
        '
        'txtProveedor
        '
        Me.txtProveedor.Cleanable = True
        Me.txtProveedor.Empty = False
        Me.txtProveedor.EnterIndex = 3
        Me.txtProveedor.LabelAssociationKey = 3
        Me.txtProveedor.Location = New System.Drawing.Point(105, 44)
        Me.txtProveedor.MaxLength = 11
        Me.txtProveedor.Name = "txtProveedor"
        Me.txtProveedor.Size = New System.Drawing.Size(76, 20)
        Me.txtProveedor.TabIndex = 29
        Me.txtProveedor.Validator = Administracion.ValidatorType.None
        '
        'txtFecha
        '
        Me.txtFecha.Cleanable = True
        Me.txtFecha.Empty = False
        Me.txtFecha.EnterIndex = 2
        Me.txtFecha.LabelAssociationKey = 2
        Me.txtFecha.Location = New System.Drawing.Point(222, 17)
        Me.txtFecha.MaxLength = 10
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(75, 20)
        Me.txtFecha.TabIndex = 6
        Me.txtFecha.Validator = Administracion.ValidatorType.DateFormat
        '
        'txtOrdenPago
        '
        Me.txtOrdenPago.Cleanable = True
        Me.txtOrdenPago.Empty = True
        Me.txtOrdenPago.EnterIndex = 1
        Me.txtOrdenPago.LabelAssociationKey = 1
        Me.txtOrdenPago.Location = New System.Drawing.Point(105, 17)
        Me.txtOrdenPago.MaxLength = 6
        Me.txtOrdenPago.Name = "txtOrdenPago"
        Me.txtOrdenPago.Size = New System.Drawing.Size(68, 20)
        Me.txtOrdenPago.TabIndex = 5
        Me.txtOrdenPago.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtOrdenPago.Validator = Administracion.ValidatorType.Numeric
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = 5
        Me.CustomLabel5.Location = New System.Drawing.Point(20, 101)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(38, 13)
        Me.CustomLabel5.TabIndex = 4
        Me.CustomLabel5.Text = "Banco"
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 4
        Me.CustomLabel4.Location = New System.Drawing.Point(20, 74)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(78, 13)
        Me.CustomLabel4.TabIndex = 3
        Me.CustomLabel4.Text = "Observaciones"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = 3
        Me.CustomLabel3.Location = New System.Drawing.Point(20, 47)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel3.TabIndex = 2
        Me.CustomLabel3.Text = "Proveedor"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = 2
        Me.CustomLabel2.Location = New System.Drawing.Point(179, 20)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(37, 13)
        Me.CustomLabel2.TabIndex = 1
        Me.CustomLabel2.Text = "Fecha"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = 1
        Me.CustomLabel1.Location = New System.Drawing.Point(20, 20)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(79, 13)
        Me.CustomLabel1.TabIndex = 0
        Me.CustomLabel1.Text = "Orden de Pago"
        '
        'CustomLabel12
        '
        Me.CustomLabel12.AutoSize = True
        Me.CustomLabel12.ControlAssociationKey = 13
        Me.CustomLabel12.Location = New System.Drawing.Point(274, 181)
        Me.CustomLabel12.Name = "CustomLabel12"
        Me.CustomLabel12.Size = New System.Drawing.Size(28, 13)
        Me.CustomLabel12.TabIndex = 67
        Me.CustomLabel12.Text = "Tipo"
        Me.CustomLabel12.Visible = False
        '
        'Tipo
        '
        Me.Tipo.HeaderText = "Tipo"
        Me.Tipo.Name = "Tipo"
        Me.Tipo.ReadOnly = True
        Me.Tipo.Width = 35
        '
        'Letra
        '
        Me.Letra.HeaderText = "Letra"
        Me.Letra.Name = "Letra"
        Me.Letra.ReadOnly = True
        Me.Letra.Width = 40
        '
        'Punto
        '
        Me.Punto.HeaderText = "Punto"
        Me.Punto.Name = "Punto"
        Me.Punto.ReadOnly = True
        Me.Punto.Width = 45
        '
        'Numero
        '
        Me.Numero.HeaderText = "Número"
        Me.Numero.Name = "Numero"
        Me.Numero.ReadOnly = True
        Me.Numero.Width = 70
        '
        'Importe
        '
        Me.Importe.HeaderText = "Importe"
        Me.Importe.Name = "Importe"
        Me.Importe.Width = 75
        '
        'Descripcion
        '
        Me.Descripcion.HeaderText = "Descripción"
        Me.Descripcion.Name = "Descripcion"
        Me.Descripcion.ReadOnly = True
        Me.Descripcion.Width = 115
        '
        'Pagos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(790, 568)
        Me.Controls.Add(Me.CustomLabel13)
        Me.Controls.Add(Me.lblDiferencia)
        Me.Controls.Add(Me.lblFormaPagos)
        Me.Controls.Add(Me.lblPagos)
        Me.Controls.Add(Me.lstSeleccion)
        Me.Controls.Add(Me.btnCarpetas)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.btnCalcular)
        Me.Controls.Add(Me.btnConsulta)
        Me.Controls.Add(Me.btnCerrar)
        Me.Controls.Add(Me.btnLimpiar)
        Me.Controls.Add(Me.btnAgregar)
        Me.Controls.Add(Me.gridFormaPagos)
        Me.Controls.Add(Me.gridPagos)
        Me.Controls.Add(Me.txtConsulta)
        Me.Controls.Add(Me.lstConsulta)
        Me.Controls.Add(Me.txtTotal)
        Me.Controls.Add(Me.CustomLabel11)
        Me.Controls.Add(Me.txtIVA)
        Me.Controls.Add(Me.txtIngresosBrutos)
        Me.Controls.Add(Me.CustomLabel10)
        Me.Controls.Add(Me.txtIBCiudad)
        Me.Controls.Add(Me.CustomLabel9)
        Me.Controls.Add(Me.txtGanancias)
        Me.Controls.Add(Me.CustomLabel8)
        Me.Controls.Add(Me.lblGanancias)
        Me.Controls.Add(Me.cmbTipo)
        Me.Controls.Add(Me.txtParidad)
        Me.Controls.Add(Me.txtFechaParidad)
        Me.Controls.Add(Me.CustomLabel7)
        Me.Controls.Add(Me.CustomLabel6)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.txtNombreBanco)
        Me.Controls.Add(Me.txtBanco)
        Me.Controls.Add(Me.txtObservaciones)
        Me.Controls.Add(Me.txtRazonSocial)
        Me.Controls.Add(Me.txtProveedor)
        Me.Controls.Add(Me.txtFecha)
        Me.Controls.Add(Me.txtOrdenPago)
        Me.Controls.Add(Me.CustomLabel5)
        Me.Controls.Add(Me.CustomLabel4)
        Me.Controls.Add(Me.CustomLabel3)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Controls.Add(Me.CustomLabel12)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Pagos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ingreso de Pagos a Proveedores"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.gridPagos, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.gridFormaPagos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents CustomLabel3 As Administracion.CustomLabel
    Friend WithEvents CustomLabel4 As Administracion.CustomLabel
    Friend WithEvents CustomLabel5 As Administracion.CustomLabel
    Friend WithEvents txtOrdenPago As Administracion.CustomTextBox
    Friend WithEvents txtFecha As Administracion.CustomTextBox
    Friend WithEvents txtRazonSocial As Administracion.CustomTextBox
    Friend WithEvents txtProveedor As Administracion.CustomTextBox
    Friend WithEvents txtObservaciones As Administracion.CustomTextBox
    Friend WithEvents txtBanco As Administracion.CustomTextBox
    Friend WithEvents txtNombreBanco As Administracion.CustomTextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents optTransferencias As System.Windows.Forms.RadioButton
    Friend WithEvents optAnticipos As System.Windows.Forms.RadioButton
    Friend WithEvents optChequeRechazado As System.Windows.Forms.RadioButton
    Friend WithEvents optVarios As System.Windows.Forms.RadioButton
    Friend WithEvents optCtaCte As System.Windows.Forms.RadioButton
    Friend WithEvents CustomLabel6 As Administracion.CustomLabel
    Friend WithEvents CustomLabel7 As Administracion.CustomLabel
    Friend WithEvents txtFechaParidad As Administracion.CustomTextBox
    Friend WithEvents txtParidad As Administracion.CustomTextBox
    Friend WithEvents cmbTipo As Administracion.CustomComboBox
    Friend WithEvents lblGanancias As Administracion.CustomLabel
    Friend WithEvents CustomLabel8 As Administracion.CustomLabel
    Friend WithEvents txtGanancias As Administracion.CustomTextBox
    Friend WithEvents CustomLabel9 As Administracion.CustomLabel
    Friend WithEvents txtIBCiudad As Administracion.CustomTextBox
    Friend WithEvents CustomLabel10 As Administracion.CustomLabel
    Friend WithEvents txtIVA As Administracion.CustomTextBox
    Friend WithEvents txtIngresosBrutos As Administracion.CustomTextBox
    Friend WithEvents CustomLabel11 As Administracion.CustomLabel
    Friend WithEvents txtTotal As Administracion.CustomTextBox
    Friend WithEvents lstConsulta As Administracion.CustomListBox
    Friend WithEvents txtConsulta As Administracion.CustomTextBox
    Friend WithEvents gridPagos As System.Windows.Forms.DataGridView
    Friend WithEvents gridFormaPagos As System.Windows.Forms.DataGridView
    Friend WithEvents btnCerrar As Administracion.CustomButton
    Friend WithEvents btnLimpiar As Administracion.CustomButton
    Friend WithEvents btnAgregar As Administracion.CustomButton
    Friend WithEvents btnConsulta As Administracion.CustomButton
    Friend WithEvents btnCalcular As Administracion.CustomButton
    Friend WithEvents btnImprimir As Administracion.CustomButton
    Friend WithEvents btnCarpetas As Administracion.CustomButton
    Friend WithEvents lstSeleccion As Administracion.CustomListBox
    Friend WithEvents Tipo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Numero2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fecha As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Banco As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nombre As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Importe2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents CustomLabel12 As Administracion.CustomLabel
    Friend WithEvents lblPagos As Administracion.CustomLabel
    Friend WithEvents lblFormaPagos As Administracion.CustomLabel
    Friend WithEvents lblDiferencia As Administracion.CustomLabel
    Friend WithEvents CustomLabel13 As Administracion.CustomLabel
    Friend WithEvents Tipo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Letra As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Punto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Numero As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Importe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Descripcion As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
