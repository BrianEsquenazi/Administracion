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
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.RadioButton5 = New System.Windows.Forms.RadioButton()
        Me.RadioButton6 = New System.Windows.Forms.RadioButton()
        Me.cmbTipo = New WindowsApplication1.CustomComboBox()
        Me.CustomTextBox7 = New WindowsApplication1.CustomTextBox()
        Me.CustomTextBox6 = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel7 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel6 = New WindowsApplication1.CustomLabel()
        Me.txtNombreBanco = New WindowsApplication1.CustomTextBox()
        Me.txtBanco = New WindowsApplication1.CustomTextBox()
        Me.txtObservaciones = New WindowsApplication1.CustomTextBox()
        Me.txtRazonSocial = New WindowsApplication1.CustomTextBox()
        Me.txtProveedor = New WindowsApplication1.CustomTextBox()
        Me.txtFecha = New WindowsApplication1.CustomTextBox()
        Me.txtOrdenPago = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel5 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel4 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel3 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel2 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel1 = New WindowsApplication1.CustomLabel()
        Me.txtGanancias = New WindowsApplication1.CustomLabel()
        Me.CustomLabel8 = New WindowsApplication1.CustomLabel()
        Me.txtNeto = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel9 = New WindowsApplication1.CustomLabel()
        Me.CustomTextBox1 = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel10 = New WindowsApplication1.CustomLabel()
        Me.CustomTextBox2 = New WindowsApplication1.CustomTextBox()
        Me.CustomTextBox3 = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel11 = New WindowsApplication1.CustomLabel()
        Me.CustomTextBox4 = New WindowsApplication1.CustomTextBox()
        Me.CustomListBox1 = New WindowsApplication1.CustomListBox()
        Me.CustomTextBox5 = New WindowsApplication1.CustomTextBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.Tipo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Letra = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Punto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Numero = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Importe = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Descripcion = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tipo2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Numero2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Banco = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nombre = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Importe2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton6)
        Me.GroupBox1.Controls.Add(Me.RadioButton5)
        Me.GroupBox1.Controls.Add(Me.RadioButton4)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Location = New System.Drawing.Point(20, 129)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(239, 99)
        Me.GroupBox1.TabIndex = 34
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tipo de Orden de Pago"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(6, 23)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(99, 17)
        Me.RadioButton1.TabIndex = 0
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Pagos Cta. Cte."
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(6, 46)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(87, 17)
        Me.RadioButton2.TabIndex = 1
        Me.RadioButton2.Text = "Pagos Varios"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(6, 69)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(104, 17)
        Me.RadioButton3.TabIndex = 2
        Me.RadioButton3.Text = "Ch. Rechazados"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(134, 23)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(68, 17)
        Me.RadioButton4.TabIndex = 3
        Me.RadioButton4.Text = "Anticipos"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'RadioButton5
        '
        Me.RadioButton5.AutoSize = True
        Me.RadioButton5.Location = New System.Drawing.Point(134, 46)
        Me.RadioButton5.Name = "RadioButton5"
        Me.RadioButton5.Size = New System.Drawing.Size(95, 17)
        Me.RadioButton5.TabIndex = 4
        Me.RadioButton5.Text = "Transferencias"
        Me.RadioButton5.UseVisualStyleBackColor = True
        '
        'RadioButton6
        '
        Me.RadioButton6.AutoSize = True
        Me.RadioButton6.Location = New System.Drawing.Point(134, 69)
        Me.RadioButton6.Name = "RadioButton6"
        Me.RadioButton6.Size = New System.Drawing.Size(102, 17)
        Me.RadioButton6.TabIndex = 5
        Me.RadioButton6.Text = "Aplic. Pago Imp."
        Me.RadioButton6.UseVisualStyleBackColor = True
        '
        'cmbTipo
        '
        Me.cmbTipo.Cleanable = False
        Me.cmbTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTipo.Empty = False
        Me.cmbTipo.EnterIndex = -1
        Me.cmbTipo.FormattingEnabled = True
        Me.cmbTipo.Items.AddRange(New Object() {"Normal"})
        Me.cmbTipo.LabelAssociationKey = -1
        Me.cmbTipo.Location = New System.Drawing.Point(268, 177)
        Me.cmbTipo.Name = "cmbTipo"
        Me.cmbTipo.Size = New System.Drawing.Size(160, 21)
        Me.cmbTipo.TabIndex = 39
        Me.cmbTipo.Validator = WindowsApplication1.ValidatorType.None
        '
        'CustomTextBox7
        '
        Me.CustomTextBox7.Cleanable = True
        Me.CustomTextBox7.Empty = True
        Me.CustomTextBox7.EnterIndex = 2
        Me.CustomTextBox7.LabelAssociationKey = -1
        Me.CustomTextBox7.Location = New System.Drawing.Point(347, 151)
        Me.CustomTextBox7.Name = "CustomTextBox7"
        Me.CustomTextBox7.Size = New System.Drawing.Size(81, 20)
        Me.CustomTextBox7.TabIndex = 38
        Me.CustomTextBox7.Validator = WindowsApplication1.ValidatorType.StrictlyPositiveFloat
        '
        'CustomTextBox6
        '
        Me.CustomTextBox6.Cleanable = True
        Me.CustomTextBox6.Empty = True
        Me.CustomTextBox6.EnterIndex = 2
        Me.CustomTextBox6.LabelAssociationKey = -1
        Me.CustomTextBox6.Location = New System.Drawing.Point(347, 126)
        Me.CustomTextBox6.Name = "CustomTextBox6"
        Me.CustomTextBox6.Size = New System.Drawing.Size(81, 20)
        Me.CustomTextBox6.TabIndex = 37
        Me.CustomTextBox6.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = -1
        Me.CustomLabel7.Location = New System.Drawing.Point(265, 156)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(43, 13)
        Me.CustomLabel7.TabIndex = 36
        Me.CustomLabel7.Text = "Paridad"
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = -1
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
        Me.txtNombreBanco.LabelAssociationKey = 3
        Me.txtNombreBanco.Location = New System.Drawing.Point(187, 97)
        Me.txtNombreBanco.Name = "txtNombreBanco"
        Me.txtNombreBanco.Size = New System.Drawing.Size(241, 20)
        Me.txtNombreBanco.TabIndex = 33
        Me.txtNombreBanco.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtBanco
        '
        Me.txtBanco.Cleanable = True
        Me.txtBanco.Empty = True
        Me.txtBanco.EnterIndex = -1
        Me.txtBanco.LabelAssociationKey = -1
        Me.txtBanco.Location = New System.Drawing.Point(105, 98)
        Me.txtBanco.Name = "txtBanco"
        Me.txtBanco.Size = New System.Drawing.Size(76, 20)
        Me.txtBanco.TabIndex = 32
        Me.txtBanco.Validator = WindowsApplication1.ValidatorType.Numeric
        '
        'txtObservaciones
        '
        Me.txtObservaciones.Cleanable = True
        Me.txtObservaciones.Empty = True
        Me.txtObservaciones.EnterIndex = 4
        Me.txtObservaciones.LabelAssociationKey = 4
        Me.txtObservaciones.Location = New System.Drawing.Point(105, 71)
        Me.txtObservaciones.Name = "txtObservaciones"
        Me.txtObservaciones.Size = New System.Drawing.Size(323, 20)
        Me.txtObservaciones.TabIndex = 31
        Me.txtObservaciones.Validator = WindowsApplication1.ValidatorType.None
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
        Me.txtRazonSocial.Validator = WindowsApplication1.ValidatorType.None
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
        Me.txtProveedor.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtFecha
        '
        Me.txtFecha.Cleanable = True
        Me.txtFecha.Empty = True
        Me.txtFecha.EnterIndex = 2
        Me.txtFecha.LabelAssociationKey = -1
        Me.txtFecha.Location = New System.Drawing.Point(222, 17)
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(75, 20)
        Me.txtFecha.TabIndex = 6
        Me.txtFecha.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'txtOrdenPago
        '
        Me.txtOrdenPago.Cleanable = True
        Me.txtOrdenPago.Empty = True
        Me.txtOrdenPago.EnterIndex = 1
        Me.txtOrdenPago.LabelAssociationKey = -1
        Me.txtOrdenPago.Location = New System.Drawing.Point(105, 17)
        Me.txtOrdenPago.MaxLength = 6
        Me.txtOrdenPago.Name = "txtOrdenPago"
        Me.txtOrdenPago.Size = New System.Drawing.Size(68, 20)
        Me.txtOrdenPago.TabIndex = 5
        Me.txtOrdenPago.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtOrdenPago.Validator = WindowsApplication1.ValidatorType.Numeric
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = -1
        Me.CustomLabel5.Location = New System.Drawing.Point(20, 101)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(38, 13)
        Me.CustomLabel5.TabIndex = 4
        Me.CustomLabel5.Text = "Banco"
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = -1
        Me.CustomLabel4.Location = New System.Drawing.Point(20, 74)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(78, 13)
        Me.CustomLabel4.TabIndex = 3
        Me.CustomLabel4.Text = "Observaciones"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = 2
        Me.CustomLabel3.Location = New System.Drawing.Point(20, 47)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel3.TabIndex = 2
        Me.CustomLabel3.Text = "Proveedor"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = -1
        Me.CustomLabel2.Location = New System.Drawing.Point(179, 20)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(37, 13)
        Me.CustomLabel2.TabIndex = 1
        Me.CustomLabel2.Text = "Fecha"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Location = New System.Drawing.Point(20, 20)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(79, 13)
        Me.CustomLabel1.TabIndex = 0
        Me.CustomLabel1.Text = "Orden de Pago"
        '
        'txtGanancias
        '
        Me.txtGanancias.AutoSize = True
        Me.txtGanancias.ControlAssociationKey = -1
        Me.txtGanancias.Location = New System.Drawing.Point(434, 154)
        Me.txtGanancias.Name = "txtGanancias"
        Me.txtGanancias.Size = New System.Drawing.Size(76, 13)
        Me.txtGanancias.TabIndex = 40
        Me.txtGanancias.Text = "Ret. Ganancia"
        '
        'CustomLabel8
        '
        Me.CustomLabel8.AutoSize = True
        Me.CustomLabel8.ControlAssociationKey = -1
        Me.CustomLabel8.Location = New System.Drawing.Point(598, 154)
        Me.CustomLabel8.Name = "CustomLabel8"
        Me.CustomLabel8.Size = New System.Drawing.Size(81, 13)
        Me.CustomLabel8.TabIndex = 41
        Me.CustomLabel8.Text = "Ret. Ing. Brutos"
        '
        'txtNeto
        '
        Me.txtNeto.Cleanable = True
        Me.txtNeto.Empty = False
        Me.txtNeto.EnterIndex = 16
        Me.txtNeto.LabelAssociationKey = 15
        Me.txtNeto.Location = New System.Drawing.Point(516, 151)
        Me.txtNeto.Name = "txtNeto"
        Me.txtNeto.Size = New System.Drawing.Size(75, 20)
        Me.txtNeto.TabIndex = 46
        Me.txtNeto.Validator = WindowsApplication1.ValidatorType.Float
        '
        'CustomLabel9
        '
        Me.CustomLabel9.AutoSize = True
        Me.CustomLabel9.ControlAssociationKey = -1
        Me.CustomLabel9.Location = New System.Drawing.Point(434, 181)
        Me.CustomLabel9.Name = "CustomLabel9"
        Me.CustomLabel9.Size = New System.Drawing.Size(76, 13)
        Me.CustomLabel9.TabIndex = 47
        Me.CustomLabel9.Text = "Ret. IB Ciudad"
        '
        'CustomTextBox1
        '
        Me.CustomTextBox1.Cleanable = True
        Me.CustomTextBox1.Empty = False
        Me.CustomTextBox1.EnterIndex = 16
        Me.CustomTextBox1.LabelAssociationKey = 15
        Me.CustomTextBox1.Location = New System.Drawing.Point(516, 178)
        Me.CustomTextBox1.Name = "CustomTextBox1"
        Me.CustomTextBox1.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox1.TabIndex = 48
        Me.CustomTextBox1.Validator = WindowsApplication1.ValidatorType.Float
        '
        'CustomLabel10
        '
        Me.CustomLabel10.AutoSize = True
        Me.CustomLabel10.ControlAssociationKey = -1
        Me.CustomLabel10.Location = New System.Drawing.Point(597, 181)
        Me.CustomLabel10.Name = "CustomLabel10"
        Me.CustomLabel10.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel10.TabIndex = 49
        Me.CustomLabel10.Text = "Ret. I.V.A."
        '
        'CustomTextBox2
        '
        Me.CustomTextBox2.Cleanable = True
        Me.CustomTextBox2.Empty = False
        Me.CustomTextBox2.EnterIndex = 16
        Me.CustomTextBox2.LabelAssociationKey = 15
        Me.CustomTextBox2.Location = New System.Drawing.Point(685, 178)
        Me.CustomTextBox2.Name = "CustomTextBox2"
        Me.CustomTextBox2.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox2.TabIndex = 51
        Me.CustomTextBox2.Validator = WindowsApplication1.ValidatorType.Float
        '
        'CustomTextBox3
        '
        Me.CustomTextBox3.Cleanable = True
        Me.CustomTextBox3.Empty = False
        Me.CustomTextBox3.EnterIndex = 16
        Me.CustomTextBox3.LabelAssociationKey = 15
        Me.CustomTextBox3.Location = New System.Drawing.Point(685, 151)
        Me.CustomTextBox3.Name = "CustomTextBox3"
        Me.CustomTextBox3.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox3.TabIndex = 50
        Me.CustomTextBox3.Validator = WindowsApplication1.ValidatorType.Float
        '
        'CustomLabel11
        '
        Me.CustomLabel11.AutoSize = True
        Me.CustomLabel11.ControlAssociationKey = -1
        Me.CustomLabel11.Location = New System.Drawing.Point(503, 211)
        Me.CustomLabel11.Name = "CustomLabel11"
        Me.CustomLabel11.Size = New System.Drawing.Size(94, 13)
        Me.CustomLabel11.TabIndex = 52
        Me.CustomLabel11.Text = "Total Retenciones"
        '
        'CustomTextBox4
        '
        Me.CustomTextBox4.Cleanable = True
        Me.CustomTextBox4.Empty = False
        Me.CustomTextBox4.EnterIndex = 16
        Me.CustomTextBox4.LabelAssociationKey = 15
        Me.CustomTextBox4.Location = New System.Drawing.Point(603, 208)
        Me.CustomTextBox4.Name = "CustomTextBox4"
        Me.CustomTextBox4.ReadOnly = True
        Me.CustomTextBox4.Size = New System.Drawing.Size(91, 20)
        Me.CustomTextBox4.TabIndex = 53
        Me.CustomTextBox4.Validator = WindowsApplication1.ValidatorType.Float
        '
        'CustomListBox1
        '
        Me.CustomListBox1.Cleanable = False
        Me.CustomListBox1.EnterIndex = -1
        Me.CustomListBox1.FormattingEnabled = True
        Me.CustomListBox1.LabelAssociationKey = -1
        Me.CustomListBox1.Location = New System.Drawing.Point(437, 38)
        Me.CustomListBox1.Name = "CustomListBox1"
        Me.CustomListBox1.Size = New System.Drawing.Size(333, 108)
        Me.CustomListBox1.TabIndex = 54
        Me.CustomListBox1.Visible = False
        '
        'CustomTextBox5
        '
        Me.CustomTextBox5.Cleanable = False
        Me.CustomTextBox5.Empty = True
        Me.CustomTextBox5.EnterIndex = -1
        Me.CustomTextBox5.LabelAssociationKey = -1
        Me.CustomTextBox5.Location = New System.Drawing.Point(437, 12)
        Me.CustomTextBox5.Name = "CustomTextBox5"
        Me.CustomTextBox5.Size = New System.Drawing.Size(333, 20)
        Me.CustomTextBox5.TabIndex = 55
        Me.CustomTextBox5.Validator = WindowsApplication1.ValidatorType.None
        Me.CustomTextBox5.Visible = False
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tipo, Me.Letra, Me.Punto, Me.Numero, Me.Importe, Me.Descripcion})
        Me.DataGridView1.Location = New System.Drawing.Point(0, 258)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(396, 298)
        Me.DataGridView1.TabIndex = 56
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Tipo2, Me.Numero2, Me.Fecha, Me.Banco, Me.Nombre, Me.Importe2})
        Me.DataGridView2.Location = New System.Drawing.Point(396, 258)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(398, 298)
        Me.DataGridView2.TabIndex = 57
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
        Me.Letra.Width = 35
        '
        'Punto
        '
        Me.Punto.HeaderText = "Punto"
        Me.Punto.Name = "Punto"
        Me.Punto.ReadOnly = True
        Me.Punto.Width = 40
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
        Me.Importe.ReadOnly = True
        Me.Importe.Width = 75
        '
        'Descripcion
        '
        Me.Descripcion.HeaderText = "Descripción"
        Me.Descripcion.Name = "Descripcion"
        Me.Descripcion.ReadOnly = True
        Me.Descripcion.Width = 98
        '
        'Tipo2
        '
        Me.Tipo2.HeaderText = "Tipo"
        Me.Tipo2.Name = "Tipo2"
        Me.Tipo2.Width = 25
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
        Me.Fecha.Width = 65
        '
        'Banco
        '
        Me.Banco.HeaderText = "Banco"
        Me.Banco.Name = "Banco"
        Me.Banco.Width = 35
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
        'Pagos
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(790, 568)
        Me.Controls.Add(Me.DataGridView2)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.CustomTextBox5)
        Me.Controls.Add(Me.CustomListBox1)
        Me.Controls.Add(Me.CustomTextBox4)
        Me.Controls.Add(Me.CustomLabel11)
        Me.Controls.Add(Me.CustomTextBox2)
        Me.Controls.Add(Me.CustomTextBox3)
        Me.Controls.Add(Me.CustomLabel10)
        Me.Controls.Add(Me.CustomTextBox1)
        Me.Controls.Add(Me.CustomLabel9)
        Me.Controls.Add(Me.txtNeto)
        Me.Controls.Add(Me.CustomLabel8)
        Me.Controls.Add(Me.txtGanancias)
        Me.Controls.Add(Me.cmbTipo)
        Me.Controls.Add(Me.CustomTextBox7)
        Me.Controls.Add(Me.CustomTextBox6)
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
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "Pagos"
        Me.Text = "Ingreso de Pagos a Proveedores"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel2 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel3 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel4 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel5 As WindowsApplication1.CustomLabel
    Friend WithEvents txtOrdenPago As WindowsApplication1.CustomTextBox
    Friend WithEvents txtFecha As WindowsApplication1.CustomTextBox
    Friend WithEvents txtRazonSocial As WindowsApplication1.CustomTextBox
    Friend WithEvents txtProveedor As WindowsApplication1.CustomTextBox
    Friend WithEvents txtObservaciones As WindowsApplication1.CustomTextBox
    Friend WithEvents txtBanco As WindowsApplication1.CustomTextBox
    Friend WithEvents txtNombreBanco As WindowsApplication1.CustomTextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton6 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton5 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents CustomLabel6 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel7 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomTextBox6 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomTextBox7 As WindowsApplication1.CustomTextBox
    Friend WithEvents cmbTipo As WindowsApplication1.CustomComboBox
    Friend WithEvents txtGanancias As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel8 As WindowsApplication1.CustomLabel
    Friend WithEvents txtNeto As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomLabel9 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomTextBox1 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomLabel10 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomTextBox2 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomTextBox3 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomLabel11 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomTextBox4 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomListBox1 As WindowsApplication1.CustomListBox
    Friend WithEvents CustomTextBox5 As WindowsApplication1.CustomTextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents Tipo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Letra As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Punto As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Numero As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Importe As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Descripcion As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Tipo2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Numero2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Fecha As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Banco As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Nombre As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Importe2 As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
