<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConsultaNumeroFactura
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
        Me.txtNombreProveedor = New Administracion.CustomTextBox()
        Me.txtCodigoProveedor = New Administracion.CustomTextBox()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.txtNumero = New Administracion.CustomTextBox()
        Me.txtPunto = New Administracion.CustomTextBox()
        Me.txtLetra = New Administracion.CustomTextBox()
        Me.cmbTipo = New Administracion.CustomComboBox()
        Me.txtTipo = New Administracion.CustomTextBox()
        Me.CustomLabel7 = New Administracion.CustomLabel()
        Me.CustomLabel6 = New Administracion.CustomLabel()
        Me.CustomLabel5 = New Administracion.CustomLabel()
        Me.CustomLabel4 = New Administracion.CustomLabel()
        Me.btnAceptar = New Administracion.CustomButton()
        Me.SuspendLayout()
        '
        'txtNombreProveedor
        '
        Me.txtNombreProveedor.Cleanable = True
        Me.txtNombreProveedor.Empty = False
        Me.txtNombreProveedor.Enabled = False
        Me.txtNombreProveedor.EnterIndex = -1
        Me.txtNombreProveedor.LabelAssociationKey = 2
        Me.txtNombreProveedor.Location = New System.Drawing.Point(159, 12)
        Me.txtNombreProveedor.Name = "txtNombreProveedor"
        Me.txtNombreProveedor.Size = New System.Drawing.Size(241, 20)
        Me.txtNombreProveedor.TabIndex = 31
        Me.txtNombreProveedor.Validator = Administracion.ValidatorType.None
        '
        'txtCodigoProveedor
        '
        Me.txtCodigoProveedor.Cleanable = True
        Me.txtCodigoProveedor.Empty = False
        Me.txtCodigoProveedor.EnterIndex = 2
        Me.txtCodigoProveedor.LabelAssociationKey = 2
        Me.txtCodigoProveedor.Location = New System.Drawing.Point(77, 12)
        Me.txtCodigoProveedor.MaxLength = 11
        Me.txtCodigoProveedor.Name = "txtCodigoProveedor"
        Me.txtCodigoProveedor.Size = New System.Drawing.Size(76, 20)
        Me.txtCodigoProveedor.TabIndex = 30
        Me.txtCodigoProveedor.Validator = Administracion.ValidatorType.None
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = 2
        Me.CustomLabel2.Location = New System.Drawing.Point(15, 15)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel2.TabIndex = 29
        Me.CustomLabel2.Text = "Proveedor"
        '
        'txtNumero
        '
        Me.txtNumero.Cleanable = True
        Me.txtNumero.Empty = False
        Me.txtNumero.EnterIndex = 6
        Me.txtNumero.LabelAssociationKey = 7
        Me.txtNumero.Location = New System.Drawing.Point(417, 40)
        Me.txtNumero.MaxLength = 8
        Me.txtNumero.Name = "txtNumero"
        Me.txtNumero.Size = New System.Drawing.Size(91, 20)
        Me.txtNumero.TabIndex = 42
        Me.txtNumero.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtNumero.Validator = Administracion.ValidatorType.Numeric
        '
        'txtPunto
        '
        Me.txtPunto.Cleanable = True
        Me.txtPunto.Empty = False
        Me.txtPunto.EnterIndex = 5
        Me.txtPunto.LabelAssociationKey = 6
        Me.txtPunto.Location = New System.Drawing.Point(297, 39)
        Me.txtPunto.MaxLength = 4
        Me.txtPunto.Name = "txtPunto"
        Me.txtPunto.Size = New System.Drawing.Size(64, 20)
        Me.txtPunto.TabIndex = 41
        Me.txtPunto.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtPunto.Validator = Administracion.ValidatorType.Numeric
        '
        'txtLetra
        '
        Me.txtLetra.Cleanable = True
        Me.txtLetra.Empty = False
        Me.txtLetra.EnterIndex = 4
        Me.txtLetra.LabelAssociationKey = 5
        Me.txtLetra.Location = New System.Drawing.Point(217, 39)
        Me.txtLetra.MaxLength = 1
        Me.txtLetra.Name = "txtLetra"
        Me.txtLetra.Size = New System.Drawing.Size(33, 20)
        Me.txtLetra.TabIndex = 40
        Me.txtLetra.Validator = Administracion.ValidatorType.None
        '
        'cmbTipo
        '
        Me.cmbTipo.Cleanable = True
        Me.cmbTipo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTipo.Empty = False
        Me.cmbTipo.EnterIndex = 3
        Me.cmbTipo.FormattingEnabled = True
        Me.cmbTipo.Items.AddRange(New Object() {"FC", "ND", "NC"})
        Me.cmbTipo.LabelAssociationKey = 4
        Me.cmbTipo.Location = New System.Drawing.Point(77, 38)
        Me.cmbTipo.Name = "cmbTipo"
        Me.cmbTipo.Size = New System.Drawing.Size(97, 21)
        Me.cmbTipo.TabIndex = 39
        Me.cmbTipo.Validator = Administracion.ValidatorType.None
        '
        'txtTipo
        '
        Me.txtTipo.Cleanable = True
        Me.txtTipo.Empty = False
        Me.txtTipo.EnterIndex = 4
        Me.txtTipo.LabelAssociationKey = 4
        Me.txtTipo.Location = New System.Drawing.Point(77, 39)
        Me.txtTipo.MaxLength = 2
        Me.txtTipo.Name = "txtTipo"
        Me.txtTipo.Size = New System.Drawing.Size(26, 20)
        Me.txtTipo.TabIndex = 38
        Me.txtTipo.Validator = Administracion.ValidatorType.None
        Me.txtTipo.Visible = False
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = 7
        Me.CustomLabel7.Location = New System.Drawing.Point(367, 42)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(44, 13)
        Me.CustomLabel7.TabIndex = 37
        Me.CustomLabel7.Text = "Número"
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = 6
        Me.CustomLabel6.Location = New System.Drawing.Point(256, 42)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(35, 13)
        Me.CustomLabel6.TabIndex = 36
        Me.CustomLabel6.Text = "Punto"
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = 5
        Me.CustomLabel5.Location = New System.Drawing.Point(180, 42)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(31, 13)
        Me.CustomLabel5.TabIndex = 35
        Me.CustomLabel5.Text = "Letra"
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 4
        Me.CustomLabel4.Location = New System.Drawing.Point(15, 41)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(28, 13)
        Me.CustomLabel4.TabIndex = 34
        Me.CustomLabel4.Text = "Tipo"
        '
        'btnAceptar
        '
        Me.btnAceptar.Cleanable = False
        Me.btnAceptar.EnterIndex = -1
        Me.btnAceptar.LabelAssociationKey = -1
        Me.btnAceptar.Location = New System.Drawing.Point(390, 76)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(118, 32)
        Me.btnAceptar.TabIndex = 43
        Me.btnAceptar.Text = "Buscar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'ConsultaNumeroFactura
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(530, 120)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.txtNumero)
        Me.Controls.Add(Me.txtPunto)
        Me.Controls.Add(Me.txtLetra)
        Me.Controls.Add(Me.cmbTipo)
        Me.Controls.Add(Me.txtTipo)
        Me.Controls.Add(Me.CustomLabel7)
        Me.Controls.Add(Me.CustomLabel6)
        Me.Controls.Add(Me.CustomLabel5)
        Me.Controls.Add(Me.CustomLabel4)
        Me.Controls.Add(Me.txtNombreProveedor)
        Me.Controls.Add(Me.txtCodigoProveedor)
        Me.Controls.Add(Me.CustomLabel2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ConsultaNumeroFactura"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ConsultaNumeroFactura"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtNombreProveedor As Administracion.CustomTextBox
    Friend WithEvents txtCodigoProveedor As Administracion.CustomTextBox
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents txtNumero As Administracion.CustomTextBox
    Friend WithEvents txtPunto As Administracion.CustomTextBox
    Friend WithEvents txtLetra As Administracion.CustomTextBox
    Friend WithEvents cmbTipo As Administracion.CustomComboBox
    Friend WithEvents txtTipo As Administracion.CustomTextBox
    Friend WithEvents CustomLabel7 As Administracion.CustomLabel
    Friend WithEvents CustomLabel6 As Administracion.CustomLabel
    Friend WithEvents CustomLabel5 As Administracion.CustomLabel
    Friend WithEvents CustomLabel4 As Administracion.CustomLabel
    Friend WithEvents btnAceptar As Administracion.CustomButton
End Class
