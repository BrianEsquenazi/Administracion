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
        Me.CustomLabel1 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.CustomLabel3 = New Administracion.CustomLabel()
        Me.txtFecha = New Administracion.CustomTextBox()
        Me.CustomTextBox1 = New Administracion.CustomTextBox()
        Me.CustomTextBox2 = New Administracion.CustomTextBox()
        Me.CustomTextBox4 = New Administracion.CustomTextBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.txtRetGanancias = New Administracion.CustomTextBox()
        Me.CustomLabel4 = New Administracion.CustomLabel()
        Me.CustomLabel5 = New Administracion.CustomLabel()
        Me.CustomTextBox3 = New Administracion.CustomTextBox()
        Me.CustomTextBox5 = New Administracion.CustomTextBox()
        Me.CustomLabel6 = New Administracion.CustomLabel()
        Me.CustomLabel7 = New Administracion.CustomLabel()
        Me.CustomTextBox6 = New Administracion.CustomTextBox()
        Me.CustomTextBox7 = New Administracion.CustomTextBox()
        Me.CustomLabel8 = New Administracion.CustomLabel()
        Me.CustomLabel9 = New Administracion.CustomLabel()
        Me.CustomTextBox8 = New Administracion.CustomTextBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = 1
        Me.CustomLabel1.Location = New System.Drawing.Point(20, 20)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(64, 13)
        Me.CustomLabel1.TabIndex = 0
        Me.CustomLabel1.Text = "Nro. Recibo"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = 3
        Me.CustomLabel2.Location = New System.Drawing.Point(20, 46)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(64, 13)
        Me.CustomLabel2.TabIndex = 1
        Me.CustomLabel2.Text = "Cod. Cliente"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = 2
        Me.CustomLabel3.Location = New System.Drawing.Point(185, 20)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(37, 13)
        Me.CustomLabel3.TabIndex = 3
        Me.CustomLabel3.Text = "Fecha"
        '
        'txtFecha
        '
        Me.txtFecha.Cleanable = True
        Me.txtFecha.Empty = False
        Me.txtFecha.EnterIndex = 2
        Me.txtFecha.LabelAssociationKey = 2
        Me.txtFecha.Location = New System.Drawing.Point(228, 17)
        Me.txtFecha.MaxLength = 10
        Me.txtFecha.Name = "txtFecha"
        Me.txtFecha.Size = New System.Drawing.Size(75, 20)
        Me.txtFecha.TabIndex = 7
        Me.txtFecha.Validator = Administracion.ValidatorType.DateFormat
        '
        'CustomTextBox1
        '
        Me.CustomTextBox1.Cleanable = True
        Me.CustomTextBox1.Empty = False
        Me.CustomTextBox1.EnterIndex = 1
        Me.CustomTextBox1.LabelAssociationKey = 1
        Me.CustomTextBox1.Location = New System.Drawing.Point(104, 17)
        Me.CustomTextBox1.MaxLength = 10
        Me.CustomTextBox1.Name = "CustomTextBox1"
        Me.CustomTextBox1.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox1.TabIndex = 8
        Me.CustomTextBox1.Validator = Administracion.ValidatorType.DateFormat
        '
        'CustomTextBox2
        '
        Me.CustomTextBox2.Cleanable = True
        Me.CustomTextBox2.Empty = False
        Me.CustomTextBox2.EnterIndex = 3
        Me.CustomTextBox2.LabelAssociationKey = 3
        Me.CustomTextBox2.Location = New System.Drawing.Point(104, 43)
        Me.CustomTextBox2.MaxLength = 10
        Me.CustomTextBox2.Name = "CustomTextBox2"
        Me.CustomTextBox2.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox2.TabIndex = 9
        Me.CustomTextBox2.Validator = Administracion.ValidatorType.DateFormat
        '
        'CustomTextBox4
        '
        Me.CustomTextBox4.Cleanable = True
        Me.CustomTextBox4.Empty = False
        Me.CustomTextBox4.Enabled = False
        Me.CustomTextBox4.EnterIndex = -1
        Me.CustomTextBox4.LabelAssociationKey = 3
        Me.CustomTextBox4.Location = New System.Drawing.Point(185, 43)
        Me.CustomTextBox4.MaxLength = 10
        Me.CustomTextBox4.Name = "CustomTextBox4"
        Me.CustomTextBox4.Size = New System.Drawing.Size(178, 20)
        Me.CustomTextBox4.TabIndex = 12
        Me.CustomTextBox4.Validator = Administracion.ValidatorType.DateFormat
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 174)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(644, 237)
        Me.DataGridView1.TabIndex = 13
        '
        'txtRetGanancias
        '
        Me.txtRetGanancias.Cleanable = False
        Me.txtRetGanancias.Empty = True
        Me.txtRetGanancias.EnterIndex = -1
        Me.txtRetGanancias.LabelAssociationKey = -1
        Me.txtRetGanancias.Location = New System.Drawing.Point(104, 69)
        Me.txtRetGanancias.Name = "txtRetGanancias"
        Me.txtRetGanancias.Size = New System.Drawing.Size(75, 20)
        Me.txtRetGanancias.TabIndex = 14
        Me.txtRetGanancias.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 3
        Me.CustomLabel4.Location = New System.Drawing.Point(20, 72)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(81, 13)
        Me.CustomLabel4.TabIndex = 15
        Me.CustomLabel4.Text = "Ret. Ganancias"
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = 3
        Me.CustomLabel5.Location = New System.Drawing.Point(20, 98)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(47, 13)
        Me.CustomLabel5.TabIndex = 16
        Me.CustomLabel5.Text = "Ret. IVA"
        '
        'CustomTextBox3
        '
        Me.CustomTextBox3.Cleanable = False
        Me.CustomTextBox3.Empty = True
        Me.CustomTextBox3.EnterIndex = -1
        Me.CustomTextBox3.LabelAssociationKey = -1
        Me.CustomTextBox3.Location = New System.Drawing.Point(104, 95)
        Me.CustomTextBox3.Name = "CustomTextBox3"
        Me.CustomTextBox3.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox3.TabIndex = 17
        Me.CustomTextBox3.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomTextBox5
        '
        Me.CustomTextBox5.Cleanable = False
        Me.CustomTextBox5.Empty = True
        Me.CustomTextBox5.EnterIndex = -1
        Me.CustomTextBox5.LabelAssociationKey = -1
        Me.CustomTextBox5.Location = New System.Drawing.Point(237, 95)
        Me.CustomTextBox5.Name = "CustomTextBox5"
        Me.CustomTextBox5.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox5.TabIndex = 21
        Me.CustomTextBox5.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = 3
        Me.CustomLabel6.Location = New System.Drawing.Point(185, 98)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(56, 13)
        Me.CustomLabel6.TabIndex = 20
        Me.CustomLabel6.Text = "Ret. Suss."
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = 3
        Me.CustomLabel7.Location = New System.Drawing.Point(185, 72)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel7.TabIndex = 19
        Me.CustomLabel7.Text = "Ret. I.B."
        '
        'CustomTextBox6
        '
        Me.CustomTextBox6.Cleanable = False
        Me.CustomTextBox6.Empty = True
        Me.CustomTextBox6.EnterIndex = -1
        Me.CustomTextBox6.LabelAssociationKey = -1
        Me.CustomTextBox6.Location = New System.Drawing.Point(237, 69)
        Me.CustomTextBox6.Name = "CustomTextBox6"
        Me.CustomTextBox6.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox6.TabIndex = 18
        Me.CustomTextBox6.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomTextBox7
        '
        Me.CustomTextBox7.Cleanable = False
        Me.CustomTextBox7.Empty = True
        Me.CustomTextBox7.EnterIndex = -1
        Me.CustomTextBox7.LabelAssociationKey = -1
        Me.CustomTextBox7.Location = New System.Drawing.Point(237, 121)
        Me.CustomTextBox7.Name = "CustomTextBox7"
        Me.CustomTextBox7.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox7.TabIndex = 25
        Me.CustomTextBox7.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'CustomLabel8
        '
        Me.CustomLabel8.AutoSize = True
        Me.CustomLabel8.ControlAssociationKey = 3
        Me.CustomLabel8.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomLabel8.Location = New System.Drawing.Point(153, 124)
        Me.CustomLabel8.Name = "CustomLabel8"
        Me.CustomLabel8.Size = New System.Drawing.Size(80, 13)
        Me.CustomLabel8.TabIndex = 24
        Me.CustomLabel8.Text = "Total Recibo"
        '
        'CustomLabel9
        '
        Me.CustomLabel9.AutoSize = True
        Me.CustomLabel9.ControlAssociationKey = 3
        Me.CustomLabel9.Location = New System.Drawing.Point(20, 124)
        Me.CustomLabel9.Name = "CustomLabel9"
        Me.CustomLabel9.Size = New System.Drawing.Size(43, 13)
        Me.CustomLabel9.TabIndex = 23
        Me.CustomLabel9.Text = "Paridad"
        '
        'CustomTextBox8
        '
        Me.CustomTextBox8.Cleanable = False
        Me.CustomTextBox8.Empty = True
        Me.CustomTextBox8.EnterIndex = -1
        Me.CustomTextBox8.LabelAssociationKey = -1
        Me.CustomTextBox8.Location = New System.Drawing.Point(104, 121)
        Me.CustomTextBox8.Name = "CustomTextBox8"
        Me.CustomTextBox8.Size = New System.Drawing.Size(75, 20)
        Me.CustomTextBox8.TabIndex = 22
        Me.CustomTextBox8.Validator = Administracion.ValidatorType.PositiveFloat
        '
        'RecibosProvisorios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(681, 423)
        Me.Controls.Add(Me.CustomTextBox7)
        Me.Controls.Add(Me.CustomLabel8)
        Me.Controls.Add(Me.CustomLabel9)
        Me.Controls.Add(Me.CustomTextBox8)
        Me.Controls.Add(Me.CustomTextBox5)
        Me.Controls.Add(Me.CustomLabel6)
        Me.Controls.Add(Me.CustomLabel7)
        Me.Controls.Add(Me.CustomTextBox6)
        Me.Controls.Add(Me.CustomTextBox3)
        Me.Controls.Add(Me.CustomLabel5)
        Me.Controls.Add(Me.CustomLabel4)
        Me.Controls.Add(Me.txtRetGanancias)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.CustomTextBox4)
        Me.Controls.Add(Me.CustomTextBox2)
        Me.Controls.Add(Me.CustomTextBox1)
        Me.Controls.Add(Me.txtFecha)
        Me.Controls.Add(Me.CustomLabel3)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "RecibosProvisorios"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Recibos Provisorios"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents CustomLabel3 As Administracion.CustomLabel
    Friend WithEvents txtFecha As Administracion.CustomTextBox
    Friend WithEvents CustomTextBox1 As Administracion.CustomTextBox
    Friend WithEvents CustomTextBox2 As Administracion.CustomTextBox
    Friend WithEvents CustomTextBox4 As Administracion.CustomTextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtRetGanancias As Administracion.CustomTextBox
    Friend WithEvents CustomLabel4 As Administracion.CustomLabel
    Friend WithEvents CustomLabel5 As Administracion.CustomLabel
    Friend WithEvents CustomTextBox3 As Administracion.CustomTextBox
    Friend WithEvents CustomTextBox5 As Administracion.CustomTextBox
    Friend WithEvents CustomLabel6 As Administracion.CustomLabel
    Friend WithEvents CustomLabel7 As Administracion.CustomLabel
    Friend WithEvents CustomTextBox6 As Administracion.CustomTextBox
    Friend WithEvents CustomTextBox7 As Administracion.CustomTextBox
    Friend WithEvents CustomLabel8 As Administracion.CustomLabel
    Friend WithEvents CustomLabel9 As Administracion.CustomLabel
    Friend WithEvents CustomTextBox8 As Administracion.CustomTextBox
End Class
