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
        Me.CustomLabel1 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel2 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel3 = New WindowsApplication1.CustomLabel()
        Me.txtFecha = New WindowsApplication1.CustomTextBox()
        Me.CustomTextBox1 = New WindowsApplication1.CustomTextBox()
        Me.CustomTextBox2 = New WindowsApplication1.CustomTextBox()
        Me.CustomLabel4 = New WindowsApplication1.CustomLabel()
        Me.CustomTextBox3 = New WindowsApplication1.CustomTextBox()
        Me.CustomTextBox4 = New WindowsApplication1.CustomTextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.optAnticipos = New System.Windows.Forms.RadioButton()
        Me.optCtaCte = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.GroupBox1.SuspendLayout()
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
        Me.txtFecha.Validator = WindowsApplication1.ValidatorType.DateFormat
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
        Me.CustomTextBox1.Validator = WindowsApplication1.ValidatorType.DateFormat
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
        Me.CustomTextBox2.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 4
        Me.CustomLabel4.Location = New System.Drawing.Point(20, 72)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(78, 13)
        Me.CustomLabel4.TabIndex = 10
        Me.CustomLabel4.Text = "Observaciones"
        '
        'CustomTextBox3
        '
        Me.CustomTextBox3.Cleanable = True
        Me.CustomTextBox3.Empty = False
        Me.CustomTextBox3.EnterIndex = 4
        Me.CustomTextBox3.LabelAssociationKey = 4
        Me.CustomTextBox3.Location = New System.Drawing.Point(104, 69)
        Me.CustomTextBox3.MaxLength = 10
        Me.CustomTextBox3.Name = "CustomTextBox3"
        Me.CustomTextBox3.Size = New System.Drawing.Size(199, 20)
        Me.CustomTextBox3.TabIndex = 11
        Me.CustomTextBox3.Validator = WindowsApplication1.ValidatorType.DateFormat
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
        Me.CustomTextBox4.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.optAnticipos)
        Me.GroupBox1.Controls.Add(Me.optCtaCte)
        Me.GroupBox1.Location = New System.Drawing.Point(23, 95)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(258, 56)
        Me.GroupBox1.TabIndex = 35
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Tipo de Recibos"
        '
        'optAnticipos
        '
        Me.optAnticipos.AutoSize = True
        Me.optAnticipos.Location = New System.Drawing.Point(109, 23)
        Me.optAnticipos.Name = "optAnticipos"
        Me.optAnticipos.Size = New System.Drawing.Size(68, 17)
        Me.optAnticipos.TabIndex = 3
        Me.optAnticipos.Text = "Anticipos"
        Me.optAnticipos.UseVisualStyleBackColor = True
        '
        'optCtaCte
        '
        Me.optCtaCte.AutoSize = True
        Me.optCtaCte.Checked = True
        Me.optCtaCte.Location = New System.Drawing.Point(6, 23)
        Me.optCtaCte.Name = "optCtaCte"
        Me.optCtaCte.Size = New System.Drawing.Size(97, 17)
        Me.optCtaCte.TabIndex = 0
        Me.optCtaCte.TabStop = True
        Me.optCtaCte.Text = "Cobro Cta. Cte."
        Me.optCtaCte.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(183, 23)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(54, 17)
        Me.RadioButton1.TabIndex = 4
        Me.RadioButton1.Text = "Varios"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RecibosProvisorios
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(790, 568)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.CustomTextBox4)
        Me.Controls.Add(Me.CustomTextBox3)
        Me.Controls.Add(Me.CustomLabel4)
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
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel2 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel3 As WindowsApplication1.CustomLabel
    Friend WithEvents txtFecha As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomTextBox1 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomTextBox2 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomLabel4 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomTextBox3 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomTextBox4 As WindowsApplication1.CustomTextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents optAnticipos As System.Windows.Forms.RadioButton
    Friend WithEvents optCtaCte As System.Windows.Forms.RadioButton
End Class
