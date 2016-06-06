<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CUFEProveedor
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
        Me.btnAceptar = New WindowsApplication1.CustomButton()
        Me.CustomLabel4 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel5 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel6 = New WindowsApplication1.CustomLabel()
        Me.txtCUFE3Fecha = New WindowsApplication1.CustomTextBox()
        Me.txtCUFE1Fecha = New WindowsApplication1.CustomTextBox()
        Me.txtCUFE2Fecha = New WindowsApplication1.CustomTextBox()
        Me.txtCUFE3 = New WindowsApplication1.CustomTextBox()
        Me.txtCUFE2 = New WindowsApplication1.CustomTextBox()
        Me.txtCUFE1 = New WindowsApplication1.CustomTextBox()
        Me.btnClose = New WindowsApplication1.CustomButton()
        Me.SuspendLayout()
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = 1
        Me.CustomLabel1.Location = New System.Drawing.Point(12, 20)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(41, 13)
        Me.CustomLabel1.TabIndex = 0
        Me.CustomLabel1.Text = "CUFE I"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = 2
        Me.CustomLabel2.Location = New System.Drawing.Point(12, 60)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(44, 13)
        Me.CustomLabel2.TabIndex = 1
        Me.CustomLabel2.Text = "CUFE II"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = 3
        Me.CustomLabel3.Location = New System.Drawing.Point(12, 100)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(47, 13)
        Me.CustomLabel3.TabIndex = 2
        Me.CustomLabel3.Text = "CUFE III"
        '
        'btnAceptar
        '
        Me.btnAceptar.Cleanable = False
        Me.btnAceptar.EnterIndex = -1
        Me.btnAceptar.LabelAssociationKey = -1
        Me.btnAceptar.Location = New System.Drawing.Point(201, 123)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(100, 35)
        Me.btnAceptar.TabIndex = 9
        Me.btnAceptar.Text = "Aceptar"
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = 4
        Me.CustomLabel4.Location = New System.Drawing.Point(17, 178)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(74, 13)
        Me.CustomLabel4.TabIndex = 10
        Me.CustomLabel4.Text = "Fecha CUFE I"
        Me.CustomLabel4.Visible = False
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = 5
        Me.CustomLabel5.Location = New System.Drawing.Point(113, 174)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(77, 13)
        Me.CustomLabel5.TabIndex = 11
        Me.CustomLabel5.Text = "Fecha CUFE II"
        Me.CustomLabel5.Visible = False
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = 6
        Me.CustomLabel6.Location = New System.Drawing.Point(97, 187)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(80, 13)
        Me.CustomLabel6.TabIndex = 12
        Me.CustomLabel6.Text = "Fecha CUFE III"
        Me.CustomLabel6.Visible = False
        '
        'txtCUFE3Fecha
        '
        Me.txtCUFE3Fecha.Cleanable = False
        Me.txtCUFE3Fecha.Empty = True
        Me.txtCUFE3Fecha.EnterIndex = 6
        Me.txtCUFE3Fecha.LabelAssociationKey = 6
        Me.txtCUFE3Fecha.Location = New System.Drawing.Point(228, 97)
        Me.txtCUFE3Fecha.MaxLength = 10
        Me.txtCUFE3Fecha.Name = "txtCUFE3Fecha"
        Me.txtCUFE3Fecha.Size = New System.Drawing.Size(73, 20)
        Me.txtCUFE3Fecha.TabIndex = 8
        Me.txtCUFE3Fecha.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'txtCUFE1Fecha
        '
        Me.txtCUFE1Fecha.Cleanable = False
        Me.txtCUFE1Fecha.Empty = True
        Me.txtCUFE1Fecha.EnterIndex = 2
        Me.txtCUFE1Fecha.LabelAssociationKey = 4
        Me.txtCUFE1Fecha.Location = New System.Drawing.Point(228, 17)
        Me.txtCUFE1Fecha.MaxLength = 10
        Me.txtCUFE1Fecha.Name = "txtCUFE1Fecha"
        Me.txtCUFE1Fecha.Size = New System.Drawing.Size(73, 20)
        Me.txtCUFE1Fecha.TabIndex = 7
        Me.txtCUFE1Fecha.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'txtCUFE2Fecha
        '
        Me.txtCUFE2Fecha.Cleanable = False
        Me.txtCUFE2Fecha.Empty = True
        Me.txtCUFE2Fecha.EnterIndex = 4
        Me.txtCUFE2Fecha.LabelAssociationKey = 5
        Me.txtCUFE2Fecha.Location = New System.Drawing.Point(228, 57)
        Me.txtCUFE2Fecha.MaxLength = 10
        Me.txtCUFE2Fecha.Name = "txtCUFE2Fecha"
        Me.txtCUFE2Fecha.Size = New System.Drawing.Size(73, 20)
        Me.txtCUFE2Fecha.TabIndex = 6
        Me.txtCUFE2Fecha.Validator = WindowsApplication1.ValidatorType.DateFormat
        '
        'txtCUFE3
        '
        Me.txtCUFE3.Cleanable = False
        Me.txtCUFE3.Empty = True
        Me.txtCUFE3.EnterIndex = 5
        Me.txtCUFE3.LabelAssociationKey = 3
        Me.txtCUFE3.Location = New System.Drawing.Point(80, 97)
        Me.txtCUFE3.Name = "txtCUFE3"
        Me.txtCUFE3.Size = New System.Drawing.Size(142, 20)
        Me.txtCUFE3.TabIndex = 5
        Me.txtCUFE3.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtCUFE2
        '
        Me.txtCUFE2.Cleanable = False
        Me.txtCUFE2.Empty = True
        Me.txtCUFE2.EnterIndex = 3
        Me.txtCUFE2.LabelAssociationKey = 2
        Me.txtCUFE2.Location = New System.Drawing.Point(80, 57)
        Me.txtCUFE2.Name = "txtCUFE2"
        Me.txtCUFE2.Size = New System.Drawing.Size(142, 20)
        Me.txtCUFE2.TabIndex = 4
        Me.txtCUFE2.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtCUFE1
        '
        Me.txtCUFE1.Cleanable = False
        Me.txtCUFE1.Empty = True
        Me.txtCUFE1.EnterIndex = 1
        Me.txtCUFE1.LabelAssociationKey = 1
        Me.txtCUFE1.Location = New System.Drawing.Point(80, 17)
        Me.txtCUFE1.Name = "txtCUFE1"
        Me.txtCUFE1.Size = New System.Drawing.Size(142, 20)
        Me.txtCUFE1.TabIndex = 3
        Me.txtCUFE1.Validator = WindowsApplication1.ValidatorType.None
        '
        'btnClose
        '
        Me.btnClose.Cleanable = False
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnClose.EnterIndex = -1
        Me.btnClose.LabelAssociationKey = -1
        Me.btnClose.Location = New System.Drawing.Point(221, 129)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(36, 18)
        Me.btnClose.TabIndex = 13
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'CUFEProveedor
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(329, 166)
        Me.Controls.Add(Me.CustomLabel6)
        Me.Controls.Add(Me.CustomLabel5)
        Me.Controls.Add(Me.CustomLabel4)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.txtCUFE3Fecha)
        Me.Controls.Add(Me.txtCUFE1Fecha)
        Me.Controls.Add(Me.txtCUFE2Fecha)
        Me.Controls.Add(Me.txtCUFE3)
        Me.Controls.Add(Me.txtCUFE2)
        Me.Controls.Add(Me.txtCUFE1)
        Me.Controls.Add(Me.CustomLabel3)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Controls.Add(Me.btnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CUFEProveedor"
        Me.Text = "CUFE Proveedor"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel2 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel3 As WindowsApplication1.CustomLabel
    Friend WithEvents txtCUFE1 As WindowsApplication1.CustomTextBox
    Friend WithEvents txtCUFE2 As WindowsApplication1.CustomTextBox
    Friend WithEvents txtCUFE3 As WindowsApplication1.CustomTextBox
    Friend WithEvents txtCUFE2Fecha As WindowsApplication1.CustomTextBox
    Friend WithEvents txtCUFE1Fecha As WindowsApplication1.CustomTextBox
    Friend WithEvents txtCUFE3Fecha As WindowsApplication1.CustomTextBox
    Friend WithEvents btnAceptar As WindowsApplication1.CustomButton
    Friend WithEvents CustomLabel4 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel5 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel6 As WindowsApplication1.CustomLabel
    Friend WithEvents btnClose As WindowsApplication1.CustomButton
End Class
