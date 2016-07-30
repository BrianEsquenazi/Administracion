<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CierreMes
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
        Me.txtMes = New Administracion.CustomTextBox()
        Me.txtAno = New Administracion.CustomTextBox()
        Me.CustomLabel1 = New Administracion.CustomLabel()
        Me.CustomLabel2 = New Administracion.CustomLabel()
        Me.Proceso = New Administracion.CustomComboBox()
        Me.btnGraba = New Administracion.CustomButton()
        Me.btnMenu = New Administracion.CustomButton()
        Me.SuspendLayout()
        '
        'txtMes
        '
        Me.txtMes.Cleanable = False
        Me.txtMes.Empty = True
        Me.txtMes.EnterIndex = -1
        Me.txtMes.LabelAssociationKey = -1
        Me.txtMes.Location = New System.Drawing.Point(144, 61)
        Me.txtMes.Name = "txtMes"
        Me.txtMes.Size = New System.Drawing.Size(66, 20)
        Me.txtMes.TabIndex = 0
        Me.txtMes.Validator = Administracion.ValidatorType.None
        '
        'txtAno
        '
        Me.txtAno.Cleanable = False
        Me.txtAno.Empty = True
        Me.txtAno.EnterIndex = -1
        Me.txtAno.LabelAssociationKey = -1
        Me.txtAno.Location = New System.Drawing.Point(216, 61)
        Me.txtAno.Name = "txtAno"
        Me.txtAno.Size = New System.Drawing.Size(71, 20)
        Me.txtAno.TabIndex = 1
        Me.txtAno.Validator = Administracion.ValidatorType.None
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Location = New System.Drawing.Point(51, 61)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(57, 13)
        Me.CustomLabel1.TabIndex = 8
        Me.CustomLabel1.Text = "Mes / Ano"
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = -1
        Me.CustomLabel2.Location = New System.Drawing.Point(51, 107)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(40, 13)
        Me.CustomLabel2.TabIndex = 11
        Me.CustomLabel2.Text = "Estado"
        '
        'Proceso
        '
        Me.Proceso.Cleanable = False
        Me.Proceso.Empty = False
        Me.Proceso.EnterIndex = -1
        Me.Proceso.FormattingEnabled = True
        Me.Proceso.LabelAssociationKey = -1
        Me.Proceso.Location = New System.Drawing.Point(144, 104)
        Me.Proceso.Name = "Proceso"
        Me.Proceso.Size = New System.Drawing.Size(151, 21)
        Me.Proceso.TabIndex = 2
        Me.Proceso.Validator = Administracion.ValidatorType.None
        '
        'btnGraba
        '
        Me.btnGraba.Cleanable = False
        Me.btnGraba.EnterIndex = -1
        Me.btnGraba.LabelAssociationKey = -1
        Me.btnGraba.Location = New System.Drawing.Point(83, 167)
        Me.btnGraba.Name = "btnGraba"
        Me.btnGraba.Size = New System.Drawing.Size(96, 34)
        Me.btnGraba.TabIndex = 12
        Me.btnGraba.Text = "Graba"
        Me.btnGraba.UseVisualStyleBackColor = True
        '
        'btnMenu
        '
        Me.btnMenu.Cleanable = False
        Me.btnMenu.EnterIndex = -1
        Me.btnMenu.LabelAssociationKey = -1
        Me.btnMenu.Location = New System.Drawing.Point(232, 167)
        Me.btnMenu.Name = "btnMenu"
        Me.btnMenu.Size = New System.Drawing.Size(89, 34)
        Me.btnMenu.TabIndex = 13
        Me.btnMenu.Text = "Menu Principal"
        Me.btnMenu.UseVisualStyleBackColor = True
        '
        'CierreMes
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(413, 318)
        Me.Controls.Add(Me.btnMenu)
        Me.Controls.Add(Me.btnGraba)
        Me.Controls.Add(Me.Proceso)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.txtAno)
        Me.Controls.Add(Me.txtMes)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Name = "CierreMes"
        Me.Text = "CierreMes"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtMes As Administracion.CustomTextBox
    Friend WithEvents txtAno As Administracion.CustomTextBox
    Friend WithEvents CustomLabel1 As Administracion.CustomLabel
    Friend WithEvents CustomLabel2 As Administracion.CustomLabel
    Friend WithEvents Proceso As Administracion.CustomComboBox
    Friend WithEvents btnGraba As Administracion.CustomButton
    Friend WithEvents btnMenu As Administracion.CustomButton
End Class
