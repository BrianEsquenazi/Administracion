<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ListadoValoresEnCartera
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
        Me.CustomLabel6 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel5 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel4 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel3 = New WindowsApplication1.CustomLabel()
        Me.txtFecha4 = New System.Windows.Forms.MaskedTextBox()
        Me.txtFecha3 = New System.Windows.Forms.MaskedTextBox()
        Me.txtFecha2 = New System.Windows.Forms.MaskedTextBox()
        Me.txtFecha1 = New System.Windows.Forms.MaskedTextBox()
        Me.txtHasta = New System.Windows.Forms.MaskedTextBox()
        Me.txtDesde = New System.Windows.Forms.MaskedTextBox()
        Me.CustomLabel2 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel1 = New WindowsApplication1.CustomLabel()
        Me.CustomLabel7 = New WindowsApplication1.CustomLabel()
        Me.CustomTextBox1 = New WindowsApplication1.CustomTextBox()
        Me.CustomTextBox2 = New WindowsApplication1.CustomTextBox()
        Me.btnConsulta = New WindowsApplication1.CustomButton()
        Me.txtAyuda = New WindowsApplication1.CustomTextBox()
        Me.btnCancela = New WindowsApplication1.CustomButton()
        Me.btnAcepta = New WindowsApplication1.CustomButton()
        Me.opcImpesora = New System.Windows.Forms.RadioButton()
        Me.opcPantalla = New System.Windows.Forms.RadioButton()
        Me.lstAyuda = New WindowsApplication1.CustomListBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CustomLabel6)
        Me.GroupBox1.Controls.Add(Me.CustomLabel5)
        Me.GroupBox1.Controls.Add(Me.CustomLabel4)
        Me.GroupBox1.Controls.Add(Me.CustomLabel3)
        Me.GroupBox1.Controls.Add(Me.txtFecha4)
        Me.GroupBox1.Controls.Add(Me.txtFecha3)
        Me.GroupBox1.Controls.Add(Me.txtFecha2)
        Me.GroupBox1.Controls.Add(Me.txtFecha1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(445, 114)
        Me.GroupBox1.TabIndex = 33
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "PARAMETROS DE FECHA"
        '
        'CustomLabel6
        '
        Me.CustomLabel6.AutoSize = True
        Me.CustomLabel6.ControlAssociationKey = -1
        Me.CustomLabel6.Location = New System.Drawing.Point(208, 68)
        Me.CustomLabel6.Name = "CustomLabel6"
        Me.CustomLabel6.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel6.TabIndex = 28
        Me.CustomLabel6.Text = "Fecha 4"
        '
        'CustomLabel5
        '
        Me.CustomLabel5.AutoSize = True
        Me.CustomLabel5.ControlAssociationKey = -1
        Me.CustomLabel5.Location = New System.Drawing.Point(208, 33)
        Me.CustomLabel5.Name = "CustomLabel5"
        Me.CustomLabel5.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel5.TabIndex = 27
        Me.CustomLabel5.Text = "Fecha 3"
        '
        'CustomLabel4
        '
        Me.CustomLabel4.AutoSize = True
        Me.CustomLabel4.ControlAssociationKey = -1
        Me.CustomLabel4.Location = New System.Drawing.Point(6, 68)
        Me.CustomLabel4.Name = "CustomLabel4"
        Me.CustomLabel4.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel4.TabIndex = 26
        Me.CustomLabel4.Text = "Fecha 2"
        '
        'CustomLabel3
        '
        Me.CustomLabel3.AutoSize = True
        Me.CustomLabel3.ControlAssociationKey = -1
        Me.CustomLabel3.Location = New System.Drawing.Point(6, 33)
        Me.CustomLabel3.Name = "CustomLabel3"
        Me.CustomLabel3.Size = New System.Drawing.Size(46, 13)
        Me.CustomLabel3.TabIndex = 25
        Me.CustomLabel3.Text = "Fecha 1"
        '
        'txtFecha4
        '
        Me.txtFecha4.Location = New System.Drawing.Point(315, 65)
        Me.txtFecha4.Mask = "##/##/####"
        Me.txtFecha4.Name = "txtFecha4"
        Me.txtFecha4.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtFecha4.Size = New System.Drawing.Size(106, 20)
        Me.txtFecha4.TabIndex = 3
        '
        'txtFecha3
        '
        Me.txtFecha3.Location = New System.Drawing.Point(315, 30)
        Me.txtFecha3.Mask = "##/##/####"
        Me.txtFecha3.Name = "txtFecha3"
        Me.txtFecha3.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtFecha3.Size = New System.Drawing.Size(106, 20)
        Me.txtFecha3.TabIndex = 2
        '
        'txtFecha2
        '
        Me.txtFecha2.Location = New System.Drawing.Point(92, 65)
        Me.txtFecha2.Mask = "##/##/####"
        Me.txtFecha2.Name = "txtFecha2"
        Me.txtFecha2.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtFecha2.Size = New System.Drawing.Size(106, 20)
        Me.txtFecha2.TabIndex = 1
        '
        'txtFecha1
        '
        Me.txtFecha1.Location = New System.Drawing.Point(92, 30)
        Me.txtFecha1.Mask = "##/##/####"
        Me.txtFecha1.Name = "txtFecha1"
        Me.txtFecha1.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtFecha1.Size = New System.Drawing.Size(106, 20)
        Me.txtFecha1.TabIndex = 0
        '
        'txtHasta
        '
        Me.txtHasta.Location = New System.Drawing.Point(327, 137)
        Me.txtHasta.Mask = "##/##/####"
        Me.txtHasta.Name = "txtHasta"
        Me.txtHasta.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtHasta.Size = New System.Drawing.Size(106, 20)
        Me.txtHasta.TabIndex = 5
        '
        'txtDesde
        '
        Me.txtDesde.Location = New System.Drawing.Point(108, 137)
        Me.txtDesde.Mask = "##/##/####"
        Me.txtDesde.Name = "txtDesde"
        Me.txtDesde.PromptChar = Global.Microsoft.VisualBasic.ChrW(32)
        Me.txtDesde.Size = New System.Drawing.Size(106, 20)
        Me.txtDesde.TabIndex = 4
        '
        'CustomLabel2
        '
        Me.CustomLabel2.AutoSize = True
        Me.CustomLabel2.ControlAssociationKey = -1
        Me.CustomLabel2.Location = New System.Drawing.Point(220, 144)
        Me.CustomLabel2.Name = "CustomLabel2"
        Me.CustomLabel2.Size = New System.Drawing.Size(68, 13)
        Me.CustomLabel2.TabIndex = 35
        Me.CustomLabel2.Text = "Hasta Fecha"
        '
        'CustomLabel1
        '
        Me.CustomLabel1.AutoSize = True
        Me.CustomLabel1.ControlAssociationKey = -1
        Me.CustomLabel1.Location = New System.Drawing.Point(18, 144)
        Me.CustomLabel1.Name = "CustomLabel1"
        Me.CustomLabel1.Size = New System.Drawing.Size(71, 13)
        Me.CustomLabel1.TabIndex = 34
        Me.CustomLabel1.Text = "Desde Fecha"
        '
        'CustomLabel7
        '
        Me.CustomLabel7.AutoSize = True
        Me.CustomLabel7.ControlAssociationKey = -1
        Me.CustomLabel7.Location = New System.Drawing.Point(18, 179)
        Me.CustomLabel7.Name = "CustomLabel7"
        Me.CustomLabel7.Size = New System.Drawing.Size(71, 13)
        Me.CustomLabel7.TabIndex = 38
        Me.CustomLabel7.Text = "Desde Fecha"
        '
        'CustomTextBox1
        '
        Me.CustomTextBox1.Cleanable = False
        Me.CustomTextBox1.Empty = True
        Me.CustomTextBox1.EnterIndex = -1
        Me.CustomTextBox1.LabelAssociationKey = -1
        Me.CustomTextBox1.Location = New System.Drawing.Point(104, 176)
        Me.CustomTextBox1.Name = "CustomTextBox1"
        Me.CustomTextBox1.Size = New System.Drawing.Size(100, 20)
        Me.CustomTextBox1.TabIndex = 6
        Me.CustomTextBox1.Validator = WindowsApplication1.ValidatorType.None
        '
        'CustomTextBox2
        '
        Me.CustomTextBox2.Cleanable = False
        Me.CustomTextBox2.Empty = True
        Me.CustomTextBox2.EnterIndex = -1
        Me.CustomTextBox2.LabelAssociationKey = -1
        Me.CustomTextBox2.Location = New System.Drawing.Point(223, 176)
        Me.CustomTextBox2.Name = "CustomTextBox2"
        Me.CustomTextBox2.Size = New System.Drawing.Size(210, 20)
        Me.CustomTextBox2.TabIndex = 40
        Me.CustomTextBox2.Validator = WindowsApplication1.ValidatorType.None
        '
        'btnConsulta
        '
        Me.btnConsulta.Cleanable = False
        Me.btnConsulta.EnterIndex = -1
        Me.btnConsulta.LabelAssociationKey = -1
        Me.btnConsulta.Location = New System.Drawing.Point(324, 255)
        Me.btnConsulta.Name = "btnConsulta"
        Me.btnConsulta.Size = New System.Drawing.Size(120, 40)
        Me.btnConsulta.TabIndex = 46
        Me.btnConsulta.Text = "Consulta"
        Me.btnConsulta.UseVisualStyleBackColor = True
        '
        'txtAyuda
        '
        Me.txtAyuda.Cleanable = False
        Me.txtAyuda.Empty = True
        Me.txtAyuda.EnterIndex = -1
        Me.txtAyuda.LabelAssociationKey = -1
        Me.txtAyuda.Location = New System.Drawing.Point(27, 306)
        Me.txtAyuda.Name = "txtAyuda"
        Me.txtAyuda.Size = New System.Drawing.Size(417, 20)
        Me.txtAyuda.TabIndex = 45
        Me.txtAyuda.Validator = WindowsApplication1.ValidatorType.None
        Me.txtAyuda.Visible = False
        '
        'btnCancela
        '
        Me.btnCancela.Cleanable = False
        Me.btnCancela.EnterIndex = -1
        Me.btnCancela.LabelAssociationKey = -1
        Me.btnCancela.Location = New System.Drawing.Point(182, 255)
        Me.btnCancela.Name = "btnCancela"
        Me.btnCancela.Size = New System.Drawing.Size(120, 40)
        Me.btnCancela.TabIndex = 44
        Me.btnCancela.Text = "Cancela"
        Me.btnCancela.UseVisualStyleBackColor = True
        '
        'btnAcepta
        '
        Me.btnAcepta.Cleanable = False
        Me.btnAcepta.EnterIndex = -1
        Me.btnAcepta.LabelAssociationKey = -1
        Me.btnAcepta.Location = New System.Drawing.Point(31, 254)
        Me.btnAcepta.Name = "btnAcepta"
        Me.btnAcepta.Size = New System.Drawing.Size(120, 41)
        Me.btnAcepta.TabIndex = 43
        Me.btnAcepta.Text = "Acepta"
        Me.btnAcepta.UseVisualStyleBackColor = True
        '
        'opcImpesora
        '
        Me.opcImpesora.AutoSize = True
        Me.opcImpesora.Location = New System.Drawing.Point(253, 217)
        Me.opcImpesora.Name = "opcImpesora"
        Me.opcImpesora.Size = New System.Drawing.Size(71, 17)
        Me.opcImpesora.TabIndex = 42
        Me.opcImpesora.TabStop = True
        Me.opcImpesora.Text = "Impresora"
        Me.opcImpesora.UseVisualStyleBackColor = True
        '
        'opcPantalla
        '
        Me.opcPantalla.AutoSize = True
        Me.opcPantalla.Location = New System.Drawing.Point(119, 217)
        Me.opcPantalla.Name = "opcPantalla"
        Me.opcPantalla.Size = New System.Drawing.Size(63, 17)
        Me.opcPantalla.TabIndex = 41
        Me.opcPantalla.TabStop = True
        Me.opcPantalla.Text = "Pantalla"
        Me.opcPantalla.UseVisualStyleBackColor = True
        '
        'lstAyuda
        '
        Me.lstAyuda.Cleanable = False
        Me.lstAyuda.EnterIndex = -1
        Me.lstAyuda.FormattingEnabled = True
        Me.lstAyuda.LabelAssociationKey = -1
        Me.lstAyuda.Location = New System.Drawing.Point(27, 332)
        Me.lstAyuda.Name = "lstAyuda"
        Me.lstAyuda.Size = New System.Drawing.Size(417, 147)
        Me.lstAyuda.TabIndex = 47
        Me.lstAyuda.Visible = False
        '
        'ListadoValoresEnCartera
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 303)
        Me.Controls.Add(Me.lstAyuda)
        Me.Controls.Add(Me.btnConsulta)
        Me.Controls.Add(Me.txtAyuda)
        Me.Controls.Add(Me.btnCancela)
        Me.Controls.Add(Me.btnAcepta)
        Me.Controls.Add(Me.opcImpesora)
        Me.Controls.Add(Me.opcPantalla)
        Me.Controls.Add(Me.CustomTextBox2)
        Me.Controls.Add(Me.CustomTextBox1)
        Me.Controls.Add(Me.CustomLabel7)
        Me.Controls.Add(Me.txtHasta)
        Me.Controls.Add(Me.txtDesde)
        Me.Controls.Add(Me.CustomLabel2)
        Me.Controls.Add(Me.CustomLabel1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ListadoValoresEnCartera"
        Me.Text = "Listado de Valores En Cartera"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CustomLabel6 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel5 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel4 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel3 As WindowsApplication1.CustomLabel
    Friend WithEvents txtFecha4 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtFecha3 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtFecha2 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtFecha1 As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtHasta As System.Windows.Forms.MaskedTextBox
    Friend WithEvents txtDesde As System.Windows.Forms.MaskedTextBox
    Friend WithEvents CustomLabel2 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel1 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomLabel7 As WindowsApplication1.CustomLabel
    Friend WithEvents CustomTextBox1 As WindowsApplication1.CustomTextBox
    Friend WithEvents CustomTextBox2 As WindowsApplication1.CustomTextBox
    Friend WithEvents btnConsulta As WindowsApplication1.CustomButton
    Friend WithEvents txtAyuda As WindowsApplication1.CustomTextBox
    Friend WithEvents btnCancela As WindowsApplication1.CustomButton
    Friend WithEvents btnAcepta As WindowsApplication1.CustomButton
    Friend WithEvents opcImpesora As System.Windows.Forms.RadioButton
    Friend WithEvents opcPantalla As System.Windows.Forms.RadioButton
    Friend WithEvents lstAyuda As WindowsApplication1.CustomListBox
End Class
