<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BancosABM
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
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtDescripcion = New WindowsApplication1.CustomTextBox()
        Me.txtCuenta = New WindowsApplication1.CustomTextBox()
        Me.txtNombre = New WindowsApplication1.CustomTextBox()
        Me.txtCodigo = New WindowsApplication1.CustomTextBox()
        Me.btnClose = New WindowsApplication1.CustomButton()
        Me.btnList = New WindowsApplication1.CustomButton()
        Me.btnQuery = New WindowsApplication1.CustomButton()
        Me.btnClean = New WindowsApplication1.CustomButton()
        Me.btnDelete = New WindowsApplication1.CustomButton()
        Me.btnAdd = New WindowsApplication1.CustomButton()
        Me.txtQuery = New WindowsApplication1.CustomTextBox()
        Me.lstQuery = New WindowsApplication1.CustomListBox()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(31, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(44, 13)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "Nombre"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "Código"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(31, 89)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(86, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Cuenta Contable"
        '
        'txtDescripcion
        '
        Me.txtDescripcion.Cleanable = True
        Me.txtDescripcion.Enabled = False
        Me.txtDescripcion.EnterIndex = -1
        Me.txtDescripcion.LabelAssociationKey = -1
        Me.txtDescripcion.Location = New System.Drawing.Point(221, 86)
        Me.txtDescripcion.MaxLength = 50
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.Size = New System.Drawing.Size(241, 20)
        Me.txtDescripcion.TabIndex = 22
        Me.txtDescripcion.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtCuenta
        '
        Me.txtCuenta.Cleanable = True
        Me.txtCuenta.EnterIndex = 3
        Me.txtCuenta.LabelAssociationKey = -1
        Me.txtCuenta.Location = New System.Drawing.Point(136, 86)
        Me.txtCuenta.MaxLength = 10
        Me.txtCuenta.Name = "txtCuenta"
        Me.txtCuenta.Size = New System.Drawing.Size(79, 20)
        Me.txtCuenta.TabIndex = 21
        Me.txtCuenta.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtNombre
        '
        Me.txtNombre.Cleanable = True
        Me.txtNombre.EnterIndex = 2
        Me.txtNombre.LabelAssociationKey = -1
        Me.txtNombre.Location = New System.Drawing.Point(136, 53)
        Me.txtNombre.MaxLength = 50
        Me.txtNombre.Name = "txtNombre"
        Me.txtNombre.Size = New System.Drawing.Size(326, 20)
        Me.txtNombre.TabIndex = 19
        Me.txtNombre.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtCodigo
        '
        Me.txtCodigo.Cleanable = True
        Me.txtCodigo.EnterIndex = 1
        Me.txtCodigo.LabelAssociationKey = -1
        Me.txtCodigo.Location = New System.Drawing.Point(136, 23)
        Me.txtCodigo.MaxLength = 5
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(79, 20)
        Me.txtCodigo.TabIndex = 18
        Me.txtCodigo.Validator = WindowsApplication1.ValidatorType.None
        '
        'btnClose
        '
        Me.btnClose.Cleanable = False
        Me.btnClose.EnterIndex = -1
        Me.btnClose.LabelAssociationKey = -1
        Me.btnClose.Location = New System.Drawing.Point(308, 164)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(110, 35)
        Me.btnClose.TabIndex = 15
        Me.btnClose.Text = "Cerrar"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnList
        '
        Me.btnList.Cleanable = False
        Me.btnList.EnterIndex = -1
        Me.btnList.LabelAssociationKey = -1
        Me.btnList.Location = New System.Drawing.Point(192, 164)
        Me.btnList.Name = "btnList"
        Me.btnList.Size = New System.Drawing.Size(110, 35)
        Me.btnList.TabIndex = 14
        Me.btnList.Text = "Lista"
        Me.btnList.UseVisualStyleBackColor = True
        '
        'btnQuery
        '
        Me.btnQuery.Cleanable = False
        Me.btnQuery.EnterIndex = -1
        Me.btnQuery.LabelAssociationKey = -1
        Me.btnQuery.Location = New System.Drawing.Point(76, 164)
        Me.btnQuery.Name = "btnQuery"
        Me.btnQuery.Size = New System.Drawing.Size(110, 35)
        Me.btnQuery.TabIndex = 13
        Me.btnQuery.Text = "Consulta"
        Me.btnQuery.UseVisualStyleBackColor = True
        '
        'btnClean
        '
        Me.btnClean.Cleanable = False
        Me.btnClean.EnterIndex = -1
        Me.btnClean.LabelAssociationKey = -1
        Me.btnClean.Location = New System.Drawing.Point(308, 123)
        Me.btnClean.Name = "btnClean"
        Me.btnClean.Size = New System.Drawing.Size(110, 35)
        Me.btnClean.TabIndex = 12
        Me.btnClean.Text = "Limpiar"
        Me.btnClean.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Cleanable = False
        Me.btnDelete.EnterIndex = -1
        Me.btnDelete.LabelAssociationKey = -1
        Me.btnDelete.Location = New System.Drawing.Point(192, 123)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 35)
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "Eliminar"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Cleanable = False
        Me.btnAdd.EnterIndex = -1
        Me.btnAdd.LabelAssociationKey = -1
        Me.btnAdd.Location = New System.Drawing.Point(76, 123)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(110, 35)
        Me.btnAdd.TabIndex = 10
        Me.btnAdd.Text = "Agregar"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'txtQuery
        '
        Me.txtQuery.Cleanable = True
        Me.txtQuery.EnterIndex = -1
        Me.txtQuery.LabelAssociationKey = -1
        Me.txtQuery.Location = New System.Drawing.Point(34, 205)
        Me.txtQuery.MaxLength = 50
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(428, 20)
        Me.txtQuery.TabIndex = 23
        Me.txtQuery.Validator = WindowsApplication1.ValidatorType.None
        '
        'lstQuery
        '
        Me.lstQuery.Cleanable = False
        Me.lstQuery.EnterIndex = -1
        Me.lstQuery.FormattingEnabled = True
        Me.lstQuery.LabelAssociationKey = -1
        Me.lstQuery.Location = New System.Drawing.Point(34, 231)
        Me.lstQuery.Name = "lstQuery"
        Me.lstQuery.Size = New System.Drawing.Size(428, 238)
        Me.lstQuery.TabIndex = 24
        '
        'BancosABM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(474, 391)
        Me.Controls.Add(Me.lstQuery)
        Me.Controls.Add(Me.txtQuery)
        Me.Controls.Add(Me.txtDescripcion)
        Me.Controls.Add(Me.txtCuenta)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.txtNombre)
        Me.Controls.Add(Me.txtCodigo)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnList)
        Me.Controls.Add(Me.btnQuery)
        Me.Controls.Add(Me.btnClean)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnAdd)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "BancosABM"
        Me.Text = "Ingreso de Bancos"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnClose As WindowsApplication1.CustomButton
    Friend WithEvents btnList As WindowsApplication1.CustomButton
    Friend WithEvents btnQuery As WindowsApplication1.CustomButton
    Friend WithEvents btnClean As WindowsApplication1.CustomButton
    Friend WithEvents btnDelete As WindowsApplication1.CustomButton
    Friend WithEvents btnAdd As WindowsApplication1.CustomButton
    Friend WithEvents txtNombre As WindowsApplication1.CustomTextBox
    Friend WithEvents txtCodigo As WindowsApplication1.CustomTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCuenta As WindowsApplication1.CustomTextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtDescripcion As WindowsApplication1.CustomTextBox
    Friend WithEvents txtQuery As WindowsApplication1.CustomTextBox
    Friend WithEvents lstQuery As WindowsApplication1.CustomListBox
End Class
