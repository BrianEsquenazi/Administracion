<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CuentaContableABM
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lstQuery = New WindowsApplication1.CustomListBox()
        Me.txtQuery = New WindowsApplication1.CustomTextBox()
        Me.txtDescripcion = New WindowsApplication1.CustomTextBox()
        Me.txtCodigo = New WindowsApplication1.CustomTextBox()
        Me.btnClose = New WindowsApplication1.CustomButton()
        Me.btnList = New WindowsApplication1.CustomButton()
        Me.btnQuery = New WindowsApplication1.CustomButton()
        Me.btnClean = New WindowsApplication1.CustomButton()
        Me.btnDelete = New WindowsApplication1.CustomButton()
        Me.btnAdd = New WindowsApplication1.CustomButton()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(30, 30)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Código"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(30, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Descripción"
        '
        'lstQuery
        '
        Me.lstQuery.Cleanable = False
        Me.lstQuery.EnterIndex = -1
        Me.lstQuery.FormattingEnabled = True
        Me.lstQuery.Location = New System.Drawing.Point(30, 211)
        Me.lstQuery.Name = "lstQuery"
        Me.lstQuery.Size = New System.Drawing.Size(341, 238)
        Me.lstQuery.TabIndex = 10
        '
        'txtQuery
        '
        Me.txtQuery.Cleanable = True
        Me.txtQuery.EnterIndex = 3
        Me.txtQuery.Location = New System.Drawing.Point(30, 182)
        Me.txtQuery.MaxLength = 50
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(342, 20)
        Me.txtQuery.TabIndex = 3
        Me.txtQuery.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtDescripcion
        '
        Me.txtDescripcion.Cleanable = True
        Me.txtDescripcion.EnterIndex = 2
        Me.txtDescripcion.Location = New System.Drawing.Point(110, 57)
        Me.txtDescripcion.MaxLength = 50
        Me.txtDescripcion.Name = "txtDescripcion"
        Me.txtDescripcion.Size = New System.Drawing.Size(265, 20)
        Me.txtDescripcion.TabIndex = 2
        Me.txtDescripcion.Validator = WindowsApplication1.ValidatorType.None
        '
        'txtCodigo
        '
        Me.txtCodigo.Cleanable = True
        Me.txtCodigo.EnterIndex = 1
        Me.txtCodigo.Location = New System.Drawing.Point(110, 27)
        Me.txtCodigo.MaxLength = 10
        Me.txtCodigo.Name = "txtCodigo"
        Me.txtCodigo.Size = New System.Drawing.Size(79, 20)
        Me.txtCodigo.TabIndex = 1
        Me.txtCodigo.Validator = WindowsApplication1.ValidatorType.None
        '
        'btnClose
        '
        Me.btnClose.Cleanable = False
        Me.btnClose.EnterIndex = -1
        Me.btnClose.Location = New System.Drawing.Point(262, 141)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(110, 35)
        Me.btnClose.TabIndex = 9
        Me.btnClose.Text = "Cerrar"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnList
        '
        Me.btnList.Cleanable = False
        Me.btnList.EnterIndex = -1
        Me.btnList.Location = New System.Drawing.Point(146, 141)
        Me.btnList.Name = "btnList"
        Me.btnList.Size = New System.Drawing.Size(110, 35)
        Me.btnList.TabIndex = 8
        Me.btnList.Text = "Lista"
        Me.btnList.UseVisualStyleBackColor = True
        '
        'btnQuery
        '
        Me.btnQuery.Cleanable = False
        Me.btnQuery.EnterIndex = -1
        Me.btnQuery.Location = New System.Drawing.Point(30, 141)
        Me.btnQuery.Name = "btnQuery"
        Me.btnQuery.Size = New System.Drawing.Size(110, 35)
        Me.btnQuery.TabIndex = 7
        Me.btnQuery.Text = "Consulta"
        Me.btnQuery.UseVisualStyleBackColor = True
        '
        'btnClean
        '
        Me.btnClean.Cleanable = False
        Me.btnClean.EnterIndex = -1
        Me.btnClean.Location = New System.Drawing.Point(262, 100)
        Me.btnClean.Name = "btnClean"
        Me.btnClean.Size = New System.Drawing.Size(110, 35)
        Me.btnClean.TabIndex = 6
        Me.btnClean.Text = "Limpiar"
        Me.btnClean.UseVisualStyleBackColor = True
        '
        'btnDelete
        '
        Me.btnDelete.Cleanable = False
        Me.btnDelete.EnterIndex = -1
        Me.btnDelete.Location = New System.Drawing.Point(146, 100)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(110, 35)
        Me.btnDelete.TabIndex = 5
        Me.btnDelete.Text = "Eliminar"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'btnAdd
        '
        Me.btnAdd.Cleanable = False
        Me.btnAdd.EnterIndex = -1
        Me.btnAdd.Location = New System.Drawing.Point(30, 100)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(110, 35)
        Me.btnAdd.TabIndex = 4
        Me.btnAdd.Text = "Agregar"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'CuentaContableABM
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(402, 402)
        Me.Controls.Add(Me.lstQuery)
        Me.Controls.Add(Me.txtQuery)
        Me.Controls.Add(Me.txtDescripcion)
        Me.Controls.Add(Me.txtCodigo)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnList)
        Me.Controls.Add(Me.btnQuery)
        Me.Controls.Add(Me.btnClean)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnAdd)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "CuentaContableABM"
        Me.Text = "CuentaContableABM"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents btnAdd As WindowsApplication1.CustomButton
    Friend WithEvents btnDelete As WindowsApplication1.CustomButton
    Friend WithEvents btnClean As WindowsApplication1.CustomButton
    Friend WithEvents btnQuery As WindowsApplication1.CustomButton
    Friend WithEvents btnList As WindowsApplication1.CustomButton
    Friend WithEvents btnClose As WindowsApplication1.CustomButton
    Friend WithEvents txtCodigo As WindowsApplication1.CustomTextBox
    Friend WithEvents txtDescripcion As WindowsApplication1.CustomTextBox
    Friend WithEvents txtQuery As WindowsApplication1.CustomTextBox
    Friend WithEvents lstQuery As WindowsApplication1.CustomListBox
End Class
