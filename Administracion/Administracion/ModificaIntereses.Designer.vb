﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ModificaIntereses
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.gridCtaCte = New System.Windows.Forms.DataGridView()
        Me.btnCancela = New Administracion.CustomButton()
        Me.btnGraba = New Administracion.CustomButton()
        Me.fechaOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.desProveOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.facturaOriginal = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.cuota = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.fecha = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.saldo = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.intereses = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ivaIntereses = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.referencia = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.clave = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.nroInterno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.InteresesControl = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IvaControl = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ReferenciaControl = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.gridCtaCte, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'gridCtaCte
        '
        Me.gridCtaCte.AllowUserToAddRows = False
        Me.gridCtaCte.AllowUserToDeleteRows = False
        Me.gridCtaCte.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.gridCtaCte.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.gridCtaCte.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.fechaOriginal, Me.desProveOriginal, Me.facturaOriginal, Me.cuota, Me.fecha, Me.saldo, Me.intereses, Me.ivaIntereses, Me.referencia, Me.clave, Me.nroInterno, Me.InteresesControl, Me.IvaControl, Me.ReferenciaControl})
        Me.gridCtaCte.Location = New System.Drawing.Point(12, 18)
        Me.gridCtaCte.Name = "gridCtaCte"
        Me.gridCtaCte.RowHeadersVisible = False
        Me.gridCtaCte.Size = New System.Drawing.Size(760, 468)
        Me.gridCtaCte.StandardTab = True
        Me.gridCtaCte.TabIndex = 4
        '
        'btnCancela
        '
        Me.btnCancela.Cleanable = False
        Me.btnCancela.EnterIndex = -1
        Me.btnCancela.LabelAssociationKey = -1
        Me.btnCancela.Location = New System.Drawing.Point(385, 502)
        Me.btnCancela.Name = "btnCancela"
        Me.btnCancela.Size = New System.Drawing.Size(122, 42)
        Me.btnCancela.TabIndex = 6
        Me.btnCancela.Text = "Cancela"
        Me.btnCancela.UseVisualStyleBackColor = True
        '
        'btnGraba
        '
        Me.btnGraba.Cleanable = False
        Me.btnGraba.EnterIndex = -1
        Me.btnGraba.LabelAssociationKey = -1
        Me.btnGraba.Location = New System.Drawing.Point(248, 502)
        Me.btnGraba.Name = "btnGraba"
        Me.btnGraba.Size = New System.Drawing.Size(122, 42)
        Me.btnGraba.TabIndex = 5
        Me.btnGraba.Text = "Graba"
        Me.btnGraba.UseVisualStyleBackColor = True
        '
        'fechaOriginal
        '
        Me.fechaOriginal.HeaderText = "Fecha"
        Me.fechaOriginal.Name = "fechaOriginal"
        Me.fechaOriginal.ReadOnly = True
        Me.fechaOriginal.Width = 62
        '
        'desProveOriginal
        '
        Me.desProveOriginal.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.desProveOriginal.HeaderText = "Razon"
        Me.desProveOriginal.Name = "desProveOriginal"
        Me.desProveOriginal.ReadOnly = True
        '
        'facturaOriginal
        '
        Me.facturaOriginal.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.facturaOriginal.HeaderText = "Factura"
        Me.facturaOriginal.Name = "facturaOriginal"
        Me.facturaOriginal.ReadOnly = True
        Me.facturaOriginal.Width = 68
        '
        'cuota
        '
        Me.cuota.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.cuota.HeaderText = "Cuota"
        Me.cuota.Name = "cuota"
        Me.cuota.ReadOnly = True
        Me.cuota.Width = 60
        '
        'fecha
        '
        Me.fecha.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.fecha.HeaderText = "Vencimiento"
        Me.fecha.Name = "fecha"
        Me.fecha.ReadOnly = True
        Me.fecha.Width = 90
        '
        'saldo
        '
        Me.saldo.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.saldo.DefaultCellStyle = DataGridViewCellStyle1
        Me.saldo.HeaderText = "Saldo"
        Me.saldo.Name = "saldo"
        Me.saldo.ReadOnly = True
        Me.saldo.Width = 59
        '
        'intereses
        '
        Me.intereses.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.intereses.DefaultCellStyle = DataGridViewCellStyle2
        Me.intereses.HeaderText = "Intereses"
        Me.intereses.Name = "intereses"
        Me.intereses.Width = 75
        '
        'ivaIntereses
        '
        Me.ivaIntereses.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        Me.ivaIntereses.DefaultCellStyle = DataGridViewCellStyle3
        Me.ivaIntereses.HeaderText = "Iva Int."
        Me.ivaIntereses.Name = "ivaIntereses"
        Me.ivaIntereses.Width = 65
        '
        'referencia
        '
        Me.referencia.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.referencia.HeaderText = "Referencia"
        Me.referencia.Name = "referencia"
        Me.referencia.Width = 84
        '
        'clave
        '
        Me.clave.HeaderText = "Clave"
        Me.clave.Name = "clave"
        Me.clave.Visible = False
        Me.clave.Width = 59
        '
        'nroInterno
        '
        Me.nroInterno.HeaderText = "N° Interno"
        Me.nroInterno.Name = "nroInterno"
        Me.nroInterno.Visible = False
        Me.nroInterno.Width = 80
        '
        'InteresesControl
        '
        Me.InteresesControl.HeaderText = "InteresesControl"
        Me.InteresesControl.Name = "InteresesControl"
        Me.InteresesControl.ReadOnly = True
        Me.InteresesControl.Visible = False
        Me.InteresesControl.Width = 108
        '
        'IvaControl
        '
        Me.IvaControl.HeaderText = "IvaControl"
        Me.IvaControl.Name = "IvaControl"
        Me.IvaControl.ReadOnly = True
        Me.IvaControl.Visible = False
        Me.IvaControl.Width = 80
        '
        'ReferenciaControl
        '
        Me.ReferenciaControl.HeaderText = "ReferenciaControl"
        Me.ReferenciaControl.Name = "ReferenciaControl"
        Me.ReferenciaControl.ReadOnly = True
        Me.ReferenciaControl.Visible = False
        Me.ReferenciaControl.Width = 117
        '
        'ModificaIntereses
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(784, 562)
        Me.Controls.Add(Me.btnCancela)
        Me.Controls.Add(Me.btnGraba)
        Me.Controls.Add(Me.gridCtaCte)
        Me.Name = "ModificaIntereses"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Modificación de Carga de Intereses"
        CType(Me.gridCtaCte, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnCancela As Administracion.CustomButton
    Friend WithEvents btnGraba As Administracion.CustomButton
    Friend WithEvents gridCtaCte As System.Windows.Forms.DataGridView
    Friend WithEvents fechaOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents desProveOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents facturaOriginal As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents cuota As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents fecha As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents saldo As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents intereses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ivaIntereses As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents referencia As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents clave As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents nroInterno As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents InteresesControl As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IvaControl As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents ReferenciaControl As System.Windows.Forms.DataGridViewTextBoxColumn
End Class
