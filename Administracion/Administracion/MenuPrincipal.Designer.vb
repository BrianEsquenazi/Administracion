﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MenuPrincipal
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
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MaestrosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IngresosDeCuentasContablesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IngresoDeProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IngresoDeBancosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IngresoDeCambiosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IngresoDeRubrosDeProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EnvioEnEMailAProveedoresToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NovedadesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PruebaToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListadosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcesosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcesoSifreToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ProcesoRetencionesOPToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.btnCambio = New System.Windows.Forms.Button()
        Me.lblCargando = New WindowsApplication1.CustomLabel()
        Me.CargarInteresesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MaestrosToolStripMenuItem, Me.NovedadesToolStripMenuItem, Me.ListadosToolStripMenuItem, Me.ProcesosToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(790, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'MaestrosToolStripMenuItem
        '
        Me.MaestrosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.IngresosDeCuentasContablesToolStripMenuItem, Me.IngresoDeProveedoresToolStripMenuItem, Me.IngresoDeBancosToolStripMenuItem, Me.IngresoDeCambiosToolStripMenuItem, Me.IngresoDeRubrosDeProveedoresToolStripMenuItem, Me.EnvioEnEMailAProveedoresToolStripMenuItem})
        Me.MaestrosToolStripMenuItem.Name = "MaestrosToolStripMenuItem"
        Me.MaestrosToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.MaestrosToolStripMenuItem.Text = "Maestros"
        '
        'IngresosDeCuentasContablesToolStripMenuItem
        '
        Me.IngresosDeCuentasContablesToolStripMenuItem.Name = "IngresosDeCuentasContablesToolStripMenuItem"
        Me.IngresosDeCuentasContablesToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.IngresosDeCuentasContablesToolStripMenuItem.Text = "Ingreso de Cuentas Contables"
        '
        'IngresoDeProveedoresToolStripMenuItem
        '
        Me.IngresoDeProveedoresToolStripMenuItem.Name = "IngresoDeProveedoresToolStripMenuItem"
        Me.IngresoDeProveedoresToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.IngresoDeProveedoresToolStripMenuItem.Text = "Ingreso de Proveedores"
        '
        'IngresoDeBancosToolStripMenuItem
        '
        Me.IngresoDeBancosToolStripMenuItem.Name = "IngresoDeBancosToolStripMenuItem"
        Me.IngresoDeBancosToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.IngresoDeBancosToolStripMenuItem.Text = "Ingreso de Bancos"
        '
        'IngresoDeCambiosToolStripMenuItem
        '
        Me.IngresoDeCambiosToolStripMenuItem.Name = "IngresoDeCambiosToolStripMenuItem"
        Me.IngresoDeCambiosToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.IngresoDeCambiosToolStripMenuItem.Text = "Ingreso de Cambios"
        '
        'IngresoDeRubrosDeProveedoresToolStripMenuItem
        '
        Me.IngresoDeRubrosDeProveedoresToolStripMenuItem.Name = "IngresoDeRubrosDeProveedoresToolStripMenuItem"
        Me.IngresoDeRubrosDeProveedoresToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.IngresoDeRubrosDeProveedoresToolStripMenuItem.Text = "Ingreso de Rubros de Proveedores"
        '
        'EnvioEnEMailAProveedoresToolStripMenuItem
        '
        Me.EnvioEnEMailAProveedoresToolStripMenuItem.Name = "EnvioEnEMailAProveedoresToolStripMenuItem"
        Me.EnvioEnEMailAProveedoresToolStripMenuItem.Size = New System.Drawing.Size(253, 22)
        Me.EnvioEnEMailAProveedoresToolStripMenuItem.Text = "Envio en E-Mail a Proveedores"
        '
        'NovedadesToolStripMenuItem
        '
        Me.NovedadesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PruebaToolStripMenuItem, Me.CargarInteresesToolStripMenuItem})
        Me.NovedadesToolStripMenuItem.Name = "NovedadesToolStripMenuItem"
        Me.NovedadesToolStripMenuItem.Size = New System.Drawing.Size(78, 20)
        Me.NovedadesToolStripMenuItem.Text = "Novedades"
        '
        'PruebaToolStripMenuItem
        '
        Me.PruebaToolStripMenuItem.Name = "PruebaToolStripMenuItem"
        Me.PruebaToolStripMenuItem.Size = New System.Drawing.Size(158, 22)
        Me.PruebaToolStripMenuItem.Text = "Prueba"
        '
        'ListadosToolStripMenuItem
        '
        Me.ListadosToolStripMenuItem.Name = "ListadosToolStripMenuItem"
        Me.ListadosToolStripMenuItem.Size = New System.Drawing.Size(62, 20)
        Me.ListadosToolStripMenuItem.Text = "Listados"
        '
        'ProcesosToolStripMenuItem
        '
        Me.ProcesosToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ProcesoSifreToolStripMenuItem, Me.ProcesoRetencionesOPToolStripMenuItem})
        Me.ProcesosToolStripMenuItem.Name = "ProcesosToolStripMenuItem"
        Me.ProcesosToolStripMenuItem.Size = New System.Drawing.Size(66, 20)
        Me.ProcesosToolStripMenuItem.Text = "Procesos"
        '
        'ProcesoSifreToolStripMenuItem
        '
        Me.ProcesoSifreToolStripMenuItem.Name = "ProcesoSifreToolStripMenuItem"
        Me.ProcesoSifreToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.ProcesoSifreToolStripMenuItem.Text = "Proceso Sifre"
        '
        'ProcesoRetencionesOPToolStripMenuItem
        '
        Me.ProcesoRetencionesOPToolStripMenuItem.Name = "ProcesoRetencionesOPToolStripMenuItem"
        Me.ProcesoRetencionesOPToolStripMenuItem.Size = New System.Drawing.Size(208, 22)
        Me.ProcesoRetencionesOPToolStripMenuItem.Text = "Proceso Retenciones O.P."
        '
        'btnCambio
        '
        Me.btnCambio.Location = New System.Drawing.Point(325, 275)
        Me.btnCambio.Name = "btnCambio"
        Me.btnCambio.Size = New System.Drawing.Size(150, 50)
        Me.btnCambio.TabIndex = 1
        Me.btnCambio.Text = "Cambio de Empresa"
        Me.btnCambio.UseVisualStyleBackColor = True
        '
        'lblCargando
        '
        Me.lblCargando.AutoSize = True
        Me.lblCargando.ControlAssociationKey = -1
        Me.lblCargando.Font = New System.Drawing.Font("Lucida Console", 27.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCargando.Location = New System.Drawing.Point(71, 43)
        Me.lblCargando.Name = "lblCargando"
        Me.lblCargando.Size = New System.Drawing.Size(684, 37)
        Me.lblCargando.TabIndex = 2
        Me.lblCargando.Text = "CARGANDO, POR FAVOR ESPERE..."
        Me.lblCargando.Visible = False
        '
        'CargarInteresesToolStripMenuItem
        '
        Me.CargarInteresesToolStripMenuItem.Name = "CargarInteresesToolStripMenuItem"
        Me.CargarInteresesToolStripMenuItem.Size = New System.Drawing.Size(158, 22)
        Me.CargarInteresesToolStripMenuItem.Text = "Cargar Intereses"
        '
        'MenuPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(790, 568)
        Me.Controls.Add(Me.lblCargando)
        Me.Controls.Add(Me.btnCambio)
        Me.Controls.Add(Me.MenuStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MainMenuStrip = Me.MenuStrip1
        Me.MaximizeBox = False
        Me.Name = "MenuPrincipal"
        Me.Text = "Sistema de Administración"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents MaestrosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IngresosDeCuentasContablesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IngresoDeProveedoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IngresoDeBancosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IngresoDeCambiosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IngresoDeRubrosDeProveedoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EnvioEnEMailAProveedoresToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NovedadesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListadosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProcesosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents btnCambio As System.Windows.Forms.Button
    Friend WithEvents lblCargando As WindowsApplication1.CustomLabel
    Friend WithEvents PruebaToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProcesoSifreToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ProcesoRetencionesOPToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CargarInteresesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
