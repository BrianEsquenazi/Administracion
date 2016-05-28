Public Class FormOrganizer

    Private form As Form
    Private height As Integer
    Private width As Integer
    Private controls As List(Of CustomControl)
    Private buttons As List(Of CustomButton)

    Public Sub New(ByVal someForm As Form, ByVal formHeight As Integer, ByVal formWidth As Integer)
        form = someForm
        height = formHeight
        width = formWidth
    End Sub

    Public Sub addControls(ByVal formControls As List(Of CustomControl))
        controls = formControls
        controls.OrderBy(Function(control) control.LabelAssociationKey)
    End Sub

    Public Sub organize()
        form.Height = height
        form.Width = width

        organizeControls()
        'organizeButtons()
    End Sub

    Private Sub organizeButtons()
        createButtons()

        Dim buttonWidth As Integer = (width - 60 - 6 * 3) \ 3 '60 = márgenes, 6 * 3 = separación entre botones
        buttons.ForEach(Sub(button) button.Width = buttonWidth)


    End Sub

    Private Sub createButtons()
        buttons.Add(addButton)
        buttons.Add(deleteButton)
        buttons.Add(cleanButton)
        buttons.Add(queryButton)
        buttons.Add(listButton)
        buttons.Add(closeButton)
    End Sub

    Private Function addButton()
        Dim btn As New CustomButton()
        btn.Parent = form
        btn.Name = "btnAdd"
        btn.Text = "Agregar"
        Return btn
    End Function

    Private Function deleteButton()
        Dim btn As New CustomButton()
        btn.Parent = form
        btn.Name = "btnDelete"
        btn.Text = "Eliminar"
        Return btn
    End Function

    Private Function cleanButton()
        Dim btn As New CustomButton()
        btn.Parent = form
        btn.Name = "btnClean"
        btn.Text = "Limpiar"
        Return btn
    End Function

    Private Function queryButton()
        Dim btn As New CustomButton()
        btn.Parent = form
        btn.Name = "btnQuery"
        btn.Text = "Consulta"
        Return btn
    End Function

    Private Function listButton()
        Dim btn As New CustomButton()
        btn.Parent = form
        btn.Name = "btnList"
        btn.Text = "Listado"
        Return btn
    End Function

    Private Function closeButton()
        Dim btn As New CustomButton()
        btn.Parent = form
        btn.Name = "btnClose"
        btn.Text = "Cerrar"
        Return btn
    End Function

    Private Sub organizeControls()
        Dim top As Integer = 0
        For Each control As CustomControl In controls
            top += 30
            control.setTop(top - 3)
            labelFor(control.LabelAssociationKey).setTop(top)
        Next
    End Sub

    Private Function labelFor(ByVal index As Integer) As CustomLabel
        Return form.Controls.OfType(Of CustomLabel).ToList.Find(Function(label) label.ControlAssociationKey = index)
    End Function
End Class
